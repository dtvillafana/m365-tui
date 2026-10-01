//! Teams chat list archive rows and accelerated hot-chat polling.
//!
//! Archive rows are a local view over the server hide state (`viewpoint.isHidden`)
//! plus an optional Stay archived set. Hot polling accelerates recently active
//! chats without letting background message GETs outrun Graph.

use std::collections::{HashMap, HashSet};
use std::path::Path;
use std::time::{Duration, Instant};

use m365_core::models::{Chat, ChatMessage};

pub const HOT_POLL_TICK_SECONDS: u64 = 1;
const TEAMS_HOT_AGE_SECONDS: u64 = 2 * 60;
const TEAMS_WARM_AGE_SECONDS: u64 = 5 * 60;
const TEAMS_COOL_AGE_SECONDS: u64 = 15 * 60;
const TEAMS_BACKGROUND_CONCURRENCY: usize = 2;
const TEAMS_BACKGROUND_PER_CHAT_MIN_SECONDS: u64 = 2;
const TEAMS_BACKGROUND_RECOVERY_SECONDS: u64 = 30;
const TEAMS_BACKGROUND_RECOVERY_STEP_RPS: f64 = 0.5;
const STAY_ARCHIVED_STATE_FILE: &str = "stay-archived.json";
const STAY_ARCHIVED_STATE_VERSION: u64 = 1;
const ARCHIVE_EXPANDED_FILE: &str = "teams-archive-expanded";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChatPollTier {
    Hot,
    Warm,
    Cool,
    Normal,
}

impl ChatPollTier {
    pub fn interval(self) -> Option<Duration> {
        match self {
            Self::Hot => Some(Duration::from_secs(2)),
            Self::Warm => Some(Duration::from_secs(5)),
            Self::Cool => Some(Duration::from_secs(10)),
            Self::Normal => None,
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub struct TeamsPollDiagnostics {
    pub hot: usize,
    pub warm: usize,
    pub cool: usize,
    pub normal: usize,
    pub hot_pending: usize,
    pub runtime_budget_rps: f64,
    pub reserved_chats: usize,
}

pub fn chat_poll_tier_from_age(age: Duration) -> ChatPollTier {
    let seconds = age.as_secs();
    if seconds <= TEAMS_HOT_AGE_SECONDS {
        ChatPollTier::Hot
    } else if seconds <= TEAMS_WARM_AGE_SECONDS {
        ChatPollTier::Warm
    } else if seconds <= TEAMS_COOL_AGE_SECONDS {
        ChatPollTier::Cool
    } else {
        ChatPollTier::Normal
    }
}

pub fn message_id_is_newer(new_id: &str, old_id: &str) -> bool {
    match (new_id.parse::<u64>(), old_id.parse::<u64>()) {
        (Ok(new_id), Ok(old_id)) => new_id > old_id,
        _ => new_id > old_id,
    }
}

pub fn teams_chat_is_hidden(chat: &Chat) -> bool {
    chat.viewpoint
        .as_ref()
        .and_then(|viewpoint| viewpoint.is_hidden)
        .unwrap_or(false)
}

pub fn teams_chat_is_archived(chat: &Chat, stay_archived: &HashSet<String>) -> bool {
    teams_chat_is_hidden(chat) || stay_archived.contains(&chat.id)
}

pub fn teams_active_chat_count(chats: &[Chat], stay_archived: &HashSet<String>) -> usize {
    chats
        .iter()
        .filter(|chat| !teams_chat_is_archived(chat, stay_archived))
        .count()
}

pub fn teams_archived_chat_count(chats: &[Chat], stay_archived: &HashSet<String>) -> usize {
    chats
        .iter()
        .filter(|chat| teams_chat_is_archived(chat, stay_archived))
        .count()
}

pub fn teams_chat_row_count(
    chats: &[Chat],
    stay_archived: &HashSet<String>,
    archive_expanded: bool,
) -> usize {
    teams_active_chat_count(chats, stay_archived)
        + 1
        + if archive_expanded {
            teams_archived_chat_count(chats, stay_archived)
        } else {
            0
        }
}

pub fn teams_chat_index_for_row(
    chats: &[Chat],
    stay_archived: &HashSet<String>,
    archive_expanded: bool,
    row: usize,
) -> Option<usize> {
    let active_count = teams_active_chat_count(chats, stay_archived);
    if row < active_count {
        return chats
            .iter()
            .enumerate()
            .filter(|(_, chat)| !teams_chat_is_archived(chat, stay_archived))
            .nth(row)
            .map(|(index, _)| index);
    }
    if row == active_count || !archive_expanded {
        return None;
    }
    chats
        .iter()
        .enumerate()
        .filter(|(_, chat)| teams_chat_is_archived(chat, stay_archived))
        .nth(row.saturating_sub(active_count + 1))
        .map(|(index, _)| index)
}

pub fn teams_chat_row_for_id(
    chats: &[Chat],
    stay_archived: &HashSet<String>,
    archive_expanded: bool,
    chat_id: &str,
) -> Option<usize> {
    let chat = chats.iter().find(|chat| chat.id == chat_id)?;
    if teams_chat_is_archived(chat, stay_archived) {
        if !archive_expanded {
            return None;
        }
        let hidden_position = chats
            .iter()
            .filter(|candidate| teams_chat_is_archived(candidate, stay_archived))
            .position(|candidate| candidate.id == chat_id)?;
        Some(teams_active_chat_count(chats, stay_archived) + 1 + hidden_position)
    } else {
        chats
            .iter()
            .filter(|candidate| !teams_chat_is_archived(candidate, stay_archived))
            .position(|candidate| candidate.id == chat_id)
    }
}

pub fn teams_stay_archived_rehide_candidates(
    chats: &[Chat],
    stay_archived: &HashSet<String>,
    in_flight: &HashSet<String>,
) -> Vec<String> {
    chats
        .iter()
        .filter(|chat| {
            stay_archived.contains(&chat.id)
                && !teams_chat_is_hidden(chat)
                && !in_flight.contains(&chat.id)
        })
        .map(|chat| chat.id.clone())
        .collect()
}

#[derive(Debug, Clone)]
pub struct TeamsHotChatState {
    pub last_activity: Instant,
    pub last_poll_started: Option<Instant>,
    pub latest_message_id: Option<String>,
    pub latest_message_at: Option<chrono::DateTime<chrono::Utc>>,
    pub hot_pending: bool,
}

pub fn activity_instant_from_graph_time(
    value: Option<chrono::DateTime<chrono::Utc>>,
    now: Instant,
) -> Instant {
    let age_seconds = value
        .map(|value| {
            chrono::Utc::now()
                .signed_duration_since(value)
                .num_seconds()
                .max(0) as u64
        })
        .unwrap_or(TEAMS_COOL_AGE_SECONDS + 1)
        .min(TEAMS_COOL_AGE_SECONDS + 1);
    now.checked_sub(Duration::from_secs(age_seconds))
        .unwrap_or(now)
}

pub fn state_message_is_newer(
    state: &TeamsHotChatState,
    message_id: &str,
    message_at: Option<chrono::DateTime<chrono::Utc>>,
) -> bool {
    if state.latest_message_id.as_deref() == Some(message_id) {
        return false;
    }
    match (message_at.as_ref(), state.latest_message_at.as_ref()) {
        (Some(new), Some(old)) if new > old => true,
        (Some(new), Some(old)) if new < old => false,
        (Some(_), Some(_)) => match state.latest_message_id.as_deref() {
            Some(old_id) => message_id_is_newer(message_id, old_id),
            None => true,
        },
        (Some(_), None) => true,
        (None, Some(_)) => false,
        (None, None) => true,
    }
}

fn message_sort_key(message: &ChatMessage) -> (i64, u64) {
    let at = message
        .created_date_time
        .as_deref()
        .and_then(|value| chrono::DateTime::parse_from_rfc3339(value).ok())
        .map(|value| value.timestamp_millis())
        .unwrap_or(0);
    (at, message.id.parse::<u64>().unwrap_or(0))
}

/// Messages that arrived after `previous_message_id`, oldest first.
///
/// A missing baseline announces only the newest message so a first hot poll
/// cannot replay a page of history.
pub fn hot_poll_new_messages<'a>(
    messages: &'a [ChatMessage],
    previous_message_id: Option<&str>,
    newest_changed: bool,
) -> Vec<&'a ChatMessage> {
    if !newest_changed || messages.is_empty() {
        return Vec::new();
    }
    let mut ordered: Vec<&ChatMessage> = messages.iter().collect();
    ordered.sort_by_key(|message| message_sort_key(message));
    if let Some(previous_message_id) = previous_message_id {
        if let Some(previous_index) = ordered
            .iter()
            .position(|message| message.id == previous_message_id)
        {
            return ordered.into_iter().skip(previous_index + 1).collect();
        }
    }
    ordered.into_iter().rev().take(1).collect()
}

#[derive(Clone)]
struct TeamsBackgroundBudgetState {
    runtime_rps: f64,
    tokens: f64,
    last_refill: Instant,
    last_recovery: Instant,
    throttle_generation: u64,
    reserved_chats: HashSet<String>,
    last_chat_request: HashMap<String, Instant>,
}

struct TeamsChatReservation {
    state: std::sync::Arc<std::sync::Mutex<TeamsBackgroundBudgetState>>,
    chat_id: String,
}

impl Drop for TeamsChatReservation {
    fn drop(&mut self) {
        if let Ok(mut state) = self.state.lock() {
            state.reserved_chats.remove(&self.chat_id);
        }
    }
}

pub struct TeamsBackgroundPermit {
    _chat: TeamsChatReservation,
    _global: tokio::sync::OwnedSemaphorePermit,
}

#[derive(Clone)]
pub struct TeamsBackgroundLimiter {
    configured_rps: f64,
    state: std::sync::Arc<std::sync::Mutex<TeamsBackgroundBudgetState>>,
    permits: std::sync::Arc<tokio::sync::Semaphore>,
}

impl TeamsBackgroundLimiter {
    pub fn new(configured_rps: f64, throttle_generation: u64) -> Self {
        let now = Instant::now();
        Self {
            configured_rps,
            state: std::sync::Arc::new(std::sync::Mutex::new(TeamsBackgroundBudgetState {
                runtime_rps: configured_rps,
                tokens: TEAMS_BACKGROUND_CONCURRENCY as f64,
                last_refill: now,
                last_recovery: now,
                throttle_generation,
                reserved_chats: HashSet::new(),
                last_chat_request: HashMap::new(),
            })),
            permits: std::sync::Arc::new(tokio::sync::Semaphore::new(TEAMS_BACKGROUND_CONCURRENCY)),
        }
    }

    pub fn diagnostics_snapshot(&self) -> (f64, usize) {
        let state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        (state.runtime_rps, state.reserved_chats.len())
    }

    async fn reserve_chat(&self, chat_id: &str) -> TeamsChatReservation {
        loop {
            let reserved = {
                let mut state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                if state.reserved_chats.contains(chat_id) {
                    false
                } else {
                    state.reserved_chats.insert(chat_id.to_string());
                    true
                }
            };
            if reserved {
                return TeamsChatReservation {
                    state: self.state.clone(),
                    chat_id: chat_id.to_string(),
                };
            }
            tokio::time::sleep(Duration::from_millis(25)).await;
        }
    }

    async fn wait_per_chat_interval(&self, chat_id: &str, min_interval: Duration) {
        loop {
            let now = Instant::now();
            let wait = {
                let state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                state.last_chat_request.get(chat_id).and_then(|last| {
                    let elapsed = now.duration_since(*last);
                    (elapsed < min_interval).then(|| min_interval - elapsed)
                })
            };
            match wait {
                Some(wait) => tokio::time::sleep(wait).await,
                None => return,
            }
        }
    }

    fn note_chat_request(&self, chat_id: &str) {
        let mut state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        state
            .last_chat_request
            .insert(chat_id.to_string(), Instant::now());
    }

    fn chat_requested_since(&self, chat_id: &str, since: Instant) -> bool {
        let state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        state
            .last_chat_request
            .get(chat_id)
            .is_some_and(|last| *last > since)
    }

    async fn acquire_inner(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
        supersede_after: Option<Instant>,
    ) -> Option<TeamsBackgroundPermit> {
        let global = self
            .permits
            .clone()
            .acquire_owned()
            .await
            .expect("Teams background limiter semaphore closed");

        loop {
            let now = Instant::now();
            let generation = graph.throttle_generation();
            let wait = {
                let mut state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                if generation != state.throttle_generation {
                    state.throttle_generation = generation;
                    state.runtime_rps =
                        (state.runtime_rps * 0.5).max(m365_core::config::TEAMS_POLL_BUDGET_MIN_RPS);
                    state.tokens = state.tokens.min(1.0);
                    state.last_recovery = now;
                    tracing::warn!(
                        "Graph throttling observed; Teams background budget reduced to {:.2} rps",
                        state.runtime_rps
                    );
                } else if state.runtime_rps < self.configured_rps {
                    let steps = now.duration_since(state.last_recovery).as_secs()
                        / TEAMS_BACKGROUND_RECOVERY_SECONDS;
                    if steps > 0 {
                        state.runtime_rps = (state.runtime_rps
                            + steps as f64 * TEAMS_BACKGROUND_RECOVERY_STEP_RPS)
                            .min(self.configured_rps);
                        state.last_recovery +=
                            Duration::from_secs(steps * TEAMS_BACKGROUND_RECOVERY_SECONDS);
                    }
                }
                let elapsed = now.duration_since(state.last_refill).as_secs_f64();
                let runtime_rps = state.runtime_rps;
                state.tokens =
                    (state.tokens + elapsed * runtime_rps).min(TEAMS_BACKGROUND_CONCURRENCY as f64);
                state.last_refill = now;
                if state.tokens >= 1.0 {
                    state.tokens -= 1.0;
                    None
                } else {
                    let seconds = ((1.0 - state.tokens) / runtime_rps).max(0.01);
                    Some(Duration::from_secs_f64(seconds))
                }
            };
            match wait {
                None => break,
                Some(wait) => tokio::time::sleep(wait).await,
            }
        }

        if supersede_after.is_some_and(|since| self.chat_requested_since(chat_id, since)) {
            return None;
        }

        let min_interval = Duration::from_secs(TEAMS_BACKGROUND_PER_CHAT_MIN_SECONDS);
        let reservation = loop {
            self.wait_per_chat_interval(chat_id, min_interval).await;
            let reservation = self.reserve_chat(chat_id).await;
            let now = Instant::now();
            let remaining = {
                let state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                state.last_chat_request.get(chat_id).and_then(|last| {
                    let elapsed = now.duration_since(*last);
                    (elapsed < min_interval).then(|| min_interval - elapsed)
                })
            };
            if let Some(wait) = remaining {
                drop(reservation);
                tokio::time::sleep(wait).await;
                continue;
            }
            break reservation;
        };

        if supersede_after.is_some_and(|since| self.chat_requested_since(chat_id, since)) {
            drop(reservation);
            return None;
        }
        self.note_chat_request(chat_id);
        Some(TeamsBackgroundPermit {
            _chat: reservation,
            _global: global,
        })
    }

    pub async fn acquire(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
    ) -> TeamsBackgroundPermit {
        self.acquire_inner(graph, chat_id, None)
            .await
            .expect("non-supersedable Teams background request disappeared")
    }

    pub async fn acquire_hot(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
        scheduled_at: Instant,
    ) -> Option<TeamsBackgroundPermit> {
        self.acquire_inner(graph, chat_id, Some(scheduled_at)).await
    }
}

pub fn load_stay_archived(cache_dir: Option<&Path>) -> HashSet<String> {
    let Some(cache_dir) = cache_dir else {
        return HashSet::new();
    };
    let path = cache_dir.join(STAY_ARCHIVED_STATE_FILE);
    let text = match std::fs::read_to_string(&path) {
        Ok(text) => text,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return HashSet::new(),
        Err(error) => {
            tracing::warn!(
                "could not read Stay archived state {}: {error}",
                path.display()
            );
            return HashSet::new();
        }
    };
    let value: serde_json::Value = match serde_json::from_str(&text) {
        Ok(value) => value,
        Err(error) => {
            tracing::warn!(
                "could not parse Stay archived state {}: {error}",
                path.display()
            );
            return HashSet::new();
        }
    };
    if value.get("version").and_then(serde_json::Value::as_u64) != Some(STAY_ARCHIVED_STATE_VERSION)
    {
        tracing::warn!(
            "unsupported Stay archived state version in {}",
            path.display()
        );
        return HashSet::new();
    }
    value
        .get("chats")
        .and_then(serde_json::Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(serde_json::Value::as_str)
        .map(str::trim)
        .filter(|chat_id| !chat_id.is_empty())
        .map(str::to_string)
        .collect()
}

pub fn store_stay_archived(
    cache_dir: &Path,
    stay_archived: &HashSet<String>,
) -> std::io::Result<()> {
    crate::teams_cache::create_private_dir(cache_dir)?;
    let mut chats: Vec<&str> = stay_archived.iter().map(String::as_str).collect();
    chats.sort_unstable();
    let value = serde_json::json!({
        "version": STAY_ARCHIVED_STATE_VERSION,
        "chats": chats,
    });
    let bytes = serde_json::to_vec_pretty(&value).map_err(std::io::Error::other)?;
    crate::teams_cache::write_private_file(&cache_dir.join(STAY_ARCHIVED_STATE_FILE), &bytes)
}

pub fn load_archive_expanded(cache_dir: Option<&Path>) -> bool {
    let Some(cache_dir) = cache_dir else {
        return false;
    };
    std::fs::read_to_string(cache_dir.join(ARCHIVE_EXPANDED_FILE))
        .ok()
        .is_some_and(|text| text.trim().eq_ignore_ascii_case("true"))
}

pub fn store_archive_expanded(cache_dir: &Path, expanded: bool) -> std::io::Result<()> {
    crate::teams_cache::create_private_dir(cache_dir)?;
    let text = if expanded { "true\n" } else { "false\n" };
    crate::teams_cache::write_private_file(
        cache_dir.join(ARCHIVE_EXPANDED_FILE).as_path(),
        text.as_bytes(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn chat(id: &str, hidden: bool) -> Chat {
        serde_json::from_value(serde_json::json!({
            "id": id,
            "chatType": "oneOnOne",
            "viewpoint": { "isHidden": hidden }
        }))
        .unwrap()
    }

    fn message(id: &str, created: &str) -> ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": id,
            "createdDateTime": created,
            "body": { "contentType": "text", "content": "hi" }
        }))
        .unwrap()
    }

    #[test]
    fn archive_drawer_rows_map_active_and_hidden_chats() {
        let chats = vec![chat("a", false), chat("b", true), chat("c", false)];
        let stay = HashSet::new();
        assert_eq!(teams_active_chat_count(&chats, &stay), 2);
        assert_eq!(teams_archived_chat_count(&chats, &stay), 1);
        assert_eq!(teams_chat_row_count(&chats, &stay, false), 3);
        assert_eq!(teams_chat_row_count(&chats, &stay, true), 4);
        assert_eq!(teams_chat_index_for_row(&chats, &stay, true, 0), Some(0));
        assert_eq!(teams_chat_index_for_row(&chats, &stay, true, 1), Some(2));
        assert_eq!(teams_chat_index_for_row(&chats, &stay, true, 2), None);
        assert_eq!(teams_chat_index_for_row(&chats, &stay, true, 3), Some(1));
        assert_eq!(teams_chat_row_for_id(&chats, &stay, false, "b"), None);
        assert_eq!(teams_chat_row_for_id(&chats, &stay, true, "b"), Some(3));
    }

    #[test]
    fn stay_archived_keeps_server_visible_chat_in_archive_rows() {
        let chats = vec![chat("a", false), chat("b", false)];
        let stay = HashSet::from(["b".to_string()]);
        assert!(teams_chat_is_archived(&chats[1], &stay));
        assert_eq!(teams_chat_row_for_id(&chats, &stay, true, "b"), Some(2));
        assert_eq!(
            teams_stay_archived_rehide_candidates(&chats, &stay, &HashSet::new()),
            vec!["b".to_string()]
        );
    }

    #[test]
    fn activity_age_maps_to_expected_poll_tier() {
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(30)),
            ChatPollTier::Hot
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(3 * 60)),
            ChatPollTier::Warm
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(10 * 60)),
            ChatPollTier::Cool
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(20 * 60)),
            ChatPollTier::Normal
        );
    }

    #[test]
    fn hot_poll_notifications_include_every_message_after_baseline() {
        let messages = vec![
            message("3", "2026-09-25T10:00:03Z"),
            message("1", "2026-09-25T10:00:01Z"),
            message("2", "2026-09-25T10:00:02Z"),
        ];
        let arrived = hot_poll_new_messages(&messages, Some("1"), true);
        let ids: Vec<&str> = arrived.iter().map(|message| message.id.as_str()).collect();
        assert_eq!(ids, ["2", "3"]);
    }

    #[test]
    fn hot_poll_notification_missing_baseline_announces_only_newest() {
        let messages = vec![
            message("1", "2026-09-25T10:00:01Z"),
            message("2", "2026-09-25T10:00:02Z"),
        ];
        let arrived = hot_poll_new_messages(&messages, None, true);
        assert_eq!(arrived.len(), 1);
        assert_eq!(arrived[0].id, "2");
    }

    #[test]
    fn stay_archived_state_round_trips() {
        let dir = std::env::temp_dir().join(format!(
            "m365-tui-stay-archived-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        let stay = HashSet::from(["chat-b".to_string(), "chat-a".to_string()]);
        store_stay_archived(&dir, &stay).unwrap();
        store_archive_expanded(&dir, true).unwrap();
        assert_eq!(load_stay_archived(Some(&dir)), stay);
        assert!(load_archive_expanded(Some(&dir)));
        let _ = std::fs::remove_dir_all(dir);
    }
}
