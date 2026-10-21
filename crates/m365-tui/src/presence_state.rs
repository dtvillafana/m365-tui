//! Local state machine for the application presence session published by m365-tui.
//!
//! This module owns only local state and transition decisions. Microsoft Graph
//! I/O stays in `app.rs`; completed writes are reported back here.

use std::time::{Duration, Instant};

pub type PresenceSession = (&'static str, &'static str);

pub const AVAILABLE: PresenceSession = ("Available", "Available");
pub const AWAY: PresenceSession = ("Away", "Away");
const RETRY_AFTER: Duration = Duration::from_secs(10);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SessionTarget {
    Absent,
    Present(PresenceSession),
}

impl SessionTarget {
    pub fn from_option(session: Option<PresenceSession>) -> Self {
        match session {
            Some(session) => Self::Present(session),
            None => Self::Absent,
        }
    }

    pub fn label(self) -> String {
        match self {
            Self::Absent => "not publishing".into(),
            Self::Present((availability, activity)) => {
                format!("{availability} · {activity}")
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresenceMode {
    Disabled,
    Automatic,
    Manual,
}

impl PresenceMode {
    pub fn label(self) -> &'static str {
        match self {
            Self::Disabled => "disabled",
            Self::Automatic => "automatic",
            Self::Manual => "manual",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteReason {
    Sync,
    Renew,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PresenceWrite {
    pub target: SessionTarget,
    pub reason: WriteReason,
}

impl PresenceWrite {
    pub fn label(self) -> String {
        let action = match self.target {
            SessionTarget::Absent => "clear".to_string(),
            SessionTarget::Present(session) => {
                format!("set {}", SessionTarget::Present(session).label())
            }
        };
        match self.reason {
            WriteReason::Sync => action,
            WriteReason::Renew => format!("renew {action}"),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct PresenceStateDiagnostics {
    pub mode: PresenceMode,
    pub desired: SessionTarget,
    pub confirmed: Option<SessionTarget>,
    pub in_flight: Option<PresenceWrite>,
    pub retry_waiting: bool,
}

pub struct PresenceStateManager {
    mode: PresenceMode,
    base: SessionTarget,
    desired: SessionTarget,
    confirmed: Option<SessionTarget>,
    confirmed_at: Option<Instant>,
    in_flight: Option<PresenceWrite>,
    retry_not_before: Option<Instant>,
    last_activity: Instant,
    locked: bool,
    /// Automatic app-session state hidden by a temporary idle Away override.
    idle_saved: Option<SessionTarget>,
    lock_saved: Option<SessionTarget>,
    lock_restore: bool,
}

impl PresenceStateManager {
    pub fn new(lock_restore: bool) -> Self {
        let now = Instant::now();
        Self {
            mode: PresenceMode::Disabled,
            base: SessionTarget::Absent,
            desired: SessionTarget::Absent,
            confirmed: None,
            confirmed_at: None,
            in_flight: None,
            retry_not_before: None,
            last_activity: now,
            locked: false,
            idle_saved: None,
            lock_saved: None,
            lock_restore,
        }
    }

    pub fn start_automatic(&mut self, base: PresenceSession, last_activity: Instant, locked: bool) {
        self.mode = PresenceMode::Automatic;
        self.base = SessionTarget::Present(base);
        self.last_activity = last_activity;
        self.locked = locked;

        // When the app starts while the desktop is already idle, the caller
        // supplies Away as the initial target. Remember the normal active state
        // so the first real activity can restore it.
        self.idle_saved = if !locked && base == AWAY {
            Some(SessionTarget::Present(AVAILABLE))
        } else {
            None
        };

        self.lock_saved = if locked && self.lock_restore {
            Some(self.base)
        } else {
            None
        };
        self.retry_not_before = None;
        self.recompute_desired();
    }

    pub fn set_manual(&mut self, session: Option<PresenceSession>) {
        self.mode = PresenceMode::Manual;
        self.base = SessionTarget::from_option(session);
        self.locked = false;
        self.idle_saved = None;
        self.lock_saved = None;
        self.retry_not_before = None;
        self.recompute_desired();
    }

    pub fn clear_manual(&mut self, resume_automatic: bool, now: Instant) {
        self.locked = false;
        self.idle_saved = None;
        self.lock_saved = None;
        self.retry_not_before = None;
        self.last_activity = now;

        if resume_automatic {
            self.mode = PresenceMode::Automatic;
            self.base = SessionTarget::Present(AVAILABLE);
        } else {
            self.mode = PresenceMode::Disabled;
            self.base = SessionTarget::Absent;
        }
        self.recompute_desired();
    }

    pub fn is_automatic(&self) -> bool {
        self.mode == PresenceMode::Automatic
    }

    pub fn on_activity(&mut self, activity_at: Instant) {
        if self.mode != PresenceMode::Automatic {
            return;
        }

        if activity_at > self.last_activity {
            self.last_activity = activity_at;
        }

        if self.locked {
            self.locked = false;
            self.base = if self.lock_restore {
                self.lock_saved
                    .take()
                    .or_else(|| self.idle_saved.take())
                    .unwrap_or(SessionTarget::Present(AVAILABLE))
            } else {
                self.lock_saved = None;
                self.idle_saved = None;
                SessionTarget::Present(AVAILABLE)
            };
            self.idle_saved = None;
        } else if let Some(saved) = self.idle_saved.take() {
            // Away caused by inactivity is temporary. Restore the exact app
            // session that was active before the idle transition.
            self.base = saved;
        }

        // With no active idle/lock override, ordinary activity only refreshes
        // last_activity. It must not erase a previously restored automatic base.
        self.recompute_desired();
    }

    pub fn on_locked(&mut self) {
        if self.mode != PresenceMode::Automatic {
            return;
        }

        if !self.locked && self.lock_restore {
            // Lock can arrive after the desktop had already become idle. In
            // that case save the state from before idle, not temporary Away.
            self.lock_saved = Some(self.idle_saved.take().unwrap_or(self.base));
        }
        if !self.lock_restore {
            self.lock_saved = None;
            self.idle_saved = None;
        }
        self.locked = true;
        self.recompute_desired();
    }

    pub fn on_idle(&mut self, idle_enabled: bool) {
        if self.mode != PresenceMode::Automatic {
            return;
        }

        if self.locked {
            self.recompute_desired();
            return;
        }

        if idle_enabled {
            if self.idle_saved.is_none() {
                self.idle_saved = Some(self.base);
            }
            self.base = SessionTarget::Present(AWAY);
            self.recompute_desired();
        }
    }

    pub fn on_unknown(&mut self, now: Instant, idle_timeout: Duration) {
        if self.mode != PresenceMode::Automatic || self.locked || idle_timeout.is_zero() {
            return;
        }

        if now.saturating_duration_since(self.last_activity) >= idle_timeout {
            if self.idle_saved.is_none() {
                self.idle_saved = Some(self.base);
            }
            self.base = SessionTarget::Present(AWAY);
            self.recompute_desired();
        }
    }

    pub fn next_write(
        &mut self,
        now: Instant,
        renew_after: Duration,
        force_sync: bool,
    ) -> Option<PresenceWrite> {
        if self.in_flight.is_some() || self.retry_not_before.is_some_and(|deadline| now < deadline)
        {
            return None;
        }

        // A fresh disabled manager has never owned a remote session. Do not
        // emit an unsolicited clear on every tick. Explicit user clear uses
        // force_sync=true, and previously managed sessions have confirmed state.
        if self.mode == PresenceMode::Disabled && self.confirmed.is_none() && !force_sync {
            return None;
        }

        let write = if force_sync || self.confirmed != Some(self.desired) {
            PresenceWrite {
                target: self.desired,
                reason: WriteReason::Sync,
            }
        } else {
            match (self.desired, self.confirmed_at) {
                (SessionTarget::Present(_), Some(at))
                    if now.saturating_duration_since(at) >= renew_after =>
                {
                    PresenceWrite {
                        target: self.desired,
                        reason: WriteReason::Renew,
                    }
                }
                _ => return None,
            }
        };

        self.in_flight = Some(write);
        Some(write)
    }

    pub fn finish_write(&mut self, write: PresenceWrite, success: bool, now: Instant) {
        if self.in_flight != Some(write) {
            return;
        }
        self.in_flight = None;

        if success {
            self.confirmed = Some(write.target);
            self.confirmed_at = match write.target {
                SessionTarget::Present(_) => Some(now),
                SessionTarget::Absent => None,
            };
            self.retry_not_before = None;
        } else if self.desired == write.target {
            self.retry_not_before = now.checked_add(RETRY_AFTER);
        } else {
            // A newer state arrived while the failed write was in flight. Do not
            // delay convergence to that newer target behind the old write's backoff.
            self.retry_not_before = None;
        }
    }

    pub fn diagnostics(&self, now: Instant) -> PresenceStateDiagnostics {
        PresenceStateDiagnostics {
            mode: self.mode,
            desired: self.desired,
            confirmed: self.confirmed,
            in_flight: self.in_flight,
            retry_waiting: self.retry_not_before.is_some_and(|deadline| now < deadline),
        }
    }

    pub fn may_have_remote_session(&self) -> bool {
        matches!(self.desired, SessionTarget::Present(_))
            || matches!(self.confirmed, Some(SessionTarget::Present(_)))
            || self
                .in_flight
                .is_some_and(|write| matches!(write.target, SessionTarget::Present(_)))
    }

    fn recompute_desired(&mut self) {
        let next = if self.mode == PresenceMode::Automatic && self.locked {
            SessionTarget::Present(AWAY)
        } else {
            self.base
        };

        if next != self.desired {
            self.desired = next;
            self.retry_not_before = None;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn automatic_idle_and_activity_round_trip() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);
        assert_eq!(
            state
                .next_write(now, Duration::from_secs(1800), false)
                .unwrap()
                .target,
            SessionTarget::Present(AVAILABLE)
        );
        let first = state.in_flight.unwrap();
        state.finish_write(first, true, now);

        state.on_idle(true);
        assert_eq!(state.diagnostics(now).desired, SessionTarget::Present(AWAY));
        state.on_activity(now + Duration::from_secs(1));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
    }

    #[test]
    fn lock_then_unlock_while_write_is_in_flight_converges_to_latest_target() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);
        let initial = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(initial, true, now);

        state.on_locked();
        let away = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        assert_eq!(away.target, SessionTarget::Present(AWAY));

        state.on_activity(now + Duration::from_secs(1));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
        assert!(state
            .next_write(now, Duration::from_secs(1800), false)
            .is_none());

        state.finish_write(away, true, now + Duration::from_secs(2));
        let available = state
            .next_write(
                now + Duration::from_secs(2),
                Duration::from_secs(1800),
                false,
            )
            .unwrap();
        assert_eq!(available.target, SessionTarget::Present(AVAILABLE));
    }

    #[test]
    fn zero_idle_timeout_still_allows_lock_override() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);
        let initial = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(initial, true, now);

        state.on_idle(false);
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
        state.on_locked();
        assert_eq!(state.diagnostics(now).desired, SessionTarget::Present(AWAY));
    }

    #[test]
    fn manual_mode_ignores_desktop_activity() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.set_manual(Some(("Busy", "InACall")));
        state.on_locked();
        state.on_idle(true);
        state.on_activity(now);
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(("Busy", "InACall"))
        );
    }

    #[test]
    fn clearing_manual_available_resumes_automatic_idle_round_trip() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);

        state.set_manual(Some(AVAILABLE));
        assert_eq!(state.diagnostics(now).mode, PresenceMode::Manual);
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );

        state.clear_manual(true, now + Duration::from_secs(1));
        assert_eq!(state.diagnostics(now).mode, PresenceMode::Automatic);
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );

        state.on_idle(true);
        assert_eq!(state.diagnostics(now).desired, SessionTarget::Present(AWAY));

        state.on_activity(now + Duration::from_secs(2));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
    }

    #[test]
    fn idle_restores_saved_automatic_session() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        let previous = ("Busy", "InACall");
        state.start_automatic(previous, now, false);

        let initial = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(initial, true, now);

        state.on_idle(true);
        assert_eq!(state.diagnostics(now).desired, SessionTarget::Present(AWAY));

        state.on_activity(now + Duration::from_secs(1));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(previous)
        );

        // A later ordinary Active sample must not erase the restored base.
        state.on_activity(now + Duration::from_secs(2));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(previous)
        );

        // A later ordinary Active sample must not erase the restored base.
        state.on_activity(now + Duration::from_secs(2));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(previous)
        );

        // A later ordinary Active sample must not erase the restored base.
        state.on_activity(now + Duration::from_secs(2));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(previous)
        );
    }

    #[test]
    fn startup_idle_restores_available_on_first_activity() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AWAY, now, false);

        state.on_activity(now + Duration::from_secs(1));
        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
    }

    #[test]
    fn lock_after_idle_restores_pre_idle_state() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);

        state.on_idle(true);
        state.on_locked();
        state.on_activity(now + Duration::from_secs(1));

        assert_eq!(
            state.diagnostics(now).desired,
            SessionTarget::Present(AVAILABLE)
        );
    }

    #[test]
    fn failed_write_retries_after_backoff() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);
        let write = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(write, false, now);
        assert!(state
            .next_write(
                now + Duration::from_secs(9),
                Duration::from_secs(1800),
                false
            )
            .is_none());
        assert!(state
            .next_write(
                now + Duration::from_secs(10),
                Duration::from_secs(1800),
                false
            )
            .is_some());
    }

    #[test]
    fn renewal_reasserts_confirmed_present_session() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.start_automatic(AVAILABLE, now, false);
        let first = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(first, true, now);

        assert!(state
            .next_write(
                now + Duration::from_secs(1799),
                Duration::from_secs(1800),
                false
            )
            .is_none());
        let renew = state
            .next_write(
                now + Duration::from_secs(1800),
                Duration::from_secs(1800),
                false,
            )
            .unwrap();
        assert_eq!(renew.reason, WriteReason::Renew);
        assert_eq!(renew.target, SessionTarget::Present(AVAILABLE));
    }

    #[test]
    fn force_sync_supports_preferred_presence_change_with_same_session_target() {
        let now = Instant::now();
        let mut state = PresenceStateManager::new(true);
        state.set_manual(Some(("Away", "Away")));
        let first = state
            .next_write(now, Duration::from_secs(1800), false)
            .unwrap();
        state.finish_write(first, true, now);

        state.set_manual(Some(("Away", "Away")));
        let forced = state
            .next_write(
                now + Duration::from_secs(1),
                Duration::from_secs(1800),
                true,
            )
            .unwrap();
        assert_eq!(forced.reason, WriteReason::Sync);
        assert_eq!(forced.target, SessionTarget::Present(("Away", "Away")));
    }
}
