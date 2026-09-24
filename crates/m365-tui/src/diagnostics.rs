//! Read-only F6 diagnostics and safe export helpers.

use std::collections::HashSet;
use std::path::PathBuf;

use chrono::{DateTime, Local, Utc};
use m365_core::auth::TokenInfo;
use m365_core::work_plan::WorkPlanDiagnostics;

use crate::app::{App, PushState};

#[derive(Debug, Clone)]
pub struct DiagnosticsRemote {
    pub generated_at: DateTime<Utc>,
    pub token: Result<TokenInfo, String>,
    pub work_plan: Result<WorkPlanDiagnostics, String>,
}

#[derive(Debug, Default)]
pub struct DiagnosticsState {
    pub loading: bool,
    pub remote: Option<DiagnosticsRemote>,
}

#[derive(Debug, Clone)]
pub enum PresenceProbe {
    Success {
        availability: String,
        activity: String,
    },
    Omitted,
    Skipped(String),
    Error(String),
}

#[derive(Debug, Clone)]
pub struct ContactDiagnosticsRemote {
    pub generated_at: DateTime<Utc>,
    pub presence_read_all: Result<bool, String>,
    pub batch: PresenceProbe,
    pub direct: PresenceProbe,
    pub teams: m365_core::teams_presence::Diagnostics,
}

#[derive(Debug, Default)]
pub struct ContactDiagnosticsState {
    pub loading: bool,
    pub member_type: String,
    pub account_type: String,
    pub member_guid: bool,
    pub preview_guid: bool,
    pub cached_guid: bool,
    pub member_id_shape: String,
    pub member_user_id_shape: String,
    pub preview_id_shape: String,
    pub cached_id_shape: String,
    pub mri_candidate_source: String,
    pub lookup_address_available: bool,
    pub probe_source: String,
    pub tenant_relation: String,
    pub cross_tenant_candidate: bool,
    pub presence_supported: bool,
    pub remote: Option<ContactDiagnosticsRemote>,
}

fn row(label: &str, value: impl AsRef<str>) -> String {
    format!("  {label:<30} {}", value.as_ref())
}

fn local_timezone_name() -> String {
    if let Ok(value) = std::env::var("TZ") {
        let value = value.trim();
        if !value.is_empty() {
            return value.to_string();
        }
    }

    if let Ok(target) = std::fs::read_link("/etc/localtime") {
        let value = target.to_string_lossy();
        if let Some((_, zone)) = value.split_once("/zoneinfo/") {
            if !zone.trim().is_empty() {
                return zone.to_string();
            }
        }
    }

    chrono::Local::now().offset().to_string()
}

fn yes_no(value: bool, yes: &str, no: &str) -> String {
    if value {
        format!("● {yes}")
    } else {
        format!("○ {no}")
    }
}

fn compact_error(value: &str) -> String {
    let flat = value.split_whitespace().collect::<Vec<_>>().join(" ");
    let mut chars = flat.chars();
    let short: String = chars.by_ref().take(120).collect();
    if chars.next().is_some() {
        format!("{short}…")
    } else {
        short
    }
}

fn permission_lines(app: &App, token: Option<&TokenInfo>) -> Vec<String> {
    const SCOPES: &[&str] = &[
        "User.Read",
        "ProfilePhoto.Read.All",
        "People.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Chat.ReadWrite",
        "ChannelMessage.Send",
        "ChannelMessage.Read.All",
        "Presence.Read.All",
        "Presence.ReadWrite",
        "Team.ReadBasic.All",
        "Files.Read.All",
        "User.ReadBasic.All",
        "User.Read.All",
    ];

    let requested: HashSet<String> = app
        .session
        .config
        .scopes
        .iter()
        .map(|value| value.to_ascii_lowercase())
        .collect();
    let granted: Option<HashSet<String>> = token.map(|token| {
        token
            .scopes
            .iter()
            .map(|value| value.to_ascii_lowercase())
            .collect()
    });

    SCOPES
        .iter()
        .map(|scope| {
            let key = scope.to_ascii_lowercase();
            let requested = requested.contains(&key);
            let granted = granted.as_ref().map(|items| items.contains(&key));
            let state = match (requested, granted) {
                (true, Some(true)) => "● granted".to_string(),
                (true, Some(false)) => "! requested, missing from token".to_string(),
                (false, Some(true)) => "! granted, not requested now".to_string(),
                (false, Some(false)) => "○ not requested".to_string(),
                (true, None) => "? requested, token scopes unavailable".to_string(),
                (false, None) => "○ not requested".to_string(),
            };
            row(scope, state)
        })
        .collect()
}

/// Classify an identity without exposing the identity value itself.
///
/// The labels are deliberately coarse. They are safe for support exports and
/// are enough to decide whether Graph presence (GUID) or Teams UPS (MRI) is the
/// next useful diagnostic path.
pub fn identifier_shape(value: Option<&str>) -> &'static str {
    let Some(value) = value.map(str::trim).filter(|value| !value.is_empty()) else {
        return "missing";
    };

    if m365_core::chats::looks_like_user_guid(value) {
        return "Entra GUID";
    }

    let lower = value.to_ascii_lowercase();
    if lower.starts_with("8:live:") {
        "MRI consumer (8:live:)"
    } else if lower.starts_with("8:orgid:") {
        "MRI orgid (8:orgid:)"
    } else if lower.starts_with("8:teamsvisitor:") || lower.starts_with("teamsvisitor:") {
        "MRI visitor (8:teamsvisitor:)"
    } else if lower.starts_with("28:") {
        "Teams MRI-like (28:)"
    } else if lower.starts_with("29:") {
        "Teams MRI-like (29:)"
    } else if lower.starts_with("gid:") {
        "Teams MRI-like (gid:)"
    } else if value
        .split_once(':')
        .is_some_and(|(prefix, rest)| !prefix.is_empty() && !rest.is_empty())
    {
        "colon-prefixed opaque ID"
    } else {
        "opaque non-GUID"
    }
}

/// True only for user-MRI forms we can plausibly send to Teams presence later.
pub fn is_presence_mri_candidate(value: Option<&str>) -> bool {
    let Some(value) = value.map(str::trim).filter(|value| !value.is_empty()) else {
        return false;
    };
    let lower = value.to_ascii_lowercase();
    lower.starts_with("8:live:")
        || lower.starts_with("8:orgid:")
        || lower.starts_with("8:teamsvisitor:")
}

fn safe_graph_error_text(raw: &str) -> String {
    let status = [
        "400 Bad Request",
        "401 Unauthorized",
        "403 Forbidden",
        "404 Not Found",
        "405 Method Not Allowed",
        "409 Conflict",
        "429 Too Many Requests",
        "500 Internal Server Error",
        "502 Bad Gateway",
        "503 Service Unavailable",
        "504 Gateway Timeout",
    ]
    .into_iter()
    .find(|candidate| raw.contains(candidate));

    let code = ["\"code\":\"", "\"code\": \""].into_iter().find_map(|needle| {
        let start = raw.find(needle)? + needle.len();
        let rest = &raw[start..];
        let end = rest.find('"')?;
        let value = &rest[..end];
        (!value.is_empty()
            && value.len() <= 80
            && value
                .chars()
                .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '_' | '-' | '.')))
        .then_some(value)
    });

    match (status, code) {
        (Some(status), Some(code)) => format!("{status} · {code}"),
        (Some(status), None) => status.to_string(),
        (None, Some(code)) => code.to_string(),
        (None, None) => "Graph request failed".to_string(),
    }
}

pub fn safe_graph_error(error: &anyhow::Error) -> String {
    safe_graph_error_text(&format!("{error:#}"))
}

pub fn presence_probe(presence: &m365_core::models::Presence) -> PresenceProbe {
    PresenceProbe::Success {
        availability: presence.availability.as_deref().unwrap_or("unknown").to_string(),
        activity: presence.activity.as_deref().unwrap_or("unknown").to_string(),
    }
}

fn probe_value(probe: &PresenceProbe) -> String {
    match probe {
        PresenceProbe::Success {
            availability,
            activity,
        } => format!("● {availability} · {activity}"),
        PresenceProbe::Omitted => "! no presence returned".to_string(),
        PresenceProbe::Skipped(reason) => format!("○ {reason}"),
        PresenceProbe::Error(error) => format!("× {error}"),
    }
}

fn teams_step_value(step: &m365_core::teams_presence::DiagnosticStep) -> String {
    match step {
        m365_core::teams_presence::DiagnosticStep::Ok(detail) => format!("● {detail}"),
        m365_core::teams_presence::DiagnosticStep::Skipped(detail) => format!("○ {detail}"),
        m365_core::teams_presence::DiagnosticStep::Error(detail) => format!("× {detail}"),
    }
}

pub fn contact_text(app: &App) -> String {
    let state = &app.contact_diagnostics;
    let mut out = Vec::new();
    out.push("m365-tui contact diagnostics".to_string());

    if let Some(remote) = state.remote.as_ref() {
        out.push(format!(
            "Generated: {}",
            remote
                .generated_at
                .with_timezone(&Local)
                .format("%Y-%m-%d %H:%M:%S %:z")
        ));
    } else if state.loading {
        out.push("Generated: loading…".into());
    }
    out.push(format!("Version: {}", env!("CARGO_PKG_VERSION")));
    out.push(String::new());

    out.push("Contact identity".into());
    out.push(row("Chat type", "oneOnOne"));
    out.push(row(
        "Account type",
        if state.account_type.is_empty() {
            "unknown"
        } else {
            &state.account_type
        },
    ));
    out.push(row(
        "Member type",
        if state.member_type.is_empty() {
            "unknown"
        } else {
            &state.member_type
        },
    ));
    out.push(row(
        "User identifier",
        if state.probe_source.is_empty() || state.probe_source == "none" {
            "! no usable Entra GUID"
        } else {
            "● Entra GUID available"
        },
    ));
    out.push(row(
        "Graph probe ID source",
        if state.probe_source.is_empty() {
            "none"
        } else {
            &state.probe_source
        },
    ));
    out.push(row(
        "UPS MRI candidate",
        if state.mri_candidate_source.is_empty() || state.mri_candidate_source == "none" {
            "○ none found"
        } else {
            "● available"
        },
    ));
    out.push(row(
        "UPS MRI source",
        if state.mri_candidate_source.is_empty() {
            "none"
        } else {
            &state.mri_candidate_source
        },
    ));
    out.push(row(
        "Tenant relation",
        if state.tenant_relation.is_empty() {
            "unknown"
        } else {
            &state.tenant_relation
        },
    ));
    out.push(row(
        "Private lookup address",
        if state.lookup_address_available {
            "● available"
        } else {
            "○ unavailable"
        },
    ));
    out.push(String::new());

    out.push("Identity sources".into());
    out.push(row("Member resource ID", &state.member_id_shape));
    out.push(row("Member userId", &state.member_user_id_shape));
    out.push(row("Message sender ID", &state.preview_id_shape));
    out.push(row("Cached contact ID", &state.cached_id_shape));
    out.push(row(
        "Graph GUID from member",
        if state.member_guid {
            "● available"
        } else {
            "○ unavailable"
        },
    ));
    out.push(row(
        "Graph GUID from sender",
        if state.preview_guid {
            "● available"
        } else {
            "○ unavailable"
        },
    ));
    out.push(row(
        "Graph GUID from cache",
        if state.cached_guid {
            "● available"
        } else {
            "○ unavailable"
        },
    ));
    out.push(String::new());

    out.push("Presence capability".into());
    match state.remote.as_ref().map(|remote| &remote.presence_read_all) {
        Some(Ok(true)) => out.push(row("Presence.Read.All", "● granted")),
        Some(Ok(false)) => out.push(row("Presence.Read.All", "! missing from token")),
        Some(Err(error)) => out.push(row("Presence.Read.All", format!("× {error}"))),
        None => out.push(row(
            "Presence.Read.All",
            if state.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }
    out.push(row(
        "Contact presence",
        yes_no(app.session.config.presence_read, "enabled", "disabled"),
    ));
    out.push(row(
        "Identity supported",
        yes_no(
            state.presence_supported,
            "Graph presence candidate",
            "unsupported / no GUID",
        ),
    ));
    out.push(row(
        "Cross-tenant candidate",
        yes_no(state.cross_tenant_candidate, "yes", "no"),
    ));
    out.push(row(
        "Teams UPS path",
        if state.mri_candidate_source.is_empty() || state.mri_candidate_source == "none" {
            "○ no user MRI candidate"
        } else {
            "● MRI candidate available; resource token not probed"
        },
    ));
    out.push(String::new());

    out.push("Teams / Skype presence path".into());
    match state.remote.as_ref() {
        Some(remote) => {
            out.push(row(
                "Skype resource token",
                teams_step_value(&remote.teams.resource_token),
            ));
            out.push(row(
                "Teams authz / Skype token",
                teams_step_value(&remote.teams.authz),
            ));
            out.push(row(
                "Middle Tier MRI lookup",
                teams_step_value(&remote.teams.middle_tier_lookup),
            ));
            out.push(row(
                "Resolved MRI shape",
                if remote.teams.resolved_mri_shape == "none" {
                    "○ none".to_string()
                } else {
                    format!("● {}", remote.teams.resolved_mri_shape)
                },
            ));
            out.push(row(
                "Teams UPS presence",
                teams_step_value(&remote.teams.ups_presence),
            ));
        }
        None => {
            let value = if state.loading { "? loading" } else { "? not loaded" };
            out.push(row("Skype resource token", value));
            out.push(row("Teams authz / Skype token", value));
            out.push(row("Middle Tier MRI lookup", value));
            out.push(row("Resolved MRI shape", value));
            out.push(row("Teams UPS presence", value));
        }
    }
    out.push(String::new());

    out.push("Presence lookup".into());
    match state.remote.as_ref() {
        Some(remote) => {
            out.push(row("Batch Graph presence", probe_value(&remote.batch)));
            out.push(row("Direct Graph presence", probe_value(&remote.direct)));
        }
        None => {
            let value = if state.loading { "? loading" } else { "? not loaded" };
            out.push(row("Batch Graph presence", value));
            out.push(row("Direct Graph presence", value));
        }
    }
    out.push(String::new());

    out.push("Result".into());
    let result = if !app.session.config.presence_read {
        "○ disabled by configuration".to_string()
    } else if !state.presence_supported {
        "○ identity cannot be queried through Graph presence".to_string()
    } else if let Some(remote) = state.remote.as_ref() {
        match (&remote.batch, &remote.direct) {
            (
                PresenceProbe::Success {
                    availability,
                    activity,
                },
                _,
            )
            | (
                _,
                PresenceProbe::Success {
                    availability,
                    activity,
                },
            ) => format!("● {availability} · {activity}"),
            _ => "× Graph presence unavailable".to_string(),
        }
    } else if state.loading {
        "? loading".to_string()
    } else {
        "? not loaded".to_string()
    };
    out.push(row("Graph presence", result));

    out.join("\n")
}

pub fn text(app: &App) -> String {
    let mut out = Vec::new();
    out.push("m365-tui diagnostics".to_string());

    let remote = app.diagnostics.remote.as_ref();
    if let Some(remote) = remote {
        out.push(format!(
            "Generated: {}",
            remote
                .generated_at
                .with_timezone(&Local)
                .format("%Y-%m-%d %H:%M:%S %:z")
        ));
    } else if app.diagnostics.loading {
        out.push("Generated: loading…".into());
    }
    out.push(format!("Version: {}", env!("CARGO_PKG_VERSION")));
    out.push(String::new());

    out.push("Identity / Token".into());
    let account = app
        .me
        .as_ref()
        .and_then(|user| user.best_email())
        .unwrap_or("unknown");
    out.push(row("Account", account));
    if let Some(name) = app
        .me
        .as_ref()
        .and_then(|user| user.display_name.as_deref())
        .filter(|value| !value.trim().is_empty())
    {
        out.push(row("Display name", name));
    }
    out.push(row("Tenant", &app.session.config.tenant_id));
    out.push(row("Client", &app.session.config.client_id));

    let token = remote.and_then(|remote| remote.token.as_ref().ok());
    match remote.map(|remote| &remote.token) {
        Some(Ok(token)) => {
            out.push(row(
                "Token",
                if token.valid { "● valid" } else { "! expired" },
            ));
            out.push(row(
                "Token expires",
                token
                    .expires_at
                    .with_timezone(&Local)
                    .format("%Y-%m-%d %H:%M:%S %:z")
                    .to_string(),
            ));
        }
        Some(Err(error)) => out.push(row(
            "Token",
            format!("× {}", compact_error(error)),
        )),
        None => out.push(row(
            "Token",
            if app.diagnostics.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }
    out.push(String::new());

    out.push("Microsoft 365 work settings".into());
    out.push(row("Local time zone", local_timezone_name()));
    match remote.map(|remote| &remote.work_plan) {
        Some(Ok(plan)) => {
            out.push(row(
                "Work plan now",
                if plan.working_now {
                    "● working"
                } else {
                    "○ outside working plan"
                },
            ));
            if plan.recurrences.is_empty() {
                out.push(row("Recurring schedule", "○ none returned"));
            } else {
                for (index, recurrence) in plan.recurrences.iter().enumerate() {
                    let days = if recurrence.days.is_empty() {
                        "?".to_string()
                    } else {
                        recurrence.days.join(" ")
                    };
                    let zone = recurrence.time_zone.as_deref().unwrap_or("unknown zone");
                    let location = recurrence.location.as_deref().unwrap_or("unspecified");
                    out.push(row(
                        &format!("Schedule {}", index + 1),
                        format!(
                            "{days} · {}-{} · {zone} · {location}",
                            recurrence.start_time, recurrence.end_time
                        ),
                    ));
                }
            }
        }
        Some(Err(error)) => out.push(row(
            "Work plan",
            format!("× {}", compact_error(error)),
        )),
        None => out.push(row(
            "Work plan",
            if app.diagnostics.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }
    out.push(String::new());

    out.push("Microsoft Graph permissions".into());
    out.extend(permission_lines(app, token));
    out.push(String::new());

    out.push("Optional features".into());
    out.push(row(
        "Contact presence",
        yes_no(app.session.config.presence_read, "enabled", "disabled"),
    ));
    out.push(row(
        "Presence write",
        yes_no(
            app.session.config.can_write_presence(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "Teams channels",
        yes_no(
            app.session.config.can_read_teams(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "People search",
        yes_no(
            app.session.config.can_search_people(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "Profile photos",
        yes_no(
            app.session.config.can_read_profile_photos(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "Cache warmup",
        yes_no(app.session.config.teams_cache_warmup, "enabled", "disabled"),
    ));
    out.push(String::new());

    out.push("Teams / Presence".into());
    match app.my_presence.as_ref() {
        Some(presence) => {
            let availability = presence.availability.as_deref().unwrap_or("unknown");
            let activity = presence.activity.as_deref().unwrap_or("unknown");
            out.push(row(
                "Graph presence",
                format!("● {availability} · {activity}"),
            ));
        }
        None => out.push(row("Graph presence", "? not loaded")),
    }
    out.push(row("Skype Presence Service", "? not verified"));
    out.push(row("Skype Presence R/W", "? no Skype resource token"));
    out.push(String::new());

    out.push("Runtime".into());
    let push = match &app.push {
        PushState::Off => "○ poll only".to_string(),
        PushState::Connecting => "? connecting".to_string(),
        PushState::Live => "● live".to_string(),
        PushState::Failed(error) => format!("× failed: {}", compact_error(error)),
    };
    out.push(row("Push", push));
    out.push(row(
        "Persistent Teams cache",
        yes_no(
            app.session.config.teams_image_cache_dir.is_some(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "Terminal image protocol",
        yes_no(app.graphics.is_some(), "available", "unavailable"),
    ));
    out.push(row(
        "Clipboard helper",
        crate::clipboard::native_backend()
            .map(|value| format!("● {value}"))
            .unwrap_or_else(|| "○ none; diagnostics use log fallback".into()),
    ));
    out.push(row(
        "Last Graph poll",
        format!("{}s ago", app.poll_elapsed().as_secs()),
    ));
    if let Some(kb) = app.rss_kb {
        out.push(row("RSS", format!("{:.1} MiB", kb as f64 / 1024.0)));
    }

    out.join("\n")
}

fn save_log_named(stem: &str, text: &str) -> anyhow::Result<PathBuf> {
    use anyhow::Context;

    let name = format!(
        "{stem}-{}.log",
        Local::now().format("%Y%m%d-%H%M%S")
    );
    let path = std::env::temp_dir().join(name);

    #[cfg(unix)]
    {
        use std::io::Write;
        use std::os::unix::fs::OpenOptionsExt;

        let mut file = std::fs::OpenOptions::new()
            .create_new(true)
            .write(true)
            .mode(0o600)
            .open(&path)
            .with_context(|| format!("creating {}", path.display()))?;
        file.write_all(text.as_bytes())
            .with_context(|| format!("writing {}", path.display()))?;
        file.flush()
            .with_context(|| format!("flushing {}", path.display()))?;
    }

    #[cfg(not(unix))]
    {
        std::fs::write(&path, text)
            .with_context(|| format!("writing {}", path.display()))?;
    }

    Ok(path)
}

/// Write generic diagnostics to a private timestamped file.
pub fn save_log(text: &str) -> anyhow::Result<PathBuf> {
    save_log_named("m365-tui-diagnostics", text)
}

/// Write selected-contact diagnostics to a private timestamped file.
pub fn save_contact_log(text: &str) -> anyhow::Result<PathBuf> {
    save_log_named("m365-tui-contact-diagnostics", text)
}

#[cfg(test)]
mod tests {
    use super::{
        compact_error, identifier_shape, is_presence_mri_candidate, safe_graph_error_text,
    };

    #[test]
    fn compact_error_is_single_line_and_bounded() {
        let value = format!("first\nsecond {}", "x".repeat(300));
        let result = compact_error(&value);
        assert!(!result.contains('\n'));
        assert!(result.chars().count() <= 121);
    }

    #[test]
    fn contact_diagnostics_error_does_not_leak_identity() {
        let raw = r#"Graph request failed (404 Not Found): {"error":{"code":"Request_ResourceNotFound","message":"user 123e4567-e89b-12d3-a456-426614174000 was not found"}}"#;
        let result = safe_graph_error_text(raw);
        assert_eq!(result, "404 Not Found · Request_ResourceNotFound");
        assert!(!result.contains("123e4567"));
        assert!(!result.contains("was not found"));
    }

    #[test]
    fn identifier_shapes_are_useful_but_do_not_echo_values() {
        assert_eq!(
            identifier_shape(Some("123e4567-e89b-12d3-a456-426614174000")),
            "Entra GUID"
        );
        assert_eq!(
            identifier_shape(Some("8:live:.cid.abcdef123456")),
            "MRI consumer (8:live:)"
        );
        assert_eq!(
            identifier_shape(Some("8:orgid:123e4567-e89b-12d3-a456-426614174000")),
            "MRI orgid (8:orgid:)"
        );
        assert_eq!(
            identifier_shape(Some("8:teamsvisitor:opaque")),
            "MRI visitor (8:teamsvisitor:)"
        );
        assert_eq!(identifier_shape(Some("something-secret")), "opaque non-GUID");
        assert_eq!(identifier_shape(None), "missing");
    }

    #[test]
    fn ups_candidate_accepts_only_user_mri_shapes() {
        assert!(is_presence_mri_candidate(Some("8:live:.cid.abcdef")));
        assert!(is_presence_mri_candidate(Some(
            "8:orgid:123e4567-e89b-12d3-a456-426614174000"
        )));
        assert!(is_presence_mri_candidate(Some("8:teamsvisitor:opaque")));
        assert!(!is_presence_mri_candidate(Some(
            "123e4567-e89b-12d3-a456-426614174000"
        )));
        assert!(!is_presence_mri_candidate(Some("29:bot")));
        assert!(!is_presence_mri_candidate(None));
    }
}
