//! Read-only F6 diagnostics and safe export helpers.

use std::collections::HashSet;
use std::path::PathBuf;

use chrono::{DateTime, Local, Utc};
use m365_core::auth::TokenInfo;
use m365_core::config::NtfyMode;
use m365_core::models::{Chat, ChatMessage, ConversationMember};
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
pub struct ContactChatNameProbe {
    pub member_count: usize,
    pub peer_member_found: bool,
    pub member_display_name: bool,
    pub member_email: bool,
    pub member_user_id_shape: String,
    pub preview_present: bool,
    pub preview_sender_present: bool,
    pub preview_sender_is_peer: bool,
    pub preview_display_name: bool,
    pub preview_sender_id_shape: String,
    pub preview_sender_type: String,
}

#[derive(Debug, Clone)]
pub struct ContactMembersNameProbe {
    pub member_count: usize,
    pub peer_member_found: bool,
    pub member_display_name: bool,
    pub member_email: bool,
    pub member_user_id_shape: String,
}

#[derive(Debug, Clone)]
pub struct ContactMessagesNameProbe {
    pub messages_scanned: usize,
    pub peer_messages: usize,
    pub peer_messages_with_display_name: usize,
    pub peer_message_found: bool,
    pub sender_display_name: bool,
    pub sender_id_shape: String,
    pub sender_type: String,
}

#[derive(Debug, Clone)]
pub struct ContactDiagnosticsRemote {
    pub generated_at: DateTime<Utc>,
    pub presence_read_all: Result<bool, String>,
    pub batch: PresenceProbe,
    pub direct: PresenceProbe,
    pub teams: m365_core::teams_presence::Diagnostics,
    pub expanded_name: Result<ContactChatNameProbe, String>,
    pub members_name: Result<ContactMembersNameProbe, String>,
    pub messages_name: Result<ContactMessagesNameProbe, String>,
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
    pub current_member_count: usize,
    pub topic_available: bool,
    pub member_display_name_available: bool,
    pub member_email_available: bool,
    pub preview_present: bool,
    pub preview_sender_available: bool,
    pub preview_sender_is_peer: bool,
    pub preview_display_name_available: bool,
    pub cached_name_available: bool,
    pub chat_label_source: String,
    pub list_label_source: String,
    /// Raw one-to-one peer ID explicitly exposed by F7 for automation setup.
    pub automation_peer_id: Option<String>,
    pub automation_peer_configured: bool,
    pub automation_config_source: String,
    pub remote: Option<ContactDiagnosticsRemote>,
}

fn nonempty(value: Option<&str>) -> Option<&str> {
    value.map(str::trim).filter(|value| !value.is_empty())
}

fn peer_member<'a>(
    members: &'a [ConversationMember],
    me_id: Option<&str>,
) -> Option<&'a ConversationMember> {
    members.iter().find(|member| {
        me_id
            .map(|id| member.user_id.as_deref() != Some(id))
            .unwrap_or(true)
    })
}

fn safe_identity_type(value: Option<&str>) -> String {
    let Some(value) = nonempty(value) else {
        return "missing".into();
    };
    if value.len() <= 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        value.to_string()
    } else {
        "other".into()
    }
}

pub fn contact_chat_name_probe(chat: &Chat, me_id: Option<&str>) -> ContactChatNameProbe {
    let peer = peer_member(&chat.members, me_id);
    let preview_user = chat
        .last_message_preview
        .as_ref()
        .and_then(|preview| preview.from.as_ref())
        .and_then(|from| from.user.as_ref());
    let preview_sender_is_peer = preview_user.is_some_and(|user| {
        me_id
            .map(|id| user.id.as_deref() != Some(id))
            .unwrap_or(true)
    });

    ContactChatNameProbe {
        member_count: chat.members.len(),
        peer_member_found: peer.is_some(),
        member_display_name: peer
            .and_then(|member| nonempty(member.display_name.as_deref()))
            .is_some(),
        member_email: peer
            .and_then(|member| nonempty(member.email.as_deref()))
            .is_some(),
        member_user_id_shape: identifier_shape(peer.and_then(|member| member.user_id.as_deref()))
            .to_string(),
        preview_present: chat.last_message_preview.is_some(),
        preview_sender_present: preview_user.is_some(),
        preview_sender_is_peer,
        preview_display_name: preview_sender_is_peer
            && preview_user
                .and_then(|user| nonempty(user.display_name.as_deref()))
                .is_some(),
        preview_sender_id_shape: identifier_shape(preview_user.and_then(|user| user.id.as_deref()))
            .to_string(),
        preview_sender_type: safe_identity_type(
            preview_user.and_then(|user| user.user_identity_type.as_deref()),
        ),
    }
}

pub fn contact_members_name_probe(
    members: &[ConversationMember],
    me_id: Option<&str>,
) -> ContactMembersNameProbe {
    let peer = peer_member(members, me_id);
    ContactMembersNameProbe {
        member_count: members.len(),
        peer_member_found: peer.is_some(),
        member_display_name: peer
            .and_then(|member| nonempty(member.display_name.as_deref()))
            .is_some(),
        member_email: peer
            .and_then(|member| nonempty(member.email.as_deref()))
            .is_some(),
        member_user_id_shape: identifier_shape(peer.and_then(|member| member.user_id.as_deref()))
            .to_string(),
    }
}

pub fn contact_messages_name_probe(
    messages: &[ChatMessage],
    me_id: Option<&str>,
) -> ContactMessagesNameProbe {
    let peers: Vec<_> = messages
        .iter()
        .filter_map(|message| message.from.as_ref()?.user.as_ref())
        .filter(|user| {
            me_id
                .map(|id| user.id.as_deref() != Some(id))
                .unwrap_or(true)
        })
        .collect();

    let named_peer = peers
        .iter()
        .copied()
        .find(|user| nonempty(user.display_name.as_deref()).is_some());
    let representative_peer = named_peer.or_else(|| peers.first().copied());
    let peer_messages_with_display_name = peers
        .iter()
        .filter(|user| nonempty(user.display_name.as_deref()).is_some())
        .count();

    ContactMessagesNameProbe {
        messages_scanned: messages.len(),
        peer_messages: peers.len(),
        peer_messages_with_display_name,
        peer_message_found: !peers.is_empty(),
        sender_display_name: named_peer.is_some(),
        sender_id_shape: identifier_shape(representative_peer.and_then(|user| user.id.as_deref()))
            .to_string(),
        sender_type: safe_identity_type(
            representative_peer.and_then(|user| user.user_identity_type.as_deref()),
        ),
    }
}

fn row(label: &str, value: impl AsRef<str>) -> String {
    format!("  {label:<30} {}", value.as_ref())
}

fn ntfy_mode(mode: NtfyMode) -> &'static str {
    match mode {
        NtfyMode::Never => "never",
        NtfyMode::Always => "always",
        NtfyMode::Away => "away",
        NtfyMode::AlwaysWd => "alwayswd",
        NtfyMode::AwayWd => "awaywd",
    }
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
        "People.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "MailboxSettings.ReadWrite",
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

    let code = ["\"code\":\"", "\"code\": \""]
        .into_iter()
        .find_map(|needle| {
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
        availability: presence
            .availability
            .as_deref()
            .unwrap_or("unknown")
            .to_string(),
        activity: presence
            .activity
            .as_deref()
            .unwrap_or("unknown")
            .to_string(),
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

    out.push(String::new());

    out.push("Automation".into());
    out.push(row(
        "Peer ID",
        state
            .automation_peer_id
            .as_deref()
            .unwrap_or("○ unavailable"),
    ));
    out.push(row(
        "Automation peer",
        if state.automation_peer_configured {
            "● configured"
        } else {
            "○ not configured"
        },
    ));
    out.push(row("Config variable", "M365_TEAMS_AUTOMATION_PEERS"));
    out.push(row(
        "Config source",
        if state.automation_config_source.is_empty() {
            "unknown"
        } else {
            &state.automation_config_source
        },
    ));

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

    out.push("Name resolution".into());
    out.push(row(
        "Current roster members",
        state.current_member_count.to_string(),
    ));
    out.push(row(
        "Current topic",
        yes_no(state.topic_available, "available", "unavailable"),
    ));
    out.push(row(
        "Current member displayName",
        yes_no(
            state.member_display_name_available,
            "available",
            "unavailable",
        ),
    ));
    out.push(row(
        "Current member email",
        yes_no(state.member_email_available, "available", "unavailable"),
    ));
    out.push(row(
        "Current preview",
        yes_no(state.preview_present, "present", "missing"),
    ));
    out.push(row(
        "Current preview sender",
        yes_no(state.preview_sender_available, "present", "missing"),
    ));
    out.push(row(
        "Preview sender is peer",
        yes_no(state.preview_sender_is_peer, "yes", "no / unknown"),
    ));
    out.push(row(
        "Preview sender displayName",
        yes_no(
            state.preview_display_name_available,
            "available",
            "unavailable",
        ),
    ));
    out.push(row(
        "Cached / learned name",
        yes_no(state.cached_name_available, "available", "unavailable"),
    ));
    out.push(row(
        "Chat::label source",
        if state.chat_label_source.is_empty() {
            "unknown"
        } else {
            &state.chat_label_source
        },
    ));
    out.push(row(
        "Rendered list source",
        if state.list_label_source.is_empty() {
            "unknown"
        } else {
            &state.list_label_source
        },
    ));
    out.push(String::new());

    out.push("Fresh Graph name probes".into());
    match state.remote.as_ref().map(|remote| &remote.expanded_name) {
        Some(Ok(probe)) => {
            out.push(row("Expanded chat probe", "● loaded"));
            out.push(row(
                "Expanded roster members",
                probe.member_count.to_string(),
            ));
            out.push(row(
                "Expanded peer member",
                yes_no(probe.peer_member_found, "found", "missing"),
            ));
            out.push(row(
                "Expanded member displayName",
                yes_no(probe.member_display_name, "available", "unavailable"),
            ));
            out.push(row(
                "Expanded member email",
                yes_no(probe.member_email, "available", "unavailable"),
            ));
            out.push(row("Expanded member userId", &probe.member_user_id_shape));
            out.push(row(
                "Expanded preview",
                yes_no(probe.preview_present, "present", "missing"),
            ));
            out.push(row(
                "Expanded preview sender",
                yes_no(probe.preview_sender_present, "present", "missing"),
            ));
            out.push(row(
                "Expanded sender is peer",
                yes_no(probe.preview_sender_is_peer, "yes", "no / unknown"),
            ));
            out.push(row(
                "Expanded sender displayName",
                yes_no(probe.preview_display_name, "available", "unavailable"),
            ));
            out.push(row("Expanded sender ID", &probe.preview_sender_id_shape));
            out.push(row("Expanded sender type", &probe.preview_sender_type));
        }
        Some(Err(error)) => out.push(row(
            "Expanded chat probe",
            format!("× {}", compact_error(error)),
        )),
        None => out.push(row(
            "Expanded chat probe",
            if state.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }

    match state.remote.as_ref().map(|remote| &remote.members_name) {
        Some(Ok(probe)) => {
            out.push(row("Dedicated /members probe", "● loaded"));
            out.push(row(
                "/members roster members",
                probe.member_count.to_string(),
            ));
            out.push(row(
                "/members peer member",
                yes_no(probe.peer_member_found, "found", "missing"),
            ));
            out.push(row(
                "/members displayName",
                yes_no(probe.member_display_name, "available", "unavailable"),
            ));
            out.push(row(
                "/members email",
                yes_no(probe.member_email, "available", "unavailable"),
            ));
            out.push(row("/members userId", &probe.member_user_id_shape));
        }
        Some(Err(error)) => out.push(row(
            "Dedicated /members probe",
            format!("× {}", compact_error(error)),
        )),
        None => out.push(row(
            "Dedicated /members probe",
            if state.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }

    match state.remote.as_ref().map(|remote| &remote.messages_name) {
        Some(Ok(probe)) => {
            out.push(row("Recent messages probe", "● loaded"));
            out.push(row("Messages scanned", probe.messages_scanned.to_string()));
            out.push(row("Peer messages", probe.peer_messages.to_string()));
            out.push(row(
                "Peer messages with displayName",
                probe.peer_messages_with_display_name.to_string(),
            ));
            out.push(row(
                "Recent peer message",
                yes_no(probe.peer_message_found, "found", "not found"),
            ));
            out.push(row(
                "Any peer sender displayName",
                yes_no(probe.sender_display_name, "available", "unavailable"),
            ));
            out.push(row("Recent sender ID", &probe.sender_id_shape));
            out.push(row("Recent sender type", &probe.sender_type));
        }
        Some(Err(error)) => out.push(row(
            "Recent messages probe",
            format!("× {}", compact_error(error)),
        )),
        None => out.push(row(
            "Recent messages probe",
            if state.loading {
                "? loading"
            } else {
                "? not loaded"
            },
        )),
    }
    out.push(String::new());

    out.push("Presence capability".into());
    match state
        .remote
        .as_ref()
        .map(|remote| &remote.presence_read_all)
    {
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
            let value = if state.loading {
                "? loading"
            } else {
                "? not loaded"
            };
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
            let value = if state.loading {
                "? loading"
            } else {
                "? not loaded"
            };
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
                if token.valid {
                    "● valid"
                } else {
                    "! expired"
                },
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
        Some(Err(error)) => out.push(row("Token", format!("× {}", compact_error(error)))),
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
    out.push(row(
        "Local time zone",
        crate::out_of_office::local_timezone_name(),
    ));
    out.push(row("NTFY mode", ntfy_mode(app.session.config.ntfy)));
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
        Some(Err(error)) => out.push(row("Work plan", format!("× {}", compact_error(error)))),
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
        "Primary presence",
        yes_no(app.session.config.presence_primary, "enabled", "disabled"),
    ));
    out.push(row(
        "Teams channels",
        yes_no(app.session.config.can_read_teams(), "enabled", "disabled"),
    ));
    out.push(row(
        "Teams file images",
        yes_no(app.session.config.teams_file_images, "enabled", "disabled"),
    ));
    out.push(row(
        "Teams Markdown",
        yes_no(app.session.config.teams_markdown, "enabled", "disabled"),
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
        "Directory profile",
        yes_no(app.session.config.directory_profile, "enabled", "disabled"),
    ));
    out.push(row(
        "Out of office",
        yes_no(app.session.config.out_of_office, "enabled", "disabled"),
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
    let managed = app.presence_state_diagnostics();
    out.push(row("Presence manager", managed.mode.label()));
    out.push(row("App session target", managed.desired.label()));
    out.push(row(
        "App session confirmed",
        managed
            .confirmed
            .map(|target| target.label())
            .unwrap_or_else(|| "unknown".into()),
    ));
    out.push(row(
        "Presence write",
        managed
            .in_flight
            .map(|write| write.label())
            .unwrap_or_else(|| {
                if managed.retry_waiting {
                    "retry waiting".into()
                } else {
                    "idle".into()
                }
            }),
    ));
    out.push(row("Skype Presence Service", "? not verified"));
    out.push(row("Skype Presence R/W", "? no Skype resource token"));

    let activity = app.presence_activity_diagnostics();
    out.push(row("Presence activity source", activity.requested));
    out.push(row(
        "Activity backend",
        if activity.degraded {
            format!("! {}", activity.backend)
        } else {
            format!("● {}", activity.backend)
        },
    ));
    out.push(row("Session type", &activity.session_type));
    out.push(row("Desktop state", activity.state));
    out.push(row(
        "Desktop idle",
        activity
            .idle_for
            .map(|idle| format!("{:.1}s", idle.as_secs_f64()))
            .unwrap_or_else(|| "unknown".into()),
    ));
    out.push(row(
        "Screen locked",
        match activity.locked {
            Some(true) => "yes",
            Some(false) => "no",
            None => "unknown",
        },
    ));
    out.push(row(
        "Lock restore",
        if app.presence_lock_restore_enabled() {
            "enabled"
        } else {
            "disabled"
        },
    ));
    if let Some(detail) = activity.detail.as_deref() {
        out.push(row("Activity detail", detail));
    }
    out.push(row(
        "Idle threshold",
        format!(
            "{}s",
            app.session
                .config
                .presence_available_timeout_min
                .saturating_mul(60)
        ),
    ));
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
            app.session.config.cache_dir.is_some(),
            "enabled",
            "disabled",
        ),
    ));
    out.push(row(
        "Persistent cache limit",
        format!("{} MiB", app.session.config.cache_max_mb),
    ));
    let teams_poll = app.teams_poll_diagnostics();
    out.push(row(
        "Teams hot chats",
        format!("{} configured", app.session.config.teams_hot_chats),
    ));
    out.push(row(
        "Teams poll budget",
        format!(
            "{:.1} rps configured · {:.1} rps runtime",
            app.session.config.teams_poll_budget_rps, teams_poll.runtime_budget_rps
        ),
    ));
    out.push(row(
        "Teams poll tiers",
        format!(
            "HOT {} · WARM {} · COOL {} · NORMAL {}",
            teams_poll.hot, teams_poll.warm, teams_poll.cool, teams_poll.normal
        ),
    ));
    out.push(row(
        "Teams poll work",
        format!(
            "{} hot pending · {} foreground · {} reserved",
            teams_poll.hot_pending, teams_poll.foreground_in_flight, teams_poll.reserved_chats
        ),
    ));

    out.push(row(
        "Kitty image protocol",
        yes_no(app.kitty_images_available(), "available", "unavailable"),
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

    let name = format!("{stem}-{}.log", Local::now().format("%Y%m%d-%H%M%S"));
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
        std::fs::write(&path, text).with_context(|| format!("writing {}", path.display()))?;
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
        assert_eq!(
            identifier_shape(Some("something-secret")),
            "opaque non-GUID"
        );
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
