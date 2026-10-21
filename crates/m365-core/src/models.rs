//! Serde models for the subset of Microsoft Graph resources the TUI uses.
//! Fields are intentionally partial — Graph returns far more than we render.

use serde::{Deserialize, Serialize};

/// The signed-in user (`/me`).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct User {
    pub id: String,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub mail: Option<String>,
    #[serde(default)]
    pub user_principal_name: Option<String>,
    #[serde(default)]
    pub job_title: Option<String>,
}

impl User {
    pub fn best_email(&self) -> Option<&str> {
        self.mail.as_deref().or(self.user_principal_name.as_deref())
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EmailAddress {
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub address: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Recipient {
    #[serde(default)]
    pub email_address: Option<EmailAddress>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ItemBody {
    #[serde(default)]
    pub content_type: Option<String>,
    #[serde(default)]
    pub content: Option<String>,
}

// ---------------------------------------------------------------------------
// Mail
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MailFolder {
    pub id: String,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub unread_item_count: Option<i64>,
    #[serde(default)]
    pub total_item_count: Option<i64>,
    #[serde(rename = "childFolderCount", default)]
    pub child_folder_count: Option<u32>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MailMessage {
    pub id: String,
    #[serde(default)]
    pub conversation_id: Option<String>,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub body_preview: Option<String>,
    #[serde(default)]
    pub from: Option<Recipient>,
    #[serde(default)]
    pub to_recipients: Vec<Recipient>,
    #[serde(default)]
    pub cc_recipients: Vec<Recipient>,
    #[serde(default)]
    pub bcc_recipients: Vec<Recipient>,
    #[serde(default)]
    pub received_date_time: Option<String>,
    #[serde(default)]
    pub sent_date_time: Option<String>,
    #[serde(default)]
    pub is_read: Option<bool>,
    #[serde(default)]
    pub has_attachments: Option<bool>,
    #[serde(default)]
    pub web_link: Option<String>,
    /// Populated only when a single message is fetched with `$select=body`.
    #[serde(default)]
    pub body: Option<ItemBody>,
}

impl MailMessage {
    pub fn mail_time(&self) -> Option<&str> {
        self.received_date_time
            .as_deref()
            .or(self.sent_date_time.as_deref())
    }
    pub fn sender_name(&self) -> String {
        self.from
            .as_ref()
            .and_then(|r| r.email_address.as_ref())
            .and_then(|e| e.name.clone().or_else(|| e.address.clone()))
            .unwrap_or_else(|| "(unknown)".into())
    }

    pub fn sender_address(&self) -> Option<String> {
        self.from
            .as_ref()
            .and_then(|r| r.email_address.as_ref())
            .and_then(|e| e.address.clone())
    }
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DateTimeTimeZone {
    pub date_time: String,
    #[serde(default)]
    pub time_zone: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct AutomaticRepliesSetting {
    #[serde(default)]
    pub status: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_audience: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheduled_start_date_time: Option<DateTimeTimeZone>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheduled_end_date_time: Option<DateTimeTimeZone>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub internal_reply_message: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_reply_message: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MailboxSettingsSnapshot {
    #[serde(default)]
    pub time_zone: Option<String>,
    #[serde(default)]
    pub automatic_replies_setting: AutomaticRepliesSetting,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Attendee {
    #[serde(default)]
    pub email_address: Option<EmailAddress>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResponseStatus {
    #[serde(default)]
    pub response: Option<String>,
    #[serde(default)]
    pub time: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Event {
    pub id: String,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub start: Option<DateTimeTimeZone>,
    #[serde(default)]
    pub end: Option<DateTimeTimeZone>,
    #[serde(default)]
    pub organizer: Option<Recipient>,
    #[serde(default)]
    pub attendees: Vec<Attendee>,
    #[serde(default)]
    pub location: Option<Location>,
    #[serde(default)]
    pub is_online_meeting: Option<bool>,
    #[serde(default)]
    pub online_meeting: Option<OnlineMeetingInfo>,
    #[serde(default)]
    pub body_preview: Option<String>,
    #[serde(default)]
    pub response_status: Option<ResponseStatus>,
    #[serde(default)]
    pub is_organizer: Option<bool>,
    #[serde(default)]
    pub is_cancelled: Option<bool>,
    #[serde(default)]
    pub is_all_day: Option<bool>,
    #[serde(default)]
    pub response_requested: Option<bool>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Location {
    #[serde(default)]
    pub display_name: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OnlineMeetingInfo {
    #[serde(default)]
    pub join_url: Option<String>,
}

// ---------------------------------------------------------------------------
// Teams: chats, channels, messages
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Chat {
    pub id: String,
    #[serde(default)]
    pub topic: Option<String>,
    #[serde(default)]
    pub chat_type: Option<String>,
    #[serde(default)]
    pub tenant_id: Option<String>,
    #[serde(default)]
    pub last_updated_date_time: Option<String>,
    #[serde(default)]
    pub members: Vec<ConversationMember>,
    #[serde(default)]
    pub last_message_preview: Option<LastMessagePreview>,
    #[serde(default)]
    pub viewpoint: Option<ChatViewpoint>,
}

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChatViewpoint {
    #[serde(default)]
    pub is_hidden: Option<bool>,
    #[serde(default)]
    pub last_message_read_date_time: Option<String>,
}

impl Chat {
    /// A display label: the explicit topic, else the member names joined.
    ///
    /// Federated one-to-one chats can omit `members[].displayName`. In that
    /// case, prefer the already-expanded last-message sender name before
    /// falling back to the peer email or the raw chat type.
    pub fn label(&self, me_id: Option<&str>) -> String {
        if let Some(t) = self.topic.as_ref().filter(|t| !t.is_empty()) {
            return t.clone();
        }

        let peer = |member: &&ConversationMember| {
            me_id
                .map(|id| member.user_id.as_deref() != Some(id))
                .unwrap_or(true)
        };

        let names: Vec<String> = self
            .members
            .iter()
            .filter(peer)
            .filter_map(|m| m.display_name.as_deref())
            .map(str::trim)
            .filter(|name| !name.is_empty())
            .map(str::to_string)
            .collect();
        if !names.is_empty() {
            return names.join(", ");
        }

        if self
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
        {
            if let Some(name) = self
                .last_message_preview
                .as_ref()
                .and_then(|preview| preview.from.as_ref())
                .and_then(|from| from.user.as_ref())
                .filter(|user| {
                    me_id
                        .map(|id| user.id.as_deref() != Some(id))
                        .unwrap_or(true)
                })
                .and_then(|user| user.display_name.as_deref())
                .map(str::trim)
                .filter(|name| !name.is_empty())
            {
                return name.to_string();
            }

            if let Some(email) = self
                .members
                .iter()
                .filter(peer)
                .filter_map(|member| member.email.as_deref())
                .map(str::trim)
                .find(|email| !email.is_empty())
            {
                return email.to_string();
            }
        }

        self.chat_type.clone().unwrap_or_else(|| "chat".into())
    }

    /// Directory user id of the other participant in a one-to-one chat.
    pub fn peer_user_id<'a>(&'a self, me_id: Option<&str>) -> Option<&'a str> {
        if !self
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
        {
            return None;
        }

        self.members
            .iter()
            .filter_map(|member| member.user_id.as_deref())
            .find(|id| me_id != Some(*id))
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConversationMember {
    #[serde(rename = "@odata.type", default)]
    pub odata_type: Option<String>,
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub user_id: Option<String>,
    #[serde(default)]
    pub tenant_id: Option<String>,
    #[serde(default)]
    pub email: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LastMessagePreview {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub body: Option<ItemBody>,
    #[serde(default)]
    pub created_date_time: Option<String>,
    #[serde(default)]
    pub from: Option<IdentitySet>,
}

/// An `@mention` inside a Teams message.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChatMessageMention {
    #[serde(default)]
    pub mentioned: Option<IdentitySet>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Team {
    pub id: String,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub description: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Channel {
    pub id: String,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub description: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct IdentitySet {
    #[serde(default)]
    pub user: Option<Identity>,
    #[serde(default)]
    pub application: Option<Identity>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Identity {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub user_identity_type: Option<String>,
    #[serde(default)]
    pub tenant_id: Option<String>,
}

/// A Teams `chatMessage` (channel or chat).
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChatMessage {
    #[serde(default)]
    pub last_edited_date_time: Option<String>,
    #[serde(default)]
    pub reply_to_id: Option<String>,
    pub id: String,
    #[serde(default)]
    pub created_date_time: Option<String>,
    #[serde(default)]
    pub from: Option<IdentitySet>,
    #[serde(default)]
    pub body: Option<ItemBody>,
    #[serde(default)]
    pub message_type: Option<String>,
    /// Raw Microsoft Graph details for a `systemEventMessage`.
    ///
    /// Keep this data-form representation in the model/cache and derive the
    /// human-readable UI text at render time so presentation changes never
    /// require a cache migration.
    #[serde(default)]
    pub event_detail: Option<serde_json::Value>,
    #[serde(default)]
    pub deleted_date_time: Option<String>,
    #[serde(default)]
    pub reactions: Vec<MessageReaction>,
    #[serde(default)]
    pub attachments: Vec<MessageAttachment>,
    #[serde(default)]
    pub mentions: Vec<ChatMessageMention>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MessageReaction {
    #[serde(default)]
    pub reaction_type: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SystemEventClass {
    Useful,
    Noise,
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SystemEventDisplay {
    pub class: SystemEventClass,
    pub text: String,
    pub url: Option<String>,
}

fn event_kind(detail: &serde_json::Value) -> Option<&str> {
    detail
        .get("@odata.type")
        .and_then(serde_json::Value::as_str)
        .and_then(|value| value.rsplit('.').next())
        .map(|value| value.trim_start_matches('#'))
}

fn identity_display_name(value: &serde_json::Value) -> Option<String> {
    for kind in ["user", "application", "device"] {
        if let Some(name) = value
            .get(kind)
            .and_then(|identity| identity.get("displayName"))
            .and_then(serde_json::Value::as_str)
            .map(str::trim)
            .filter(|value| !value.is_empty())
        {
            return Some(name.to_string());
        }
    }

    value
        .get("displayName")
        .and_then(serde_json::Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
}

fn event_initiator(detail: &serde_json::Value) -> Option<String> {
    detail.get("initiator").and_then(identity_display_name)
}

fn event_members(detail: &serde_json::Value) -> Vec<String> {
    detail
        .get("members")
        .and_then(serde_json::Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(identity_display_name)
        .collect()
}

fn names_label(names: &[String], singular: &str, plural: &str) -> String {
    match names {
        [] => singular.to_string(),
        [one] => one.clone(),
        [one, two] => format!("{one} and {two}"),
        _ => format!("{} {plural}", names.len()),
    }
}

fn with_initiator(text: String, detail: &serde_json::Value) -> String {
    match event_initiator(detail) {
        Some(name) => format!("{text} (by {name})"),
        None => text,
    }
}

fn system_event_label(kind: &str) -> String {
    let value = kind
        .trim_end_matches("EventMessageDetail")
        .trim_end_matches("MessageDetail");
    let mut out = String::new();
    for (index, ch) in value.chars().enumerate() {
        if index > 0 && ch.is_uppercase() {
            out.push(' ');
        }
        if index == 0 {
            out.extend(ch.to_uppercase());
        } else {
            out.push(ch);
        }
    }
    if out.is_empty() {
        "System event".to_string()
    } else {
        out
    }
}

impl ChatMessage {
    /// The sender's user id, when the message came from a person.
    pub fn author_id(&self) -> Option<&str> {
        self.from.as_ref()?.user.as_ref()?.id.as_deref()
    }

    pub fn author(&self) -> String {
        self.from
            .as_ref()
            .and_then(|f| f.user.as_ref().or(f.application.as_ref()))
            .and_then(|i| i.display_name.clone())
            .unwrap_or_else(|| "(system)".into())
    }

    pub fn text(&self) -> String {
        self.body
            .as_ref()
            .and_then(|b| b.content.clone())
            .unwrap_or_default()
    }

    pub fn is_system_event(&self) -> bool {
        if self
            .message_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("systemEventMessage"))
        {
            return true;
        }

        // eventDetail is specific to Teams system-event messages and also
        // covers Graph responses where messageType arrived as unknownFutureValue.
        if self.event_detail.is_some() {
            return true;
        }

        // Old persistent-cache entries written before eventDetail was stored
        // still carry the Teams marker in the message body.
        self.body
            .as_ref()
            .and_then(|body| body.content.as_deref())
            .is_some_and(|content| content.to_ascii_lowercase().contains("<systemeventmessage"))
    }

    /// Classify and humanize a Teams system event without changing the raw
    /// Graph payload stored in this message.
    pub fn system_event_display(&self) -> Option<SystemEventDisplay> {
        if !self.is_system_event() {
            return None;
        }

        let Some(detail) = self.event_detail.as_ref() else {
            return Some(SystemEventDisplay {
                class: SystemEventClass::Unknown,
                text: "System event (details not cached)".to_string(),
                url: None,
            });
        };

        let Some(kind) = event_kind(detail) else {
            return Some(SystemEventDisplay {
                class: SystemEventClass::Unknown,
                text: "System event (unknown type)".to_string(),
                url: None,
            });
        };

        let members = event_members(detail);
        let (class, text, url) = match kind {
            "callRecordingEventMessageDetail" => {
                let name = detail
                    .get("callRecordingDisplayName")
                    .and_then(serde_json::Value::as_str)
                    .map(str::trim)
                    .filter(|value| !value.is_empty());
                let text = match name {
                    Some(name) => format!("🎥 Recording available: {name}"),
                    None => "🎥 Meeting recording available".to_string(),
                };
                let url = detail
                    .get("callRecordingUrl")
                    .and_then(serde_json::Value::as_str)
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(str::to_string);
                (SystemEventClass::Useful, text, url)
            }
            "callTranscriptEventMessageDetail" => (
                SystemEventClass::Useful,
                "📝 Meeting transcript available".to_string(),
                None,
            ),
            "chatRenamedEventMessageDetail" => {
                let name = detail
                    .get("chatDisplayName")
                    .and_then(serde_json::Value::as_str)
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .unwrap_or("(unnamed)");
                (
                    SystemEventClass::Useful,
                    with_initiator(format!("Chat renamed to \"{name}\""), detail),
                    None,
                )
            }
            "membersAddedEventMessageDetail" => (
                SystemEventClass::Useful,
                with_initiator(
                    format!(
                        "{} added to the chat",
                        names_label(&members, "A member was", "members were")
                    ),
                    detail,
                ),
                None,
            ),
            "membersDeletedEventMessageDetail" => (
                SystemEventClass::Useful,
                with_initiator(
                    format!(
                        "{} removed from the chat",
                        names_label(&members, "A member was", "members were")
                    ),
                    detail,
                ),
                None,
            ),
            "messagePinnedEventMessageDetail" => (
                SystemEventClass::Useful,
                with_initiator("📌 A message was pinned".to_string(), detail),
                None,
            ),
            "messageUnpinnedEventMessageDetail" => (
                SystemEventClass::Useful,
                with_initiator("A message was unpinned".to_string(), detail),
                None,
            ),

            "callStartedEventMessageDetail" => (
                SystemEventClass::Noise,
                with_initiator("Call started".to_string(), detail),
                None,
            ),
            "callEndedEventMessageDetail" => {
                (SystemEventClass::Noise, "Call ended".to_string(), None)
            }
            "membersJoinedEventMessageDetail" => (
                SystemEventClass::Noise,
                format!(
                    "{} joined the chat",
                    names_label(&members, "A member", "members")
                ),
                None,
            ),
            "membersLeftEventMessageDetail" => (
                SystemEventClass::Noise,
                format!(
                    "{} left the chat",
                    names_label(&members, "A member", "members")
                ),
                None,
            ),
            "meetingPolicyUpdatedEventMessageDetail" => (
                SystemEventClass::Noise,
                "Meeting policy updated".to_string(),
                None,
            ),
            "tabUpdatedEventMessageDetail" => (
                SystemEventClass::Noise,
                "Teams tab updated".to_string(),
                None,
            ),
            "teamsAppInstalledEventMessageDetail" => (
                SystemEventClass::Noise,
                "Teams app installed".to_string(),
                None,
            ),
            "teamsAppRemovedEventMessageDetail" => (
                SystemEventClass::Noise,
                "Teams app removed".to_string(),
                None,
            ),
            "teamsAppUpgradedEventMessageDetail" => (
                SystemEventClass::Noise,
                "Teams app upgraded".to_string(),
                None,
            ),
            _ => (
                SystemEventClass::Unknown,
                format!("System event: {}", system_event_label(kind)),
                None,
            ),
        };

        Some(SystemEventDisplay { class, text, url })
    }

    /// A short plain-text excerpt of the body, for quoting.
    pub fn text_preview(&self, max: usize) -> String {
        let mut out = String::new();
        let mut in_tag = false;
        for c in self.text().chars() {
            match c {
                '<' => in_tag = true,
                '>' => in_tag = false,
                _ if !in_tag => out.push(c),
                _ => {}
            }
        }
        let flat: String = out.split_whitespace().collect::<Vec<_>>().join(" ");
        if flat.chars().count() <= max {
            flat
        } else {
            flat.chars().take(max).collect::<String>() + "…"
        }
    }

    /// The message this one replies to, if any.
    ///
    /// Teams models a chat reply as a `messageReference` attachment — the body
    /// only carries an empty `<attachment>` tag — so the quote has to be read
    /// from there rather than from the HTML.
    pub fn quoted(&self) -> Option<QuotedMessage> {
        let att = self.attachments.iter().find(|a| {
            a.content_type
                .as_deref()
                .is_some_and(|t| t.eq_ignore_ascii_case("messageReference"))
        })?;
        let reference: MessageReference = serde_json::from_str(att.content.as_deref()?).ok()?;
        Some(QuotedMessage {
            message_id: reference.message_id.unwrap_or_default(),
            author: reference
                .message_sender
                .as_ref()
                .and_then(|s| s.user.as_ref())
                .and_then(|u| u.display_name.clone())
                .unwrap_or_else(|| "(unknown)".into()),
            preview: reference.message_preview.unwrap_or_default(),
        })
    }

    /// A short display of reactions, e.g. `👍 ❤️`. Maps the classic Teams
    /// reaction names to emoji; unicode reactions pass through as-is.
    pub fn reactions_summary(&self) -> Option<String> {
        if self.reactions.is_empty() {
            return None;
        }
        let mut out = String::new();
        for r in &self.reactions {
            let ty = r.reaction_type.as_deref().unwrap_or("");
            let e = match ty {
                "like" => "👍",
                "heart" => "❤️",
                "laugh" => "😆",
                "surprised" => "😮",
                "sad" => "😢",
                "angry" => "😠",
                other => other,
            };
            if !e.is_empty() {
                out.push_str(e);
                out.push(' ');
            }
        }
        let out = out.trim_end().to_string();
        if out.is_empty() {
            None
        } else {
            Some(out)
        }
    }
}

// ---------------------------------------------------------------------------
// People & presence
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Person {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub given_name: Option<String>,
    #[serde(default)]
    pub surname: Option<String>,
    #[serde(default)]
    pub job_title: Option<String>,
    #[serde(default)]
    pub company_name: Option<String>,
    #[serde(default)]
    pub department: Option<String>,
    #[serde(default)]
    pub office_location: Option<String>,
    #[serde(default)]
    pub user_principal_name: Option<String>,
    #[serde(default)]
    pub im_address: Option<String>,
    #[serde(default)]
    pub phones: Vec<Phone>,
    #[serde(default)]
    pub person_type: Option<PersonType>,
    #[serde(default, rename = "scoredEmailAddresses")]
    pub scored_email_addresses: Vec<ScoredEmailAddress>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Phone {
    #[serde(default, rename = "type")]
    pub phone_type: Option<String>,
    #[serde(default)]
    pub number: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PersonType {
    #[serde(default)]
    pub class: Option<String>,
    #[serde(default)]
    pub subclass: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ScoredEmailAddress {
    #[serde(default)]
    pub address: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutOfOfficeSettings {
    #[serde(default)]
    pub is_out_of_office: Option<bool>,
    #[serde(default)]
    pub message: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresenceStatusMessage {
    #[serde(default)]
    pub message: Option<ItemBody>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Presence {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub availability: Option<String>,
    #[serde(default)]
    pub activity: Option<String>,
    #[serde(default)]
    pub out_of_office_settings: Option<OutOfOfficeSettings>,
    #[serde(default)]
    pub status_message: Option<PresenceStatusMessage>,
}

impl Presence {
    /// User-authored Teams presence status message, if one is currently set.
    pub fn status_message_text(&self) -> Option<&str> {
        self.status_message
            .as_ref()?
            .message
            .as_ref()?
            .content
            .as_deref()
            .map(str::trim)
            .filter(|message| !message.is_empty())
    }
}

/// A mail attachment. Listing deliberately omits `contentBytes` — those are
/// fetched separately so a big file isn't pulled just to show its name.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Attachment {
    pub id: String,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub content_type: Option<String>,
    #[serde(default)]
    pub size: Option<i64>,
    #[serde(default)]
    pub is_inline: Option<bool>,
    #[serde(default)]
    pub content_id: Option<String>,
    /// Present when Graph returned the file body (inline image listing).
    #[serde(default)]
    pub content_bytes: Option<String>,
    /// `#microsoft.graph.fileAttachment`, `itemAttachment`, `referenceAttachment`.
    #[serde(default, rename = "@odata.type")]
    pub odata_type: Option<String>,
}

impl Attachment {
    pub fn display_name(&self) -> String {
        self.name.clone().unwrap_or_else(|| "(unnamed)".into())
    }

    /// Only file attachments have bytes to download.
    pub fn is_file(&self) -> bool {
        self.odata_type
            .as_deref()
            .map(|t| t.ends_with("fileAttachment"))
            .unwrap_or(true)
    }

    pub fn human_size(&self) -> String {
        match self.size {
            Some(b) if b >= 1024 * 1024 => format!("{:.1} MB", b as f64 / (1024.0 * 1024.0)),
            Some(b) if b >= 1024 => format!("{:.0} KB", b as f64 / 1024.0),
            Some(b) => format!("{b} B"),
            None => String::new(),
        }
    }
}

/// Something attached to a Teams message: a shared file, or — for a reply —
/// a `messageReference` pointing at the message being answered.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MessageAttachment {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub content_url: Option<String>,
    #[serde(default)]
    pub content_type: Option<String>,
    /// For `messageReference`, a JSON *string* describing the quoted message.
    #[serde(default)]
    pub content: Option<String>,
}

/// The message a reply is answering.
#[derive(Debug, Clone)]
pub struct QuotedMessage {
    pub message_id: String,
    pub author: String,
    pub preview: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MessageReference {
    #[serde(default)]
    message_id: Option<String>,
    #[serde(default)]
    message_preview: Option<String>,
    #[serde(default)]
    message_sender: Option<IdentitySet>,
}

/// Outbound draft used by the compose views for both mail and Teams messages.
#[derive(Debug, Clone, Default, Serialize)]
pub struct Draft {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
}

impl ChatMessage {
    /// Whether this message has been edited after sending.
    pub fn was_edited(&self) -> bool {
        self.last_edited_date_time
            .as_deref()
            .is_some_and(|s| !s.is_empty())
    }

    /// The signed-in user can edit their own, still-present chat messages.
    pub fn can_edit(&self, my_user_id: Option<&str>) -> bool {
        let Some(me) = my_user_id.filter(|id| !id.is_empty()) else {
            return false;
        };
        self.deleted_date_time.is_none()
            && self.author_id() == Some(me)
            && self.message_type.as_deref().unwrap_or("message") == "message"
    }

    /// Body text for the composer: HTML tags stripped, `<br>`/`</p>` as newlines.
    pub fn editable_text(&self) -> String {
        html_to_editable(&self.text())
    }
}
/// Convert a Teams HTML body into the plain text a user would retype.
fn html_to_editable(html: &str) -> String {
    let chars: Vec<char> = html.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '<' => {
                i += 1;
                let tag_start = i;
                while i < chars.len() && chars[i] != '>' {
                    i += 1;
                }
                let tag: String = chars[tag_start..i].iter().collect();
                if i < chars.len() {
                    i += 1;
                }
                let trimmed = tag.trim();
                let is_close = trimmed.starts_with('/');
                let name = trimmed
                    .trim_start_matches('/')
                    .split(|c: char| c.is_whitespace() || c == '/')
                    .next()
                    .unwrap_or("")
                    .to_ascii_lowercase();
                let block_end = is_close
                    && matches!(
                        name.as_str(),
                        "p" | "div"
                            | "li"
                            | "tr"
                            | "h1"
                            | "h2"
                            | "h3"
                            | "h4"
                            | "h5"
                            | "h6"
                            | "blockquote"
                    );
                if (name == "br" || block_end) && !out.ends_with('\n') {
                    out.push('\n');
                }
            }
            '&' => {
                let rest: String = chars[i..].iter().take(10).collect();
                let (ch, skip) = if rest.starts_with("&amp;") {
                    ('&', 5)
                } else if rest.starts_with("&lt;") {
                    ('<', 4)
                } else if rest.starts_with("&gt;") {
                    ('>', 4)
                } else if rest.starts_with("&quot;") {
                    ('"', 6)
                } else if rest.starts_with("&#39;") {
                    ('\'', 5)
                } else if rest.starts_with("&apos;") {
                    ('\'', 6)
                } else if rest.starts_with("&nbsp;") {
                    (' ', 6)
                } else {
                    ('&', 1)
                };
                out.push(ch);
                i += skip;
            }
            _ => {
                out.push(chars[i]);
                i += 1;
            }
        }
    }
    out.lines()
        .map(str::trim_end)
        .collect::<Vec<_>>()
        .join("\n")
        .trim()
        .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn one_on_one_label_uses_last_message_sender_for_federated_peer() {
        let chat: Chat = serde_json::from_value(serde_json::json!({
            "id": "chat-1",
            "chatType": "oneOnOne",
            "members": [
                {
                    "userId": "me",
                    "displayName": "Local User",
                    "email": "me@example.com"
                },
                {
                    "userId": "external-user"
                }
            ],
            "lastMessagePreview": {
                "id": "message-1",
                "from": {
                    "user": {
                        "id": "external-user",
                        "displayName": "External Person"
                    }
                }
            }
        }))
        .unwrap();

        assert_eq!(chat.label(Some("me")), "External Person");
    }

    #[test]
    fn one_on_one_label_does_not_use_own_preview_name() {
        let chat: Chat = serde_json::from_value(serde_json::json!({
            "id": "chat-2",
            "chatType": "oneOnOne",
            "members": [
                {
                    "userId": "me",
                    "displayName": "Local User",
                    "email": "me@example.com"
                },
                {
                    "userId": "external-user",
                    "email": "external@example.net"
                }
            ],
            "lastMessagePreview": {
                "id": "message-2",
                "from": {
                    "user": {
                        "id": "me",
                        "displayName": "Local User"
                    }
                }
            }
        }))
        .unwrap();

        assert_eq!(chat.label(Some("me")), "external@example.net");
    }

    #[test]
    fn one_on_one_label_keeps_chat_type_as_final_fallback() {
        let chat: Chat = serde_json::from_value(serde_json::json!({
            "id": "chat-3",
            "chatType": "oneOnOne",
            "members": [
                {
                    "userId": "me",
                    "displayName": "Local User"
                },
                {
                    "userId": "external-user"
                }
            ]
        }))
        .unwrap();

        assert_eq!(chat.label(Some("me")), "oneOnOne");
    }

    /// The exact shape Graph returns for a reply in a chat.
    fn reply_message() -> ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": "1785859178276",
            "createdDateTime": "2026-08-04T15:59:38.276Z",
            "body": {
                "contentType": "html",
                "content": "<attachment id=\"1785858892876\"></attachment><p>Confirma por favor</p>"
            },
            "attachments": [{
                "id": "1785858892876",
                "contentType": "messageReference",
                "content": "{\"messageId\":\"1785858892876\",\"messagePreview\":\"Sounds good to me\",\"messageSender\":{\"user\":{\"userIdentityType\":\"aadUser\",\"id\":\"abc\",\"displayName\":\"Alex Rivera\"}}}"
            }]
        }))
        .unwrap()
    }

    #[test]
    fn reads_the_quoted_message_from_a_reference_attachment() {
        let q = reply_message()
            .quoted()
            .expect("reply should carry a quote");
        assert_eq!(q.author, "Alex Rivera");
        assert_eq!(q.preview, "Sounds good to me");
        assert_eq!(q.message_id, "1785858892876");
    }

    #[test]
    fn an_ordinary_message_has_no_quote() {
        let plain: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "1",
            "body": { "contentType": "text", "content": "hello" }
        }))
        .unwrap();
        assert!(plain.quoted().is_none());
    }

    #[test]
    fn humanizes_useful_system_events() {
        let message: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-1",
            "messageType": "systemEventMessage",
            "body": { "contentType": "html", "content": "<systemEventMessage/>" },
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callRecordingEventMessageDetail",
                "callRecordingDisplayName": "Weekly sync.mp4",
                "callRecordingUrl": "https://example.invalid/recording"
            }
        }))
        .unwrap();

        let display = message.system_event_display().unwrap();
        assert_eq!(display.class, SystemEventClass::Useful);
        assert!(display.text.contains("Weekly sync.mp4"));
        assert_eq!(
            display.url.as_deref(),
            Some("https://example.invalid/recording")
        );
    }

    #[test]
    fn classifies_known_noise_and_unknown_system_events() {
        let noise: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-2",
            "messageType": "systemEventMessage",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callStartedEventMessageDetail"
            }
        }))
        .unwrap();
        assert_eq!(
            noise.system_event_display().unwrap().class,
            SystemEventClass::Noise
        );

        let unknown: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-3",
            "messageType": "systemEventMessage",
            "eventDetail": {
                "@odata.type": "#microsoft.graph.futureImportantEventMessageDetail"
            }
        }))
        .unwrap();
        let display = unknown.system_event_display().unwrap();
        assert_eq!(display.class, SystemEventClass::Unknown);
        assert!(display.text.contains("Future Important"));
    }

    #[test]
    fn old_cached_system_event_without_detail_is_unknown() {
        let message: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-old",
            "messageType": "unknownFutureValue",
            "body": {
                "contentType": "html",
                "content": "<systemEventMessage/>"
            }
        }))
        .unwrap();

        assert!(message.is_system_event());
        assert_eq!(
            message.system_event_display().unwrap().class,
            SystemEventClass::Unknown
        );
    }

    #[test]
    fn event_detail_identifies_system_event_even_with_unknown_message_type() {
        let message: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-future",
            "messageType": "unknownFutureValue",
            "body": { "contentType": "html", "content": "" },
            "eventDetail": {
                "@odata.type": "#microsoft.graph.callStartedEventMessageDetail"
            }
        }))
        .unwrap();

        assert!(message.is_system_event());
        assert_eq!(
            message.system_event_display().unwrap().class,
            SystemEventClass::Noise
        );
    }

    #[test]
    fn chat_message_round_trips_for_persistent_cache() {
        let original = reply_message();
        let encoded = serde_json::to_vec(&original).expect("serialize chat message");
        let restored: ChatMessage =
            serde_json::from_slice(&encoded).expect("deserialize chat message");

        assert_eq!(restored.id, original.id);
        assert_eq!(restored.text(), original.text());
        assert_eq!(
            restored.quoted().map(|quote| quote.message_id),
            original.quoted().map(|quote| quote.message_id)
        );
    }

    #[test]
    fn preview_flattens_html_and_bounds_length() {
        let m: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "1",
            "body": { "contentType": "html", "content": "<p>one   two</p>\n<p>three</p>" }
        }))
        .unwrap();
        assert_eq!(m.text_preview(100), "one two three");
        assert_eq!(m.text_preview(5).chars().count(), 6); // 5 + the ellipsis
    }
    #[test]
    fn editable_text_keeps_line_breaks_and_drops_quote_markup() {
        let m = reply_message();
        assert_eq!(m.editable_text(), "Confirma por favor");

        let html: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "1",
            "body": {
                "contentType": "html",
                "content": "<p>one</p><p>two &amp; three<br>four</p><img alt=\"x\" src=\"../hostedContents/1/$value\">"
            }
        }))
        .unwrap();
        assert_eq!(html.editable_text(), "one\ntwo & three\nfour");
    }
    #[test]
    fn only_the_author_can_edit_a_live_message() {
        let mut m = reply_message();
        m.from = Some(IdentitySet {
            user: Some(Identity {
                id: Some("me".into()),
                display_name: Some("Me".into()),
                user_identity_type: None,
                tenant_id: None,
            }),
            application: None,
        });
        assert!(m.can_edit(Some("me")));
        assert!(!m.can_edit(Some("other")));
        assert!(!m.can_edit(None));
        m.deleted_date_time = Some("2026-08-04T16:00:00Z".into());
        assert!(!m.can_edit(Some("me")));
    }
    #[test]
    fn edited_flag_follows_last_edited_date_time() {
        let mut m = reply_message();
        assert!(!m.was_edited());
        m.last_edited_date_time = Some("2026-08-04T16:00:00Z".into());
        assert!(m.was_edited());
    }
}
