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

#[derive(Debug, Clone, Deserialize)]
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
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MailMessage {
    pub id: String,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub body_preview: Option<String>,
    #[serde(default)]
    pub from: Option<Recipient>,
    #[serde(default)]
    pub to_recipients: Vec<Recipient>,
    #[serde(default)]
    pub received_date_time: Option<String>,
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DateTimeTimeZone {
    pub date_time: String,
    #[serde(default)]
    pub time_zone: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Attendee {
    #[serde(default)]
    pub email_address: Option<EmailAddress>,
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
    pub last_updated_date_time: Option<String>,
    #[serde(default)]
    pub members: Vec<ConversationMember>,
    #[serde(default)]
    pub last_message_preview: Option<LastMessagePreview>,
    #[serde(default)]
    pub online_meeting_info: Option<ChatOnlineMeetingInfo>,
}

/// Join information on a meeting chat (`chatType` is `meeting`).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChatOnlineMeetingInfo {
    #[serde(default)]
    pub join_web_url: Option<String>,
}

/// Who a chat's `c` (call) key should ring, or which meeting to join.
#[derive(Debug, Clone)]
pub enum ChatCall {
    Meeting { join_url: String, label: String },
    Users { ids: Vec<String>, label: String },
}

impl Chat {
    /// A display label: the explicit topic, else the member names joined.
    pub fn label(&self, me_id: Option<&str>) -> String {
        if let Some(t) = self.topic.as_ref().filter(|t| !t.is_empty()) {
            return t.clone();
        }
        let names: Vec<String> = self
            .members
            .iter()
            .filter(|m| {
                me_id
                    .map(|id| m.user_id.as_deref() != Some(id))
                    .unwrap_or(true)
            })
            .filter_map(|m| m.display_name.clone())
            .collect();
        if names.is_empty() {
            self.chat_type.clone().unwrap_or_else(|| "chat".into())
        } else {
            names.join(", ")
        }
    }

    /// A call target for this chat: a meeting join URL if it is a meeting
    /// chat, otherwise the other members' Entra ids.
    pub fn call_target(&self, me_id: Option<&str>) -> Option<ChatCall> {
        let label = self.label(me_id);
        if let Some(url) = self
            .online_meeting_info
            .as_ref()
            .and_then(|o| o.join_web_url.as_deref())
            .filter(|s| !s.is_empty())
        {
            return Some(ChatCall::Meeting {
                join_url: url.to_string(),
                label,
            });
        }
        let ids: Vec<String> = self
            .members
            .iter()
            .filter(|m| {
                me_id
                    .map(|id| m.user_id.as_deref() != Some(id))
                    .unwrap_or(true)
            })
            .filter_map(|m| m.user_id.clone())
            .filter(|id| !id.is_empty())
            .collect();
        if ids.is_empty() {
            None
        } else {
            Some(ChatCall::Users { ids, label })
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConversationMember {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub display_name: Option<String>,
    #[serde(default)]
    pub user_id: Option<String>,
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
#[derive(Debug, Clone, Deserialize)]
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct IdentitySet {
    #[serde(default)]
    pub user: Option<Identity>,
    #[serde(default)]
    pub application: Option<Identity>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Identity {
    #[serde(default)]
    pub id: Option<String>,
    #[serde(default)]
    pub display_name: Option<String>,
}

/// A Teams `chatMessage` (channel or chat).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChatMessage {
    pub id: String,
    #[serde(default)]
    pub created_date_time: Option<String>,
    #[serde(default)]
    pub from: Option<IdentitySet>,
    #[serde(default)]
    pub body: Option<ItemBody>,
    #[serde(default)]
    pub message_type: Option<String>,
    #[serde(default)]
    pub deleted_date_time: Option<String>,
    #[serde(default)]
    pub reactions: Vec<MessageReaction>,
    #[serde(default)]
    pub attachments: Vec<MessageAttachment>,
    #[serde(default)]
    pub mentions: Vec<ChatMessageMention>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MessageReaction {
    #[serde(default)]
    pub reaction_type: Option<String>,
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
    #[serde(default, rename = "scoredEmailAddresses")]
    pub scored_email_addresses: Vec<ScoredEmailAddress>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ScoredEmailAddress {
    #[serde(default)]
    pub address: Option<String>,
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
#[derive(Debug, Clone, Deserialize)]
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

#[cfg(test)]
mod tests {
    use super::*;

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
    fn call_target_prefers_a_meeting_join_url() {
        let chat: Chat = serde_json::from_value(serde_json::json!({
            "id": "19:meeting",
            "topic": "Standup",
            "chatType": "meeting",
            "onlineMeetingInfo": {
                "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc"
            },
            "members": [{ "userId": "me", "displayName": "Me" }]
        }))
        .unwrap();
        match chat.call_target(Some("me")) {
            Some(ChatCall::Meeting { join_url, label }) => {
                assert!(join_url.contains("meetup-join"));
                assert_eq!(label, "Standup");
            }
            other => panic!("expected a meeting, got {other:?}"),
        }
    }

    #[test]
    fn call_target_on_a_one_to_one_chat_is_the_other_person() {
        let chat: Chat = serde_json::from_value(serde_json::json!({
            "id": "19:chat",
            "chatType": "oneOnOne",
            "members": [
                { "userId": "me", "displayName": "Me" },
                { "userId": "them", "displayName": "Alex Rivera" }
            ]
        }))
        .unwrap();
        match chat.call_target(Some("me")) {
            Some(ChatCall::Users { ids, label }) => {
                assert_eq!(ids, vec!["them".to_string()]);
                assert_eq!(label, "Alex Rivera");
            }
            other => panic!("expected users, got {other:?}"),
        }
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
}
