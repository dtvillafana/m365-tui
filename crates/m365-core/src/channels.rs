//! Teams team/channel endpoints.

use crate::hosted::{outgoing_payload, update_payload, OutgoingBody};
use anyhow::Result;
use serde_json::json;

use crate::graph::GraphClient;
use crate::models::{Channel, ChatMessage, Team};
use crate::teams_format::outgoing_body;

/// Teams the signed-in user has joined.
pub async fn joined_teams(graph: &GraphClient) -> Result<Vec<Team>> {
    graph
        .get_collection("me/joinedTeams?$select=id,displayName,description")
        .await
}

/// Channels within a team.
pub async fn list_channels(graph: &GraphClient, team_id: &str) -> Result<Vec<Channel>> {
    graph
        .get_collection(&format!(
            "teams/{team_id}/channels?$select=id,displayName,description"
        ))
        .await
}

/// Top-level messages in a channel (replies are fetched separately by Graph).
/// Also returns the `@odata.nextLink` for fetching older messages.
pub async fn list_messages(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    top: u32,
) -> Result<(Vec<ChatMessage>, Option<String>)> {
    let path = format!("teams/{team_id}/channels/{channel_id}/messages?$top={top}");
    graph.get_page_with_next(&path).await
}

/// Fetch the next (older) page from an `@odata.nextLink`.
pub async fn list_messages_more(
    graph: &GraphClient,
    next_link: &str,
) -> Result<(Vec<ChatMessage>, Option<String>)> {
    graph.get_page_with_next(next_link).await
}

/// Reply to a channel message. Channels have a real replies collection, unlike
/// chats, so this threads properly.
pub async fn send_reply(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    text: &str,
    markdown: bool,
) -> Result<ChatMessage> {
    let body = outgoing_body(text, markdown);
    let payload = json!({
        "body": {
            "contentType": body.content_type,
            "content": body.content,
        }
    });
    graph
        .post_json(
            &format!("teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"),
            &payload,
        )
        .await
}

/// React to a channel message with an emoji (unicode, e.g. "👍").
pub async fn set_reaction(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    emoji: &str,
) -> Result<()> {
    graph
        .post_action(
            &format!("teams/{team_id}/channels/{channel_id}/messages/{message_id}/setReaction"),
            &json!({ "reactionType": emoji }),
        )
        .await
}

/// Post a plain-text message to a channel.
pub async fn send_message(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    text: &str,
    markdown: bool,
) -> Result<ChatMessage> {
    let body = outgoing_body(text, markdown);
    let payload = json!({
        "body": {
            "contentType": body.content_type,
            "content": body.content,
        }
    });
    graph
        .post_json(
            &format!("teams/{team_id}/channels/{channel_id}/messages"),
            &payload,
        )
        .await
}

/// Download one Teams-hosted inline content item from a channel message.
pub async fn hosted_content_bytes(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    hosted_content_id: &str,
) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!(
            "teams/{team_id}/channels/{channel_id}/messages/{message_id}/hostedContents/{hosted_content_id}/$value"
        ))
        .await
}

// Hosted-content variants retain support for inline image composition.
pub async fn send_message_with_body(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    body: OutgoingBody<'_>,
) -> Result<ChatMessage> {
    graph
        .post_json(
            &format!("teams/{team_id}/channels/{channel_id}/messages"),
            &outgoing_payload(body),
        )
        .await
}
pub async fn send_reply_with_body(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    body: OutgoingBody<'_>,
) -> Result<ChatMessage> {
    graph
        .post_json(
            &format!("teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"),
            &outgoing_payload(body),
        )
        .await
}
pub async fn update_message_with_body(
    graph: &GraphClient,
    team_id: &str,
    channel_id: &str,
    original: &ChatMessage,
    body: OutgoingBody<'_>,
) -> Result<()> {
    let path = match original.reply_to_id.as_deref().filter(|s| !s.is_empty()) {
        Some(parent) => format!(
            "teams/{team_id}/channels/{channel_id}/messages/{parent}/replies/{}",
            original.id
        ),
        None => format!(
            "teams/{team_id}/channels/{channel_id}/messages/{}",
            original.id
        ),
    };
    graph.patch(&path, &update_payload(original, body)).await
}
