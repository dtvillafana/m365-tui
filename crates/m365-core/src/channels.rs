//! Teams team/channel endpoints.

use anyhow::Result;
use serde_json::json;

use crate::graph::GraphClient;
use crate::hosted::{outgoing_payload, OutgoingBody};
use crate::models::{Channel, ChatMessage, Team};

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
    body: OutgoingBody<'_>,
) -> Result<ChatMessage> {
    graph
        .post_json(
            &format!("teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"),
            &outgoing_payload(body),
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

/// Post a message to a channel — plain text, or HTML with inline images.
pub async fn send_message(
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
