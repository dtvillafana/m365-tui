//! Teams 1:1 / group chat endpoints.

use anyhow::Result;
use serde_json::json;

use crate::graph::{DeltaPage, GraphClient};
use crate::hosted::{images_of, outgoing_payload, with_hosted_contents, OutgoingBody};
use crate::models::{Chat, ChatMessage};
use crate::util::html_escape;

/// List the signed-in user's chats, most-recently-updated first, with member
/// names and a last-message preview expanded for display.
pub async fn list_chats(graph: &GraphClient, top: u32) -> Result<Vec<Chat>> {
    let path = format!(
        "me/chats?$top={top}&$orderby=lastMessagePreview/createdDateTime desc\
         &$expand=members,lastMessagePreview"
    );
    graph.get_page(&path).await
}

/// List the first page of messages in a chat, newest first. Also returns the
/// `@odata.nextLink` for fetching older messages, if there are any.
pub async fn list_messages(
    graph: &GraphClient,
    chat_id: &str,
    top: u32,
) -> Result<(Vec<ChatMessage>, Option<String>)> {
    let path = format!("me/chats/{chat_id}/messages?$top={top}");
    graph.get_page_with_next(&path).await
}

/// Fetch the next (older) page from an `@odata.nextLink`.
pub async fn list_messages_more(
    graph: &GraphClient,
    next_link: &str,
) -> Result<(Vec<ChatMessage>, Option<String>)> {
    graph.get_page_with_next(next_link).await
}

/// Incremental sync of a chat's messages.
pub async fn delta_messages(
    graph: &GraphClient,
    chat_id: &str,
    delta_link: Option<&str>,
) -> Result<DeltaPage<ChatMessage>> {
    let path = match delta_link {
        Some(link) => link.to_string(),
        None => format!("me/chats/{chat_id}/messages/delta"),
    };
    graph.delta(&path).await
}

/// Send a message to a chat — plain text, or HTML with inline images.
pub async fn send_message(
    graph: &GraphClient,
    chat_id: &str,
    body: OutgoingBody<'_>,
) -> Result<ChatMessage> {
    graph
        .post_json(
            &format!("me/chats/{chat_id}/messages"),
            &outgoing_payload(body),
        )
        .await
}

/// Reply to a message in a chat.
///
/// Chats have no replies endpoint. Teams represents a reply as a
/// `messageReference` attachment plus an empty `<attachment>` tag in the body —
/// exactly what `ChatMessage::quoted` reads back — so that shape is what gets
/// posted here.
pub async fn send_reply(
    graph: &GraphClient,
    chat_id: &str,
    original: &ChatMessage,
    body: OutgoingBody<'_>,
) -> Result<ChatMessage> {
    let message_id = &original.id;
    let preview: String = original.text_preview(250);
    let inner = body.inner_html();
    let images = images_of(body);
    let sender = json!({
        "user": {
            "userIdentityType": "aadUser",
            "id": original.author_id().unwrap_or_default(),
            "displayName": original.author(),
        }
    });
    let reference = json!({
        "messageId": message_id,
        "messagePreview": preview,
        "messageSender": sender,
    })
    .to_string();

    let payload = with_hosted_contents(
        json!({
            "body": {
                "contentType": "html",
                "content": format!(
                    "<attachment id=\"{message_id}\"></attachment><p>{inner}</p>",
                ),
            },
            "attachments": [{
                "id": message_id,
                "contentType": "messageReference",
                "content": reference,
            }],
        }),
        images,
    );

    match graph
        .post_json(&format!("me/chats/{chat_id}/messages"), &payload)
        .await
    {
        Ok(m) => Ok(m),
        // If the tenant rejects the reference attachment, still deliver the
        // message rather than losing what the user typed.
        Err(e) => {
            tracing::warn!("native reply rejected, sending as a quote instead: {e:#}");
            let quoted = format!(
                "<blockquote><b>{}</b><br>{}</blockquote><p>{inner}</p>",
                html_escape(&original.author()),
                html_escape(&preview),
            );
            graph
                .post_json(
                    &format!("me/chats/{chat_id}/messages"),
                    &outgoing_payload(OutgoingBody::Html {
                        html: &quoted,
                        images,
                    }),
                )
                .await
        }
    }
}

/// React to a chat message with an emoji (unicode, e.g. "👍").
pub async fn set_reaction(
    graph: &GraphClient,
    chat_id: &str,
    message_id: &str,
    emoji: &str,
) -> Result<()> {
    graph
        .post_action(
            &format!("chats/{chat_id}/messages/{message_id}/setReaction"),
            &json!({ "reactionType": emoji }),
        )
        .await
}

/// Create (or return existing) 1:1 chat with another user by their id.
pub async fn create_one_on_one(
    graph: &GraphClient,
    my_user_id: &str,
    other_user_id: &str,
) -> Result<Chat> {
    let member = |uid: &str| {
        json!({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": format!("https://graph.microsoft.com/v1.0/users('{uid}')"),
        })
    };
    let payload = json!({
        "chatType": "oneOnOne",
        "members": [member(my_user_id), member(other_user_id)],
    });
    graph.post_json("chats", &payload).await
}
