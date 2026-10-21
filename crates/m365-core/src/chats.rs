//! Teams 1:1 / group chat endpoints.

use crate::hosted::{
    images_of, outgoing_payload, update_payload, with_hosted_contents, OutgoingBody,
};
use anyhow::Result;
use serde_json::json;

use crate::graph::{DeltaPage, GraphClient};
use crate::models::{Chat, ChatMessage, ConversationMember};
use crate::teams_format::outgoing_body;
use crate::util::base64_url_no_pad;
use crate::util::html_escape;

/// List the signed-in user's chats, most-recently-updated first, with member
/// names and a last-message preview expanded for display.
pub async fn list_chats(graph: &GraphClient, top: u32) -> Result<Vec<Chat>> {
    let path = format!(
        "me/chats?$top={top}&$orderby=lastMessagePreview/createdDateTime desc\
         &$expand=members,lastMessagePreview"
    );
    let mut chats: Vec<Chat> = graph.get_page(&path).await?;

    // Federated chats can expose an opaque Teams identity (for example a
    // non-GUID id) in the expanded roster/message sender. Presence endpoints,
    // however, require the user's Entra object GUID. Refresh only suspicious
    // 1:1 rosters through the federation-aware /members endpoint.
    for chat in &mut chats {
        let one_on_one = chat
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"));
        let has_non_guid_user = chat
            .members
            .iter()
            .filter_map(|member| member.user_id.as_deref())
            .any(|id| !looks_like_user_guid(id));

        if !one_on_one || !has_non_guid_user {
            continue;
        }

        match list_members(graph, &chat.id).await {
            Ok(members) if !members.is_empty() => {
                tracing::debug!(
                    "refreshed federated chat roster for {}: {} member(s)",
                    chat.id,
                    members.len()
                );
                chat.members = members;
            }
            Ok(_) => {
                tracing::debug!(
                    "federated chat roster refresh returned no members for {}",
                    chat.id
                );
            }
            Err(error) => {
                // Keep the original expanded roster: naming/chat use must not
                // fail merely because optional presence enrichment did.
                tracing::debug!(
                    "federated chat roster refresh failed for {}: {error:#}",
                    chat.id
                );
            }
        }
    }

    Ok(chats)
}

/// List the exact chat roster. This endpoint supports federation and returns
/// aadUserConversationMember fields such as userId and tenantId when available.
pub async fn list_members(graph: &GraphClient, chat_id: &str) -> Result<Vec<ConversationMember>> {
    graph
        .get_collection(&format!("me/chats/{chat_id}/members"))
        .await
}

pub fn looks_like_user_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    if bytes.len() != 36 {
        return false;
    }

    for (index, byte) in bytes.iter().copied().enumerate() {
        match index {
            8 | 13 | 18 | 23 if byte == b'-' => {}
            8 | 13 | 18 | 23 => return false,
            _ if byte.is_ascii_hexdigit() => {}
            _ => return false,
        }
    }

    true
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

/// Mark the signed-in user's chat as read in Teams.
///
/// Graph expects the signed-in user's object id and tenant id.
pub async fn mark_read(
    graph: &GraphClient,
    chat_id: &str,
    user_id: &str,
    tenant_id: &str,
) -> Result<()> {
    graph
        .post_action(
            &format!("chats/{chat_id}/markChatReadForUser"),
            &json!({
                "user": {
                    "id": user_id,
                    "tenantId": tenant_id
                }
            }),
        )
        .await
}

/// Mark a chat unread for the signed-in user. Omitting
/// lastMessageReadDateTime makes Graph mark the last message unread.
pub async fn mark_unread(
    graph: &GraphClient,
    chat_id: &str,
    user_id: &str,
    tenant_id: &str,
) -> Result<()> {
    graph
        .post_action(
            &format!("chats/{chat_id}/markChatUnreadForUser"),
            &json!({
                "user": {
                    "id": user_id,
                    "tenantId": tenant_id
                }
            }),
        )
        .await
}

/// Hide or unhide a chat for the signed-in user.
///
/// This is the server-side Teams archive state surfaced by
/// `chat.viewpoint.isHidden` in the list-chats response.
pub async fn set_hidden(
    graph: &GraphClient,
    chat_id: &str,
    user_id: &str,
    tenant_id: &str,
    hidden: bool,
) -> Result<()> {
    let action = if hidden {
        "hideForUser"
    } else {
        "unhideForUser"
    };

    graph
        .post_action(
            &format!("chats/{chat_id}/{action}"),
            &json!({
                "user": {
                    "id": user_id,
                    "tenantId": tenant_id
                }
            }),
        )
        .await
}

/// List messages newest-first by creation time. Used for unread counting so
/// paging can stop as soon as the chat viewpoint is reached.
pub async fn list_messages_created_desc(
    graph: &GraphClient,
    chat_id: &str,
    top: u32,
) -> Result<(Vec<ChatMessage>, Option<String>)> {
    let path = format!("me/chats/{chat_id}/messages?$top={top}&$orderby=createdDateTime desc");
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

/// Send a message to a chat, optionally applying the safe Markdown subset.
pub async fn send_message(
    graph: &GraphClient,
    chat_id: &str,
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
        .post_json(&format!("me/chats/{chat_id}/messages"), &payload)
        .await
}

/// Change the topic of a group chat.
pub async fn set_topic(graph: &GraphClient, chat_id: &str, topic: &str) -> Result<()> {
    graph
        .patch(&format!("chats/{chat_id}"), &json!({ "topic": topic }))
        .await
}

/// Replace the body of a chat message, optionally applying the safe Markdown subset.
pub async fn edit_message(
    graph: &GraphClient,
    chat_id: &str,
    message_id: &str,
    text: &str,
    markdown: bool,
) -> Result<()> {
    let body = outgoing_body(text, markdown);
    graph
        .patch(
            &format!("chats/{chat_id}/messages/{message_id}"),
            &json!({
                "body": {
                    "contentType": body.content_type,
                    "content": body.content
                }
            }),
        )
        .await
}

/// Soft-delete one of the signed-in user's chat messages.
pub async fn soft_delete_message(
    graph: &GraphClient,
    user_id: &str,
    chat_id: &str,
    message_id: &str,
) -> Result<()> {
    graph
        .post_empty_action(&format!(
            "users/{user_id}/chats/{chat_id}/messages/{message_id}/softDelete"
        ))
        .await
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct PinnedChatMessageInfo {
    id: String,
}

/// List pin IDs first, then fetch each chatMessage separately.
///
/// This deliberately avoids `$expand=message`: some Teams thread shapes return
/// `Thread property pinnedItems was not found` through that expanded path even
/// though Graph documents the plain pinnedMessages collection separately.
pub async fn list_pinned_messages(graph: &GraphClient, chat_id: &str) -> Result<Vec<ChatMessage>> {
    let pins: Vec<PinnedChatMessageInfo> = graph
        .get_collection(&format!("chats/{chat_id}/pinnedMessages"))
        .await?;

    let mut messages = Vec::with_capacity(pins.len());
    for pin in pins {
        match graph
            .get_json::<ChatMessage>(&format!("chats/{chat_id}/messages/{}", pin.id))
            .await
        {
            Ok(message) => messages.push(message),
            Err(error) => {
                tracing::warn!(
                    "could not load pinned Teams message {} in {}: {error:#}",
                    pin.id,
                    chat_id
                );
            }
        }
    }
    Ok(messages)
}

/// Pin an existing message in a chat.
pub async fn pin_message(graph: &GraphClient, chat_id: &str, message_id: &str) -> Result<()> {
    graph
        .post_action(
            &format!("chats/{chat_id}/pinnedMessages"),
            &json!({
                "message@odata.bind": format!(
                    "https://graph.microsoft.com/v1.0/chats/{chat_id}/messages/{message_id}"
                )
            }),
        )
        .await
}

/// Unpin a message. Graph defines pinnedChatMessageInfo.id as the chatMessage id.
pub async fn unpin_message(graph: &GraphClient, chat_id: &str, message_id: &str) -> Result<()> {
    graph
        .delete(&format!("chats/{chat_id}/pinnedMessages/{message_id}"))
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
    text: &str,
    markdown: bool,
) -> Result<ChatMessage> {
    let message_id = &original.id;
    let formatted = outgoing_body(text, markdown);
    let formatted_html = formatted.html_fragment();
    let preview: String = original.text_preview(250);
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

    let payload = json!({
        "body": {
            "contentType": "html",
            "content": format!(
                "<attachment id=\"{message_id}\"></attachment>{formatted_html}"
            ),
        },
        "attachments": [{
            "id": message_id,
            "contentType": "messageReference",
            "content": reference,
        }],
    });

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
                "<blockquote><b>{}</b><br>{}</blockquote>{formatted_html}",
                html_escape(&original.author()),
                html_escape(&preview),
            );
            graph
                .post_json(
                    &format!("me/chats/{chat_id}/messages"),
                    &json!({ "body": { "contentType": "html", "content": quoted } }),
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

/// Download one Teams-hosted inline content item from a chat message.
pub async fn hosted_content_bytes(
    graph: &GraphClient,
    chat_id: &str,
    message_id: &str,
    hosted_content_id: &str,
) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!(
            "chats/{chat_id}/messages/{message_id}/hostedContents/{hosted_content_id}/$value"
        ))
        .await
}

/// Extract a path relative to a OneDrive for Business drive root from a
/// `driveItem.webDavUrl`. The web URL contains the document-library component
/// (`/Documents/`), while Graph's `/users/{id}/drive/root:` already points at
/// that library root.
fn personal_webdav_item_path(content_url: &str) -> Option<String> {
    let no_fragment = content_url.split('#').next()?;
    let no_query = no_fragment.split('?').next()?;
    let scheme_end = no_query.find("://")? + 3;
    let path_start = no_query[scheme_end..].find('/')? + scheme_end;
    let path = &no_query[path_start..];

    if !path.contains("/personal/") {
        return None;
    }

    let marker = "/Documents/";
    let marker_pos = path.find(marker)?;
    let item_path = &path[marker_pos + marker.len()..];
    (!item_path.is_empty()).then(|| item_path.to_string())
}

/// Download a real SharePoint/OneDrive sharing link through Graph's `/shares`
/// endpoint.
pub async fn shared_file_bytes(graph: &GraphClient, content_url: &str) -> Result<Vec<u8>> {
    let encoded = format!("u!{}", base64_url_no_pad(content_url.as_bytes()));
    graph
        .get_bytes(&format!("shares/{encoded}/driveItem/content"))
        .await
}

/// Download a Teams `reference` attachment.
///
/// Teams normally emits `driveItem.webDavUrl` for ordinary file attachments.
/// In a 1:1/group chat that URL points into the sender's OneDrive. True sharing
/// links are still handled by `/shares` as a fallback.
#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DriveSearchPage {
    #[serde(default)]
    value: Vec<DriveSearchItem>,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DriveSearchItem {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    remote_item: Option<RemoteDriveItem>,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct RemoteDriveItem {
    id: String,
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    web_dav_url: Option<String>,
    #[serde(default)]
    web_url: Option<String>,
    #[serde(default)]
    parent_reference: Option<DriveParentReference>,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DriveParentReference {
    #[serde(default)]
    drive_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RemoteDriveTarget {
    drive_id: String,
    item_id: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum ReferenceRoute {
    SenderDrive,
    AccessibleDriveSearch,
    MicrosoftSearch,
    Shares,
}

impl ReferenceRoute {
    fn label(self) -> &'static str {
        match self {
            Self::SenderDrive => "sender-drive",
            Self::AccessibleDriveSearch => "accessible-drive-search",
            Self::MicrosoftSearch => "microsoft-search",
            Self::Shares => "shares",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum ReferenceRouteClass {
    ShareLink,
    PersonalWebDav,
    Other,
}

impl ReferenceRouteClass {
    fn label(self) -> &'static str {
        match self {
            Self::ShareLink => "share-link",
            Self::PersonalWebDav => "personal-webdav",
            Self::Other => "other-reference",
        }
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct ReferenceRouteStats {
    successes: u32,
    failures: u32,
}

type ReferenceRouteStatsKey = (ReferenceRouteClass, String, ReferenceRoute);

fn reference_route_stats(
) -> &'static std::sync::Mutex<std::collections::HashMap<ReferenceRouteStatsKey, ReferenceRouteStats>>
{
    static STATS: std::sync::OnceLock<
        std::sync::Mutex<std::collections::HashMap<ReferenceRouteStatsKey, ReferenceRouteStats>>,
    > = std::sync::OnceLock::new();

    STATS.get_or_init(|| std::sync::Mutex::new(std::collections::HashMap::new()))
}

fn comparable_reference_url(url: &str) -> String {
    url.split('#')
        .next()
        .unwrap_or(url)
        .split('?')
        .next()
        .unwrap_or(url)
        .trim_end_matches('/')
        .to_ascii_lowercase()
}

fn reference_url_host(content_url: &str) -> String {
    let no_fragment = content_url.split('#').next().unwrap_or(content_url);
    let no_query = no_fragment.split('?').next().unwrap_or(no_fragment);
    let Some(scheme_end) = no_query.find("://").map(|pos| pos + 3) else {
        return String::new();
    };
    no_query[scheme_end..]
        .split('/')
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase()
}

fn looks_like_share_link(content_url: &str) -> bool {
    let no_fragment = content_url.split('#').next().unwrap_or(content_url);
    let no_query = no_fragment.split('?').next().unwrap_or(no_fragment);
    let Some(scheme_end) = no_query.find("://").map(|pos| pos + 3) else {
        return false;
    };
    let Some(path_offset) = no_query[scheme_end..].find('/') else {
        return false;
    };
    let path = &no_query[scheme_end + path_offset..];

    path.starts_with("/:") && path[2..].contains(":/")
}

fn reference_route_class(content_url: &str) -> ReferenceRouteClass {
    if looks_like_share_link(content_url) {
        ReferenceRouteClass::ShareLink
    } else if personal_webdav_item_path(content_url).is_some() {
        ReferenceRouteClass::PersonalWebDav
    } else {
        ReferenceRouteClass::Other
    }
}

fn base_reference_routes(class: ReferenceRouteClass) -> [ReferenceRoute; 4] {
    match class {
        ReferenceRouteClass::ShareLink => [
            ReferenceRoute::Shares,
            ReferenceRoute::MicrosoftSearch,
            ReferenceRoute::AccessibleDriveSearch,
            ReferenceRoute::SenderDrive,
        ],
        ReferenceRouteClass::PersonalWebDav => [
            ReferenceRoute::SenderDrive,
            ReferenceRoute::AccessibleDriveSearch,
            ReferenceRoute::MicrosoftSearch,
            ReferenceRoute::Shares,
        ],
        ReferenceRouteClass::Other => [
            ReferenceRoute::AccessibleDriveSearch,
            ReferenceRoute::MicrosoftSearch,
            ReferenceRoute::Shares,
            ReferenceRoute::SenderDrive,
        ],
    }
}

fn promoted_route_score(stats: ReferenceRouteStats) -> Option<i64> {
    let score = stats.successes as i64 * 2 - stats.failures as i64;
    (stats.successes >= 2 && score > 0).then_some(score)
}

fn ordered_reference_routes(content_url: &str) -> Vec<ReferenceRoute> {
    let class = reference_route_class(content_url);
    let host = reference_url_host(content_url);
    let base = base_reference_routes(class);

    let stats = reference_route_stats()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

    let mut routes: Vec<(usize, ReferenceRoute, Option<i64>)> = base
        .into_iter()
        .enumerate()
        .map(|(index, route)| {
            let route_stats = stats
                .get(&(class, host.clone(), route))
                .copied()
                .unwrap_or_default();
            (index, route, promoted_route_score(route_stats))
        })
        .collect();

    routes.sort_by(|a, b| match (a.2, b.2) {
        (Some(a_score), Some(b_score)) => b_score.cmp(&a_score).then_with(|| a.0.cmp(&b.0)),
        (Some(_), None) => std::cmp::Ordering::Less,
        (None, Some(_)) => std::cmp::Ordering::Greater,
        (None, None) => a.0.cmp(&b.0),
    });

    routes.into_iter().map(|(_, route, _)| route).collect()
}

fn record_reference_route_result(content_url: &str, route: ReferenceRoute, success: bool) {
    let class = reference_route_class(content_url);
    let host = reference_url_host(content_url);
    let mut stats = reference_route_stats()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let entry = stats.entry((class, host, route)).or_default();

    if success {
        entry.successes = entry.successes.saturating_add(1);
    } else {
        entry.failures = entry.failures.saturating_add(1);
    }
}

fn graph_search_term(input: &str) -> String {
    let escaped = input.replace('\'', "''");
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut out = String::with_capacity(escaped.len());

    for byte in escaped.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.' | b'~') {
            out.push(byte as char);
        } else {
            out.push('%');
            out.push(HEX[(byte >> 4) as usize] as char);
            out.push(HEX[(byte & 0x0f) as usize] as char);
        }
    }

    out
}

fn push_unique_target(targets: &mut Vec<RemoteDriveTarget>, target: RemoteDriveTarget) {
    if !targets.contains(&target) {
        targets.push(target);
    }
}

fn remote_drive_target_from_search(
    page: &DriveSearchPage,
    name: &str,
    content_url: &str,
) -> Option<RemoteDriveTarget> {
    let wanted_url = comparable_reference_url(content_url);
    let mut url_matches = Vec::new();
    let mut name_matches = Vec::new();

    for item in &page.value {
        let Some(remote) = item.remote_item.as_ref() else {
            continue;
        };
        let Some(drive_id) = remote
            .parent_reference
            .as_ref()
            .and_then(|reference| reference.drive_id.as_ref())
        else {
            continue;
        };

        let target = RemoteDriveTarget {
            drive_id: drive_id.clone(),
            item_id: remote.id.clone(),
        };

        let remote_url_matches = remote
            .web_dav_url
            .as_deref()
            .into_iter()
            .chain(remote.web_url.as_deref())
            .any(|url| comparable_reference_url(url) == wanted_url);

        if remote_url_matches {
            push_unique_target(&mut url_matches, target.clone());
        }

        let remote_name = remote.name.as_deref().or(item.name.as_deref());
        if remote_name.is_some_and(|candidate| candidate.eq_ignore_ascii_case(name)) {
            push_unique_target(&mut name_matches, target);
        }
    }

    if url_matches.len() == 1 {
        return url_matches.into_iter().next();
    }

    if name_matches.len() == 1 {
        return name_matches.into_iter().next();
    }

    None
}

fn remote_drive_target_from_microsoft_search(
    response: &serde_json::Value,
    name: &str,
    content_url: &str,
) -> Option<RemoteDriveTarget> {
    let wanted_url = comparable_reference_url(content_url);
    let mut url_matches = Vec::new();
    let mut name_matches = Vec::new();

    let search_responses = response.get("value").and_then(|value| value.as_array())?;

    for search_response in search_responses {
        let Some(containers) = search_response
            .get("hitsContainers")
            .and_then(|value| value.as_array())
        else {
            continue;
        };

        for container in containers {
            let Some(hits) = container.get("hits").and_then(|value| value.as_array()) else {
                continue;
            };

            for hit in hits {
                let Some(resource) = hit.get("resource") else {
                    continue;
                };
                let Some(drive_id) = resource
                    .get("parentReference")
                    .and_then(|value| value.get("driveId"))
                    .and_then(|value| value.as_str())
                else {
                    continue;
                };
                let item_id = resource
                    .get("id")
                    .and_then(|value| value.as_str())
                    .or_else(|| hit.get("hitId").and_then(|value| value.as_str()));
                let Some(item_id) = item_id else {
                    continue;
                };

                let target = RemoteDriveTarget {
                    drive_id: drive_id.to_string(),
                    item_id: item_id.to_string(),
                };

                let remote_url_matches = ["webDavUrl", "webUrl"].into_iter().any(|field| {
                    resource
                        .get(field)
                        .and_then(|value| value.as_str())
                        .is_some_and(|url| comparable_reference_url(url) == wanted_url)
                });
                if remote_url_matches {
                    push_unique_target(&mut url_matches, target.clone());
                }

                if resource
                    .get("name")
                    .and_then(|value| value.as_str())
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(name))
                {
                    push_unique_target(&mut name_matches, target);
                }
            }
        }
    }

    if url_matches.len() == 1 {
        return url_matches.into_iter().next();
    }

    if name_matches.len() == 1 {
        return name_matches.into_iter().next();
    }

    None
}

async fn search_accessible_remote_drive_item(
    graph: &GraphClient,
    name: &str,
    content_url: &str,
) -> Result<Option<RemoteDriveTarget>> {
    let encoded_name = graph_search_term(name);
    let search_path = format!("me/drive/search(q='{encoded_name}')?$top=25");
    tracing::debug!("Teams reference attachment searching accessible drive items: {search_path}");

    let page = graph.get_json::<DriveSearchPage>(&search_path).await?;
    tracing::debug!(
        "Teams reference attachment accessible-file search returned {} result(s)",
        page.value.len()
    );

    Ok(remote_drive_target_from_search(&page, name, content_url))
}

async fn search_microsoft_remote_drive_item(
    graph: &GraphClient,
    name: &str,
    content_url: &str,
) -> Result<Option<RemoteDriveTarget>> {
    let payload = serde_json::json!({
        "requests": [{
            "entityTypes": ["driveItem"],
            "query": { "queryString": name },
            "from": 0,
            "size": 25,
            "fields": ["id", "name", "webUrl", "webDavUrl", "parentReference"]
        }]
    });

    tracing::debug!("Teams reference attachment searching Microsoft Search API for {name}");
    let response: serde_json::Value = graph.post_json("search/query", &payload).await?;
    Ok(remote_drive_target_from_microsoft_search(
        &response,
        name,
        content_url,
    ))
}

async fn sender_drive_file_bytes(
    graph: &GraphClient,
    sender_user_id: Option<&str>,
    content_url: &str,
) -> Result<Option<Vec<u8>>> {
    let Some(item_path) = personal_webdav_item_path(content_url) else {
        return Ok(None);
    };
    let Some(user_id) = sender_user_id else {
        return Ok(None);
    };

    #[derive(serde::Deserialize)]
    struct DriveItemId {
        id: String,
    }

    let resolve_path = format!("users/{user_id}/drive/root:/{item_path}");
    tracing::debug!("Teams reference attachment resolving sender OneDrive path: {resolve_path}");
    let item: DriveItemId = graph.get_json(&resolve_path).await?;

    let content_path = format!("users/{user_id}/drive/items/{}/content", item.id);
    tracing::debug!("Teams reference attachment downloading sender driveItem: {content_path}");
    Ok(Some(graph.get_bytes(&content_path).await?))
}

async fn remote_target_file_bytes(
    graph: &GraphClient,
    target: RemoteDriveTarget,
    route_label: &str,
) -> Result<Vec<u8>> {
    let content_path = format!(
        "drives/{}/items/{}/content",
        target.drive_id, target.item_id
    );
    tracing::debug!("Teams reference attachment downloading {route_label} target: {content_path}");
    graph.get_bytes(&content_path).await
}

async fn try_reference_route(
    graph: &GraphClient,
    route: ReferenceRoute,
    sender_user_id: Option<&str>,
    name: &str,
    content_url: &str,
) -> Result<Option<Vec<u8>>> {
    match route {
        ReferenceRoute::SenderDrive => {
            sender_drive_file_bytes(graph, sender_user_id, content_url).await
        }
        ReferenceRoute::AccessibleDriveSearch => {
            let Some(target) =
                search_accessible_remote_drive_item(graph, name, content_url).await?
            else {
                return Ok(None);
            };
            remote_target_file_bytes(graph, target, route.label())
                .await
                .map(Some)
        }
        ReferenceRoute::MicrosoftSearch => {
            let Some(target) = search_microsoft_remote_drive_item(graph, name, content_url).await?
            else {
                return Ok(None);
            };
            remote_target_file_bytes(graph, target, route.label())
                .await
                .map(Some)
        }
        ReferenceRoute::Shares => shared_file_bytes(graph, content_url).await.map(Some),
    }
}

/// Download a Teams `reference` attachment through a session-adaptive resolver.
///
/// Four strategies are retained. Their initial order depends on URL shape, then a
/// strategy that succeeds at least twice for the same route class + host is
/// promoted for the rest of the current process/session. Nothing is persisted.
pub async fn reference_file_bytes(
    graph: &GraphClient,
    sender_user_id: Option<&str>,
    name: &str,
    content_url: &str,
) -> Result<Vec<u8>> {
    let class = reference_route_class(content_url);
    let host = reference_url_host(content_url);
    let routes = ordered_reference_routes(content_url);

    tracing::debug!(
        "Teams reference attachment resolver class={} host={} order={}",
        class.label(),
        if host.is_empty() { "(unknown)" } else { &host },
        routes
            .iter()
            .map(|route| route.label())
            .collect::<Vec<_>>()
            .join(" -> ")
    );

    let mut last_error = None;

    for route in routes {
        match try_reference_route(graph, route, sender_user_id, name, content_url).await {
            Ok(Some(bytes)) => {
                record_reference_route_result(content_url, route, true);
                tracing::debug!(
                    "Teams reference attachment resolver strategy {} succeeded",
                    route.label()
                );
                return Ok(bytes);
            }
            Ok(None) => {
                tracing::debug!(
                    "Teams reference attachment resolver strategy {} had no applicable/unique target",
                    route.label()
                );
            }
            Err(error) => {
                record_reference_route_result(content_url, route, false);
                tracing::debug!(
                    "Teams reference attachment resolver strategy {} failed: {error:#}",
                    route.label()
                );
                last_error = Some(error);
            }
        }
    }

    Err(last_error.unwrap_or_else(|| {
        anyhow::anyhow!(
            "Teams reference attachment {name} could not be resolved by any file strategy"
        )
    }))
}

// Hosted-content variants retain support for inline image composition.
pub async fn send_message_with_body(
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
pub async fn send_reply_with_body(
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
pub async fn update_message_with_body(
    graph: &GraphClient,
    chat_id: &str,
    original: &ChatMessage,
    body: OutgoingBody<'_>,
) -> Result<()> {
    graph
        .patch(
            &format!("chats/{chat_id}/messages/{}", original.id),
            &update_payload(original, body),
        )
        .await
}

#[cfg(test)]
mod reference_search_tests {
    use super::{
        graph_search_term, looks_like_share_link, remote_drive_target_from_search,
        DriveParentReference, DriveSearchItem, DriveSearchPage, RemoteDriveItem, RemoteDriveTarget,
    };

    fn remote(name: &str, web_dav_url: Option<&str>, drive: &str, item: &str) -> DriveSearchItem {
        DriveSearchItem {
            name: Some(name.to_string()),
            remote_item: Some(RemoteDriveItem {
                id: item.to_string(),
                name: Some(name.to_string()),
                web_dav_url: web_dav_url.map(str::to_string),
                web_url: None,
                parent_reference: Some(DriveParentReference {
                    drive_id: Some(drive.to_string()),
                }),
            }),
        }
    }

    #[test]
    fn distinguishes_webdav_from_real_share_link() {
        assert!(!looks_like_share_link(
            "https://tenant-my.sharepoint.com/personal/alice_example_com/Documents/Microsoft%20Teams%20Chat%20Files/test.jpg"
        ));
        assert!(looks_like_share_link(
            "https://tenant-my.sharepoint.com/:i:/g/personal/alice_example_com/EXAMPLE?e=abc123"
        ));
    }

    #[test]
    fn encodes_drive_search_term() {
        assert_eq!(graph_search_term("test one.jpg"), "test%20one.jpg");
        assert_eq!(graph_search_term("O'Connor.png"), "O%27%27Connor.png");
    }

    #[test]
    fn prefers_exact_webdav_url_match() {
        let wanted = "https://tenant-my.sharepoint.com/personal/alice/Documents/Microsoft%20Teams%20Chat%20Files/test.jpg";
        let page = DriveSearchPage {
            value: vec![
                remote(
                    "test.jpg",
                    Some("https://other.example/Documents/test.jpg"),
                    "drive-a",
                    "item-a",
                ),
                remote("test.jpg", Some(wanted), "drive-b", "item-b"),
            ],
        };

        assert_eq!(
            remote_drive_target_from_search(&page, "test.jpg", wanted),
            Some(RemoteDriveTarget {
                drive_id: "drive-b".to_string(),
                item_id: "item-b".to_string(),
            })
        );
    }

    #[test]
    fn accepts_one_unique_exact_name_when_webdav_is_missing() {
        let page = DriveSearchPage {
            value: vec![
                remote("other.jpg", None, "drive-a", "item-a"),
                remote("test.jpg", None, "drive-b", "item-b"),
            ],
        };

        assert_eq!(
            remote_drive_target_from_search(
                &page,
                "test.jpg",
                "https://tenant.example/Documents/test.jpg"
            ),
            Some(RemoteDriveTarget {
                drive_id: "drive-b".to_string(),
                item_id: "item-b".to_string(),
            })
        );
    }

    #[test]
    fn refuses_ambiguous_exact_name_matches() {
        let page = DriveSearchPage {
            value: vec![
                remote("test.jpg", None, "drive-a", "item-a"),
                remote("test.jpg", None, "drive-b", "item-b"),
            ],
        };

        assert_eq!(
            remote_drive_target_from_search(
                &page,
                "test.jpg",
                "https://tenant.example/Documents/test.jpg"
            ),
            None
        );
    }
}

#[cfg(test)]
mod adaptive_reference_route_tests {
    use super::{
        base_reference_routes, ordered_reference_routes, record_reference_route_result,
        remote_drive_target_from_microsoft_search, ReferenceRoute, ReferenceRouteClass,
        RemoteDriveTarget,
    };

    #[test]
    fn personal_webdav_starts_with_sender_drive() {
        assert_eq!(
            base_reference_routes(ReferenceRouteClass::PersonalWebDav),
            [
                ReferenceRoute::SenderDrive,
                ReferenceRoute::AccessibleDriveSearch,
                ReferenceRoute::MicrosoftSearch,
                ReferenceRoute::Shares,
            ]
        );
    }

    #[test]
    fn one_success_does_not_promote_but_two_do() {
        let url = "https://adaptive-one.example/personal/alice/Documents/Microsoft%20Teams%20Chat%20Files/x.jpg";

        let initial = ordered_reference_routes(url);
        assert_eq!(initial[0], ReferenceRoute::SenderDrive);

        record_reference_route_result(url, ReferenceRoute::MicrosoftSearch, true);
        let after_one = ordered_reference_routes(url);
        assert_eq!(after_one[0], ReferenceRoute::SenderDrive);

        record_reference_route_result(url, ReferenceRoute::MicrosoftSearch, true);
        let after_two = ordered_reference_routes(url);
        assert_eq!(after_two[0], ReferenceRoute::MicrosoftSearch);
    }

    #[test]
    fn route_learning_is_scoped_by_host() {
        let learned = "https://adaptive-learned.example/personal/alice/Documents/x.jpg";
        let fresh = "https://adaptive-fresh.example/personal/alice/Documents/x.jpg";

        record_reference_route_result(learned, ReferenceRoute::MicrosoftSearch, true);
        record_reference_route_result(learned, ReferenceRoute::MicrosoftSearch, true);

        assert_eq!(
            ordered_reference_routes(learned)[0],
            ReferenceRoute::MicrosoftSearch
        );
        assert_eq!(
            ordered_reference_routes(fresh)[0],
            ReferenceRoute::SenderDrive
        );
    }

    #[test]
    fn microsoft_search_prefers_exact_url_match() {
        let wanted = "https://tenant-my.sharepoint.com/personal/alice/Documents/x.jpg";
        let response = serde_json::json!({
            "value": [{
                "hitsContainers": [{
                    "hits": [
                        {
                            "hitId": "item-a",
                            "resource": {
                                "id": "item-a",
                                "name": "x.jpg",
                                "webDavUrl": "https://other.example/Documents/x.jpg",
                                "parentReference": { "driveId": "drive-a" }
                            }
                        },
                        {
                            "hitId": "item-b",
                            "resource": {
                                "id": "item-b",
                                "name": "x.jpg",
                                "webDavUrl": wanted,
                                "parentReference": { "driveId": "drive-b" }
                            }
                        }
                    ]
                }]
            }]
        });

        assert_eq!(
            remote_drive_target_from_microsoft_search(&response, "x.jpg", wanted),
            Some(RemoteDriveTarget {
                drive_id: "drive-b".into(),
                item_id: "item-b".into(),
            })
        );
    }

    #[test]
    fn microsoft_search_refuses_ambiguous_name_only_matches() {
        let response = serde_json::json!({
            "value": [{
                "hitsContainers": [{
                    "hits": [
                        {
                            "hitId": "item-a",
                            "resource": {
                                "id": "item-a",
                                "name": "x.jpg",
                                "parentReference": { "driveId": "drive-a" }
                            }
                        },
                        {
                            "hitId": "item-b",
                            "resource": {
                                "id": "item-b",
                                "name": "x.jpg",
                                "parentReference": { "driveId": "drive-b" }
                            }
                        }
                    ]
                }]
            }]
        });

        assert_eq!(
            remote_drive_target_from_microsoft_search(
                &response,
                "x.jpg",
                "https://unknown.example/x.jpg"
            ),
            None
        );
    }
}
