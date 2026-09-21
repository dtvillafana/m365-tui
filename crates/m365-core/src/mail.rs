//! Outlook mail endpoints.

use anyhow::Result;
use serde_json::json;

use crate::graph::{DeltaPage, GraphClient};
use crate::models::{Attachment, MailFolder, MailMessage};
use crate::util::{base64_encode, html_escape};
use bytes::Bytes;

/// List all mail folders, including hidden folders and nested custom folders.
/// Nested display names include their parent path.
pub async fn list_folders(graph: &GraphClient) -> Result<Vec<MailFolder>> {
    let query = "$top=100&includeHiddenFolders=true\
                 &$select=id,displayName,unreadItemCount,totalItemCount,childFolderCount";
    let mut pending: Vec<MailFolder> = graph
        .get_collection(&format!("me/mailFolders?{query}"))
        .await?;
    pending.reverse();
    let mut folders = Vec::new();
    while let Some(folder) = pending.pop() {
        if folder.child_folder_count.unwrap_or(0) > 0 {
            let children: Vec<MailFolder> = graph
                .get_collection(&format!(
                    "me/mailFolders/{}/childFolders?{query}",
                    folder.id
                ))
                .await?;
            for mut child in children.into_iter().rev() {
                child.display_name = Some(format!(
                    "{} / {}",
                    folder.display_name.as_deref().unwrap_or_default(),
                    child.display_name.as_deref().unwrap_or_default()
                ));
                pending.push(child);
            }
        }
        folders.push(folder);
    }
    Ok(folders)
}

/// List the first page of messages in a folder, newest first. Returns the page
/// and the `@odata.nextLink` for "load more" (if the folder has more).
pub async fn list_messages(
    graph: &GraphClient,
    folder_id: &str,
    top: u32,
) -> Result<(Vec<MailMessage>, Option<String>)> {
    let path = format!(
        "me/mailFolders/{folder_id}/messages?$top={top}&$orderby=receivedDateTime desc\
         &$select=id,conversationId,subject,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,webLink"
    );
    // Single page only — `$top` bounds it; we don't want to walk the whole folder.
    graph.get_page_with_next(&path).await
}

/// Fetch the next page of messages from an `@odata.nextLink` returned by
/// [`list_messages`].
pub async fn list_messages_more(
    graph: &GraphClient,
    next_link: &str,
) -> Result<(Vec<MailMessage>, Option<String>)> {
    graph.get_page_with_next(next_link).await
}

/// Fetch a single message including its full body.
pub async fn get_message(graph: &GraphClient, id: &str) -> Result<MailMessage> {
    let path = format!(
        "me/messages/{id}?$select=id,conversationId,subject,body,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,webLink"
    );
    graph.get_json(&path).await
}

/// Incremental sync of a folder. Pass `None` for the first call, then feed back
/// the returned `delta_link` on each subsequent poll.
pub async fn delta_messages(
    graph: &GraphClient,
    folder_id: &str,
    delta_link: Option<&str>,
) -> Result<DeltaPage<MailMessage>> {
    let path = match delta_link {
        Some(link) => link.to_string(),
        None => format!("me/mailFolders/{folder_id}/messages/delta?$select=id,conversationId,subject,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead"),
    };
    graph.delta(&path).await
}

/// List a message's attachments. `$select` keeps `contentBytes` out of the
/// response so listing stays cheap regardless of attachment size.
///
/// `contentId` is not on the base `attachment` type. Graph rejects it in
/// `$select` even after an OData `fileAttachment` cast, so CID values are
/// read from unselected inline attachments instead — those payloads are
/// small, and extra fields such as `contentBytes` are ignored here.
pub async fn list_attachments(graph: &GraphClient, message_id: &str) -> Result<Vec<Attachment>> {
    let path =
        format!("me/messages/{message_id}/attachments?$select=id,name,contentType,size,isInline");
    let mut attachments: Vec<Attachment> = graph.get_collection(&path).await?;
    if !attachments.iter().any(|a| a.is_inline.unwrap_or(false)) {
        return Ok(attachments);
    }
    let inline: Vec<Attachment> = graph
        .get_collection(&format!(
            "me/messages/{message_id}/attachments?$filter=isInline eq true"
        ))
        .await?;
    for file in inline {
        if let Some(attachment) = attachments.iter_mut().find(|a| a.id == file.id) {
            attachment.content_id = file.content_id;
        }
    }
    Ok(attachments)
}

/// Download one attachment's raw bytes.
pub async fn download_attachment(
    graph: &GraphClient,
    message_id: &str,
    attachment_id: &str,
) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!(
            "me/messages/{message_id}/attachments/{attachment_id}/$value"
        ))
        .await
}

/// Search across the mailbox using Graph `$search`.
pub async fn search(graph: &GraphClient, query: &str, top: u32) -> Result<Vec<MailMessage>> {
    // $search requires ConsistencyLevel semantics; Graph accepts the quoted form.
    let escaped = query.replace('"', "");
    let path = format!(
        "me/messages?$search=\"{escaped}\"&$top={top}\
         &$select=id,conversationId,subject,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead"
    );
    graph.get_page(&path).await
}

/// Fetch one mailbox-wide reply thread. Graph cannot aggregate folder messages
/// into conversations, so callers collapse folder pages by `conversationId`
/// and use this filtered query only when a thread is opened.
pub async fn list_conversation(
    graph: &GraphClient,
    conversation_id: &str,
    top: u32,
) -> Result<(Vec<MailMessage>, bool)> {
    let escaped = escape_odata_string(conversation_id);
    let path = format!(
        "me/messages?$filter=conversationId eq '{escaped}'&$top={top}\
         &$select=id,conversationId,subject,body,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,webLink"
    );
    let (mut messages, next): (Vec<MailMessage>, Option<String>) =
        graph.get_page_with_next(&path).await?;
    // Although `/me/messages` is mailbox-wide, explicitly merge Sent Items.
    // Some mailbox configurations omit sent copies from that collection.
    let sent_path = format!(
        "me/mailFolders/sentitems/messages?$filter=conversationId eq '{escaped}'&$top={top}\
         &$select=id,conversationId,subject,body,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,webLink"
    );
    let (sent, sent_next): (Vec<MailMessage>, Option<String>) =
        graph.get_page_with_next(&sent_path).await?;
    let known: std::collections::HashSet<String> =
        messages.iter().map(|message| message.id.clone()).collect();
    messages.extend(
        sent.into_iter()
            .filter(|message| !known.contains(&message.id)),
    );
    sort_conversation_messages(&mut messages);
    Ok((messages, next.is_some() || sent_next.is_some()))
}

fn escape_odata_string(value: &str) -> String {
    value.replace('\'', "''")
}

fn sort_conversation_messages(messages: &mut [MailMessage]) {
    messages.sort_by(|a, b| {
        b.mail_time()
            .cmp(&a.mail_time())
            .then_with(|| b.id.cmp(&a.id))
    });
}

/// Mark a message as read or unread.
pub async fn mark_read(graph: &GraphClient, id: &str, read: bool) -> Result<()> {
    graph
        .patch(&format!("me/messages/{id}"), &json!({ "isRead": read }))
        .await
}

/// Move a message into another folder. `destination_id` is a folder id or a
/// well-known name such as `deleteditems`.
pub async fn move_message(graph: &GraphClient, id: &str, destination_id: &str) -> Result<()> {
    graph
        .post_action(
            &format!("me/messages/{id}/move"),
            &json!({ "destinationId": destination_id }),
        )
        .await
}

/// Graph treats mail bodies and reply comments as HTML, so composer newlines
/// have to become `<br>` or they collapse on the way out.
fn html_body(text: &str) -> serde_json::Value {
    json!({ "contentType": "HTML", "content": html_escape(text) })
}

fn outgoing_body(text: &str, is_html: bool) -> serde_json::Value {
    if is_html {
        json!({ "contentType": "HTML", "content": text })
    } else {
        html_body(text)
    }
}

/// Send a new message.
pub async fn send_mail(
    graph: &GraphClient,
    to: &[String],
    cc: &[String],
    bcc: &[String],
    subject: &str,
    body: &str,
    body_is_html: bool,
) -> Result<()> {
    let payload = json!({
        "message": {
            "subject": subject,
            "body": outgoing_body(body, body_is_html),
            "toRecipients": recipient_values(to),
            "ccRecipients": recipient_values(cc),
            "bccRecipients": recipient_values(bcc),
        },
        "saveToSentItems": true,
    });
    graph.post_action("me/sendMail", &payload).await
}

/// Reply to a message (Graph fills quoting + recipients automatically).
pub async fn reply(graph: &GraphClient, id: &str, comment: &str) -> Result<()> {
    graph
        .post_action(
            &format!("me/messages/{id}/reply"),
            &json!({ "comment": html_escape(comment) }),
        )
        .await
}

/// Reply-all to a message.
pub async fn reply_all(graph: &GraphClient, id: &str, comment: &str) -> Result<()> {
    graph
        .post_action(
            &format!("me/messages/{id}/replyAll"),
            &json!({ "comment": html_escape(comment) }),
        )
        .await
}

/// Forward a message to new recipients with an optional comment.
pub async fn forward(graph: &GraphClient, id: &str, to: &[String], comment: &str) -> Result<()> {
    let recipients: Vec<_> = to
        .iter()
        .map(|addr| json!({ "emailAddress": { "address": addr } }))
        .collect();
    graph
        .post_action(
            &format!("me/messages/{id}/forward"),
            &json!({ "comment": html_escape(comment), "toRecipients": recipients }),
        )
        .await
}

// ---------------------------------------------------------------------------
// Sending with attachments
// ---------------------------------------------------------------------------

/// Graph accepts attachment bytes inline below ~3 MB; larger ones need an
/// upload session.
const INLINE_ATTACHMENT_LIMIT: usize = 3 * 1024 * 1024;

/// Upload chunks must be a multiple of 320 KiB. 10 x 320 KiB = 3.2 MB.
const UPLOAD_CHUNK: usize = 10 * 320 * 1024;

/// What kind of message is being sent.
#[derive(Debug, Clone)]
pub enum Outgoing {
    New {
        to: Vec<String>,
        cc: Vec<String>,
        bcc: Vec<String>,
        subject: String,
    },
    Reply {
        id: String,
        recipients: Option<OutgoingRecipients>,
    },
    ReplyAll {
        id: String,
        recipients: Option<OutgoingRecipients>,
    },
    Forward {
        id: String,
        to: Vec<String>,
        cc: Vec<String>,
        bcc: Vec<String>,
    },
}

#[derive(Debug, Clone)]
pub struct OutgoingRecipients {
    pub to: Vec<String>,
    pub cc: Vec<String>,
    pub bcc: Vec<String>,
}

/// A file to attach: display name plus its bytes.
#[derive(Debug, Clone)]
pub struct OutgoingAttachment {
    pub name: String,
    pub bytes: Vec<u8>,
    pub content_type: Option<String>,
    pub content_id: Option<String>,
    pub is_inline: bool,
}

/// Send a message, with or without attachments.
///
/// Without attachments this uses the one-shot `sendMail`/`reply`/`replyAll`/
/// `forward` actions. With attachments it must go through a draft, because
/// those actions cannot carry files: create the draft, attach each file
/// (inline for small ones, an upload session for large ones), then send it.
pub async fn send_message(
    graph: &GraphClient,
    kind: Outgoing,
    body: &str,
    body_is_html: bool,
    attachments: Vec<OutgoingAttachment>,
) -> Result<()> {
    let needs_recipient_draft = matches!(
        &kind,
        Outgoing::Forward { cc, bcc, .. } if !cc.is_empty() || !bcc.is_empty()
    ) || matches!(
        &kind,
        Outgoing::Reply {
            recipients: Some(_),
            ..
        } | Outgoing::ReplyAll {
            recipients: Some(_),
            ..
        }
    );
    if attachments.is_empty() && !needs_recipient_draft {
        return send_simple(graph, kind, body, body_is_html).await;
    }

    let draft = create_draft(graph, kind, body, body_is_html).await?;
    for att in attachments {
        attach_to_draft(graph, &draft.id, att).await?;
    }
    graph
        .post_action(&format!("me/messages/{}/send", draft.id), &json!({}))
        .await
}

async fn send_simple(
    graph: &GraphClient,
    kind: Outgoing,
    body: &str,
    body_is_html: bool,
) -> Result<()> {
    match kind {
        Outgoing::New {
            to,
            cc,
            bcc,
            subject,
        } => send_mail(graph, &to, &cc, &bcc, &subject, body, body_is_html).await,
        Outgoing::Reply { id, .. } => reply(graph, &id, body).await,
        Outgoing::ReplyAll { id, .. } => reply_all(graph, &id, body).await,
        Outgoing::Forward { id, to, .. } => forward(graph, &id, &to, body).await,
    }
}

/// Create a draft for the given kind, with the user's text in place.
async fn create_draft(
    graph: &GraphClient,
    kind: Outgoing,
    body: &str,
    body_is_html: bool,
) -> Result<MailMessage> {
    match kind {
        Outgoing::New {
            to,
            cc,
            bcc,
            subject,
        } => {
            graph
                .post_json(
                    "me/messages",
                    &json!({
                        "subject": subject,
                        "body": outgoing_body(body, body_is_html),
                        "toRecipients": recipient_values(&to),
                        "ccRecipients": recipient_values(&cc),
                        "bccRecipients": recipient_values(&bcc),
                    }),
                )
                .await
        }
        Outgoing::Reply { id, recipients } => {
            let draft: MailMessage = graph
                .post_json(&format!("me/messages/{id}/createReply"), &json!({}))
                .await?;
            patch_draft_recipients(graph, &draft.id, recipients.as_ref()).await?;
            prepend_comment(graph, draft, body, body_is_html).await
        }
        Outgoing::ReplyAll { id, recipients } => {
            let draft: MailMessage = graph
                .post_json(&format!("me/messages/{id}/createReplyAll"), &json!({}))
                .await?;
            patch_draft_recipients(graph, &draft.id, recipients.as_ref()).await?;
            prepend_comment(graph, draft, body, body_is_html).await
        }
        Outgoing::Forward { id, to, cc, bcc } => {
            let draft: MailMessage = graph
                .post_json(&format!("me/messages/{id}/createForward"), &json!({}))
                .await?;
            graph
                .patch(
                    &format!("me/messages/{}", draft.id),
                    &json!({
                        "toRecipients": recipient_values(&to),
                        "ccRecipients": recipient_values(&cc),
                        "bccRecipients": recipient_values(&bcc),
                    }),
                )
                .await?;
            prepend_comment(graph, draft, body, body_is_html).await
        }
    }
}

async fn patch_draft_recipients(
    graph: &GraphClient,
    draft_id: &str,
    recipients: Option<&OutgoingRecipients>,
) -> Result<()> {
    let Some(recipients) = recipients else {
        return Ok(());
    };
    graph
        .patch(
            &format!("me/messages/{draft_id}"),
            &json!({
                "toRecipients": recipient_values(&recipients.to),
                "ccRecipients": recipient_values(&recipients.cc),
                "bccRecipients": recipient_values(&recipients.bcc),
            }),
        )
        .await
}

fn recipient_values(addresses: &[String]) -> Vec<serde_json::Value> {
    addresses
        .iter()
        .map(|address| json!({ "emailAddress": { "address": address } }))
        .collect()
}

/// A reply/forward draft already contains the quoted original; put the user's
/// text above it (the one-shot actions do this for us, drafts do not).
async fn prepend_comment(
    graph: &GraphClient,
    draft: MailMessage,
    comment: &str,
    comment_is_html: bool,
) -> Result<MailMessage> {
    if comment.trim().is_empty() {
        return Ok(draft);
    }
    let original = draft
        .body
        .as_ref()
        .and_then(|b| b.content.clone())
        .unwrap_or_default();
    let comment = if comment_is_html {
        comment.to_string()
    } else {
        html_escape(comment)
    };
    let merged = format!("<div>{comment}</div>{original}");
    graph
        .patch(
            &format!("me/messages/{}", draft.id),
            &json!({ "body": { "contentType": "HTML", "content": merged } }),
        )
        .await?;
    Ok(draft)
}

async fn attach_to_draft(
    graph: &GraphClient,
    draft_id: &str,
    att: OutgoingAttachment,
) -> Result<()> {
    if att.bytes.len() <= INLINE_ATTACHMENT_LIMIT {
        let mut payload = json!({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": att.name,
            "isInline": att.is_inline,
            "contentBytes": base64_encode(&att.bytes),
        });
        if let Some(content_type) = att.content_type {
            payload["contentType"] = json!(content_type);
        }
        if let Some(content_id) = att.content_id {
            payload["contentId"] = json!(content_id);
        }
        graph
            .post_action(&format!("me/messages/{draft_id}/attachments"), &payload)
            .await
    } else {
        upload_large_attachment(graph, draft_id, att).await
    }
}

#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct UploadSession {
    upload_url: String,
}

async fn upload_large_attachment(
    graph: &GraphClient,
    draft_id: &str,
    att: OutgoingAttachment,
) -> Result<()> {
    let total = att.bytes.len() as u64;
    let session: UploadSession = graph
        .post_json(
            &format!("me/messages/{draft_id}/attachments/createUploadSession"),
            &json!({
                "AttachmentItem": {
                    "attachmentType": "file",
                    "name": att.name,
                    "size": total,
                }
            }),
        )
        .await?;

    // Move the file buffer into `Bytes` once; each chunk is then a refcounted
    // view of it, so the upload never copies the payload again.
    let data = Bytes::from(att.bytes);
    let mut start = 0usize;
    while start < data.len() {
        let end = (start + UPLOAD_CHUNK).min(data.len());
        let done = graph
            .put_upload_chunk(
                &session.upload_url,
                start as u64,
                total,
                data.slice(start..end),
            )
            .await?;
        start = end;
        if done {
            break;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn html_body_turns_newlines_into_breaks() {
        let body = html_body("line1\nline2");
        assert_eq!(body["contentType"], "HTML");
        assert_eq!(body["content"], "line1<br>line2");
    }

    #[test]
    fn prepared_html_body_is_not_escaped_again() {
        let body = outgoing_body("<img src=\"cid:x\">", true);
        assert_eq!(body["content"], "<img src=\"cid:x\">");
    }

    #[test]
    fn conversation_filter_escapes_odata_quotes() {
        assert_eq!(escape_odata_string("a'b"), "a''b");
    }

    #[test]
    fn conversation_messages_sort_newest_first() {
        let message = |id: &str, received: &str| {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "receivedDateTime": received
            }))
            .unwrap()
        };
        let mut messages = [
            message("new", "2026-09-17T15:00:00Z"),
            message("old", "2026-09-17T13:00:00Z"),
            message("middle", "2026-09-17T14:00:00Z"),
            serde_json::from_value(serde_json::json!({
                "id": "sent",
                "sentDateTime": "2026-09-17T16:00:00Z"
            }))
            .unwrap(),
        ];
        sort_conversation_messages(&mut messages);
        assert_eq!(
            messages
                .iter()
                .map(|message| message.id.as_str())
                .collect::<Vec<_>>(),
            vec!["sent", "new", "middle", "old"]
        );
    }

    #[test]
    fn recipient_values_build_graph_addresses() {
        let values = recipient_values(&["a@example.com".into(), "b@example.com".into()]);
        assert_eq!(values[0]["emailAddress"]["address"], "a@example.com");
        assert_eq!(values[1]["emailAddress"]["address"], "b@example.com");
    }
}
