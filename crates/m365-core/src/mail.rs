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
         &$select=id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,webLink"
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
        "me/messages/{id}?$select=id,subject,body,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments,webLink"
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
        None => format!("me/mailFolders/{folder_id}/messages/delta?$select=id,subject,bodyPreview,from,receivedDateTime,isRead"),
    };
    graph.delta(&path).await
}

/// List a message's attachments. `$select` keeps `contentBytes` out of the
/// response so listing stays cheap regardless of attachment size.
pub async fn list_attachments(graph: &GraphClient, message_id: &str) -> Result<Vec<Attachment>> {
    let path =
        format!("me/messages/{message_id}/attachments?$select=id,name,contentType,size,isInline");
    graph.get_collection(&path).await
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
         &$select=id,subject,bodyPreview,from,receivedDateTime,isRead"
    );
    graph.get_page(&path).await
}

pub async fn mark_read(graph: &GraphClient, id: &str, read: bool) -> Result<()> {
    graph
        .patch(&format!("me/messages/{id}"), &json!({ "isRead": read }))
        .await
}

/// Send a new message.
pub async fn send_mail(
    graph: &GraphClient,
    to: &[String],
    subject: &str,
    body: &str,
) -> Result<()> {
    let recipients: Vec<_> = to
        .iter()
        .map(|addr| json!({ "emailAddress": { "address": addr } }))
        .collect();
    let payload = json!({
        "message": {
            "subject": subject,
            "body": { "contentType": "Text", "content": body },
            "toRecipients": recipients,
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
            &json!({ "comment": comment }),
        )
        .await
}

/// Reply-all to a message.
pub async fn reply_all(graph: &GraphClient, id: &str, comment: &str) -> Result<()> {
    graph
        .post_action(
            &format!("me/messages/{id}/replyAll"),
            &json!({ "comment": comment }),
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
            &json!({ "comment": comment, "toRecipients": recipients }),
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
    New { to: Vec<String>, subject: String },
    Reply { id: String },
    ReplyAll { id: String },
    Forward { id: String, to: Vec<String> },
}

/// A file to attach: display name plus its bytes.
#[derive(Debug, Clone)]
pub struct OutgoingAttachment {
    pub name: String,
    pub bytes: Vec<u8>,
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
    attachments: Vec<OutgoingAttachment>,
) -> Result<()> {
    if attachments.is_empty() {
        return send_simple(graph, kind, body).await;
    }

    let draft = create_draft(graph, kind, body).await?;
    for att in attachments {
        attach_to_draft(graph, &draft.id, att).await?;
    }
    graph
        .post_action(&format!("me/messages/{}/send", draft.id), &json!({}))
        .await
}

async fn send_simple(graph: &GraphClient, kind: Outgoing, body: &str) -> Result<()> {
    match kind {
        Outgoing::New { to, subject } => send_mail(graph, &to, &subject, body).await,
        Outgoing::Reply { id } => reply(graph, &id, body).await,
        Outgoing::ReplyAll { id } => reply_all(graph, &id, body).await,
        Outgoing::Forward { id, to } => forward(graph, &id, &to, body).await,
    }
}

/// Create a draft for the given kind, with the user's text in place.
async fn create_draft(graph: &GraphClient, kind: Outgoing, body: &str) -> Result<MailMessage> {
    match kind {
        Outgoing::New { to, subject } => {
            let recipients: Vec<_> = to
                .iter()
                .map(|a| json!({ "emailAddress": { "address": a } }))
                .collect();
            graph
                .post_json(
                    "me/messages",
                    &json!({
                        "subject": subject,
                        "body": { "contentType": "Text", "content": body },
                        "toRecipients": recipients,
                    }),
                )
                .await
        }
        Outgoing::Reply { id } => {
            let draft: MailMessage = graph
                .post_json(&format!("me/messages/{id}/createReply"), &json!({}))
                .await?;
            prepend_comment(graph, draft, body).await
        }
        Outgoing::ReplyAll { id } => {
            let draft: MailMessage = graph
                .post_json(&format!("me/messages/{id}/createReplyAll"), &json!({}))
                .await?;
            prepend_comment(graph, draft, body).await
        }
        Outgoing::Forward { id, to } => {
            let recipients: Vec<_> = to
                .iter()
                .map(|a| json!({ "emailAddress": { "address": a } }))
                .collect();
            let draft: MailMessage = graph
                .post_json(
                    &format!("me/messages/{id}/createForward"),
                    &json!({ "toRecipients": recipients }),
                )
                .await?;
            prepend_comment(graph, draft, body).await
        }
    }
}

/// A reply/forward draft already contains the quoted original; put the user's
/// text above it (the one-shot actions do this for us, drafts do not).
async fn prepend_comment(
    graph: &GraphClient,
    draft: MailMessage,
    comment: &str,
) -> Result<MailMessage> {
    if comment.trim().is_empty() {
        return Ok(draft);
    }
    let original = draft
        .body
        .as_ref()
        .and_then(|b| b.content.clone())
        .unwrap_or_default();
    let merged = format!("<div>{}</div>{}", html_escape(comment), original);
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
        graph
            .post_action(
                &format!("me/messages/{draft_id}/attachments"),
                &json!({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att.name,
                    "contentBytes": base64_encode(&att.bytes),
                }),
            )
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
