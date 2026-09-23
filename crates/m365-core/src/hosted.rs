//! Inline images on a Teams chat/channel message (`hostedContents`).
//!
//! Graph cannot take a raw file attachment on a chat message. Images go in the
//! body as `<img src="../hostedContents/{id}/$value">` plus a matching
//! `hostedContents` entry holding base64 bytes.

use serde_json::{json, Value};

use crate::models::ChatMessage;
use crate::util::{base64_encode, html_escape};

/// Raw bytes plus the MIME type Graph should store. The TUI validates these
/// before they get here.
#[derive(Debug, Clone)]
pub struct HostedImage {
    pub name: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

/// A Teams message body: plain text, or HTML with inline images.
#[derive(Debug, Clone, Copy)]
pub enum OutgoingBody<'a> {
    Text(&'a str),
    Html {
        html: &'a str,
        images: &'a [HostedImage],
    },
}

impl OutgoingBody<'_> {
    /// Inner HTML used when wrapping a reply (`<p>…</p>`). Plain text is
    /// escaped; HTML is passed through.
    pub fn inner_html(&self) -> String {
        match self {
            OutgoingBody::Text(text) => html_escape(text),
            OutgoingBody::Html { html, .. } => (*html).to_string(),
        }
    }
}

/// JSON body for `POST …/messages` (and channel replies).
pub fn outgoing_payload(body: OutgoingBody<'_>) -> Value {
    match body {
        OutgoingBody::Text(text) => {
            json!({ "body": { "contentType": "text", "content": text } })
        }
        OutgoingBody::Html { html, images } => {
            let mut payload = json!({
                "body": { "contentType": "html", "content": html },
            });
            if !images.is_empty() {
                payload["hostedContents"] = hosted_array(images);
            }
            payload
        }
    }
}

/// Attach `hostedContents` to an existing payload (chat replies already have
/// a `body` and `attachments`).
pub fn with_hosted_contents(mut payload: Value, images: &[HostedImage]) -> Value {
    if !images.is_empty() {
        payload["hostedContents"] = hosted_array(images);
    }
    payload
}

pub fn hosted_array(images: &[HostedImage]) -> Value {
    Value::Array(
        images
            .iter()
            .enumerate()
            .map(|(i, img)| {
                json!({
                    "@microsoft.graph.temporaryId": (i + 1).to_string(),
                    "contentBytes": base64_encode(&img.bytes),
                    "contentType": img.content_type,
                })
            })
            .collect(),
    )
}

pub fn images_of(body: OutgoingBody<'_>) -> &[HostedImage] {
    match body {
        OutgoingBody::Text(_) => &[],
        OutgoingBody::Html { images, .. } => images,
    }
}

/// JSON body for `PATCH …/messages/{id}`.
///
/// Graph treats omitted fields as cleared, so existing attachments (the quote
/// on a chat reply) and `<img>` tags have to go back with the new text.
pub fn update_payload(original: &ChatMessage, body: OutgoingBody<'_>) -> Value {
    let quote_id = original
        .quoted()
        .map(|q| q.message_id)
        .filter(|id| !id.is_empty());
    let retained = retained_img_html(&original.text());
    let images = images_of(body);
    let needs_html = quote_id.is_some()
        || !retained.is_empty()
        || !images.is_empty()
        || matches!(body, OutgoingBody::Html { .. });

    if !needs_html {
        return outgoing_payload(body);
    }

    let inner = body.inner_html();
    let mut content = String::new();
    if let Some(id) = &quote_id {
        content.push_str(&format!("<attachment id=\"{id}\"></attachment>"));
    }
    if !inner.is_empty() {
        content.push_str(&format!("<p>{inner}</p>"));
    }
    content.push_str(&retained);

    let mut payload = json!({
        "body": {
            "contentType": "html",
            "content": content,
        }
    });
    if !original.attachments.is_empty() {
        payload["attachments"] =
            serde_json::to_value(&original.attachments).expect("attachments are serializable");
    }
    with_hosted_contents(payload, images)
}

/// Existing `<img>` tags in a body, kept so a text edit does not drop images.
fn retained_img_html(html: &str) -> String {
    let lower = html.to_ascii_lowercase();
    let mut out = String::new();
    let mut from = 0;
    while let Some(rel) = lower[from..].find("<img") {
        let start = from + rel;
        let Some(rel_end) = html[start..].find('>') else {
            break;
        };
        let end = start + rel_end + 1;
        out.push_str(&html[start..end]);
        from = end;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png() -> HostedImage {
        HostedImage {
            name: "shot.png".into(),
            content_type: "image/png".into(),
            bytes: b"\x89PNG".to_vec(),
        }
    }

    #[test]
    fn text_payload_is_unchanged() {
        let v = outgoing_payload(OutgoingBody::Text("hello"));
        assert_eq!(v["body"]["contentType"], "text");
        assert_eq!(v["body"]["content"], "hello");
        assert!(v.get("hostedContents").is_none());
    }

    #[test]
    fn html_payload_carries_temporary_ids() {
        let img = png();
        let html = "<img alt=\"shot.png\" src=\"../hostedContents/1/$value\">";
        let v = outgoing_payload(OutgoingBody::Html {
            html,
            images: std::slice::from_ref(&img),
        });
        assert_eq!(v["body"]["contentType"], "html");
        assert_eq!(v["body"]["content"], html);
        assert_eq!(v["hostedContents"][0]["@microsoft.graph.temporaryId"], "1");
        assert_eq!(v["hostedContents"][0]["contentType"], "image/png");
        assert_eq!(v["hostedContents"][0]["contentBytes"], "iVBORw==");
    }

    #[test]
    fn reply_wrapper_adds_hosted_contents() {
        let img = png();
        let payload = json!({
            "body": { "contentType": "html", "content": "<p>hi</p>" },
            "attachments": [{ "contentType": "messageReference" }],
        });
        let v = with_hosted_contents(payload, std::slice::from_ref(&img));
        assert_eq!(v["hostedContents"][0]["@microsoft.graph.temporaryId"], "1");
        assert_eq!(v["attachments"][0]["contentType"], "messageReference");
    }

    #[test]
    fn image_only_message() {
        let img = png();
        let v = outgoing_payload(OutgoingBody::Html {
            html: "<img src=\"../hostedContents/1/$value\">",
            images: std::slice::from_ref(&img),
        });
        assert_eq!(v["hostedContents"].as_array().unwrap().len(), 1);
    }

    fn reply_message() -> crate::models::ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": "1785859178276",
            "body": {
                "contentType": "html",
                "content": "<attachment id=\"1785858892876\"></attachment><p>old</p><img alt=\"shot.png\" src=\"../hostedContents/1/$value\">"
            },
            "attachments": [{
                "id": "1785858892876",
                "contentType": "messageReference",
                "content": "{\"messageId\":\"1785858892876\",\"messagePreview\":\"Sounds good\",\"messageSender\":{\"user\":{\"displayName\":\"Alex\"}}}"
            }]
        }))
        .unwrap()
    }

    #[test]
    fn edit_keeps_the_quote_attachment_and_existing_images() {
        let original = reply_message();
        let v = update_payload(&original, OutgoingBody::Text("new text"));
        let content = v["body"]["content"].as_str().unwrap();
        assert!(
            content.contains("<attachment id=\"1785858892876\"></attachment>"),
            "quote tag missing: {content}"
        );
        assert!(
            content.contains("<p>new text</p>"),
            "new body missing: {content}"
        );
        assert!(
            content.contains("<img alt=\"shot.png\" src=\"../hostedContents/1/$value\">"),
            "existing image dropped: {content}"
        );
        assert_eq!(v["attachments"][0]["contentType"], "messageReference");
        assert_eq!(v["attachments"][0]["id"], "1785858892876");
    }

    #[test]
    fn plain_edit_stays_plain_text() {
        let original: crate::models::ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "1",
            "body": { "contentType": "text", "content": "hello" }
        }))
        .unwrap();
        let v = update_payload(&original, OutgoingBody::Text("hello world"));
        assert_eq!(v["body"]["contentType"], "text");
        assert_eq!(v["body"]["content"], "hello world");
        assert!(v.get("attachments").is_none());
    }
}
