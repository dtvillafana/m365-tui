//! Inline images on a Teams chat/channel message (`hostedContents`).
//!
//! Graph cannot take a raw file attachment on a chat message. Images go in the
//! body as `<img src="../hostedContents/{id}/$value">` plus a matching
//! `hostedContents` entry holding base64 bytes.

use serde_json::{json, Value};

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
}
