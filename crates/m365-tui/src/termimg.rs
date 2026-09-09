//! Terminal-protocol image rendering (Kitty, Sixel, iTerm2).
//!
//! Unicode half-blocks are not used: if the terminal cannot display pixels,
//! callers show a short `[image unavailable]` label instead.

use std::io::Cursor;

use image::ImageReader;
use ratatui::layout::Rect;
use ratatui::Frame;
use ratatui_image::picker::{Picker, ProtocolType};
use ratatui_image::protocol::StatefulProtocol;
use ratatui_image::StatefulImage;

pub const CONVO_IMAGE_ROWS: u16 = 8;
pub const COMPOSER_IMAGE_ROWS: u16 = 5;

pub struct Graphics {
    picker: Picker,
}

pub struct ReadyImage {
    pub protocol: StatefulProtocol,
    pub rows: u16,
}

impl Graphics {
    /// Query the terminal. `None` if there is no pixel graphics protocol.
    pub fn detect() -> Option<Self> {
        let picker = Picker::from_query_stdio().ok()?;
        if picker.protocol_type() == ProtocolType::Halfblocks {
            return None;
        }
        Some(Self { picker })
    }

    pub fn decode(&self, bytes: &[u8], max_rows: u16) -> Result<ReadyImage, String> {
        let reader = ImageReader::new(Cursor::new(bytes))
            .with_guessed_format()
            .map_err(|e| e.to_string())?;
        let dyn_img = reader.decode().map_err(|e| e.to_string())?;
        let font = self.picker.font_size();
        let cell_h = u32::from(font.1.max(1));
        let natural = dyn_img.height().div_ceil(cell_h) as u16;
        let rows = natural.clamp(1, max_rows.max(1));
        let protocol = self.picker.new_resize_protocol(dyn_img);
        Ok(ReadyImage { protocol, rows })
    }
}

pub fn render(f: &mut Frame, area: Rect, image: &mut ReadyImage) {
    if area.width == 0 || area.height == 0 {
        return;
    }
    f.render_stateful_widget(StatefulImage::default(), area, &mut image.protocol);
}

pub fn cache_key(src: &str, hosted_id: Option<&str>) -> String {
    if let Some(id) = hosted_id.filter(|s| !s.is_empty()) {
        format!("hosted:{id}")
    } else {
        src.to_string()
    }
}

/// Pull a Graph hosted-content id out of an `<img src>`.
pub fn hosted_content_id(src: &str) -> Option<&str> {
    if let Some(rest) = src.split("hostedContents('").nth(1) {
        let id = rest.split("')").next()?;
        if !id.is_empty() {
            return Some(id);
        }
    }
    if let Some(rest) = src.split("hostedContents/").nth(1) {
        let id = rest
            .split('/')
            .next()?
            .trim_end_matches("$value")
            .trim_end_matches(')');
        if !id.is_empty() {
            return Some(id);
        }
    }
    None
}

pub fn fetch_path(
    src: &str,
    hosted_id: Option<&str>,
    message_id: &str,
    chat_id: Option<&str>,
    channel: Option<&(String, String)>,
) -> Option<String> {
    let src = src.trim();
    if src.starts_with("http://") || src.starts_with("https://") {
        return Some(src.to_string());
    }
    let id = hosted_id.or_else(|| hosted_content_id(src))?;
    if let Some(chat) = chat_id {
        return Some(format!(
            "me/chats/{chat}/messages/{message_id}/hostedContents/{id}/$value"
        ));
    }
    if let Some((team, channel)) = channel {
        return Some(format!(
            "teams/{team}/channels/{channel}/messages/{message_id}/hostedContents/{id}/$value"
        ));
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_graph_hosted_content_urls() {
        assert_eq!(
            hosted_content_id(
                "https://graph.microsoft.com/v1.0/chats('19:x')/messages('1')/hostedContents('abc123')/$value"
            ),
            Some("abc123")
        );
        assert_eq!(hosted_content_id("../hostedContents/1/$value"), Some("1"));
        assert_eq!(
            hosted_content_id(
                "https://graph.microsoft.com/v1.0/chats/19:x/messages/1/hostedContents/xyz/$value"
            ),
            Some("xyz")
        );
        assert!(hosted_content_id("https://example.com/pic.png").is_none());
    }

    #[test]
    fn fetch_path_prefers_absolute_src() {
        let url = "https://graph.microsoft.com/v1.0/chats/c/messages/m/hostedContents/h/$value";
        assert_eq!(
            fetch_path(url, Some("h"), "m", Some("c"), None).as_deref(),
            Some(url)
        );
    }

    #[test]
    fn fetch_path_builds_relative_chat_and_channel() {
        assert_eq!(
            fetch_path("../hostedContents/1/$value", None, "mid", Some("cid"), None),
            Some("me/chats/cid/messages/mid/hostedContents/1/$value".into())
        );
        assert_eq!(
            fetch_path(
                "../hostedContents/9/$value",
                None,
                "mid",
                None,
                Some(&("t".into(), "ch".into())),
            ),
            Some("teams/t/channels/ch/messages/mid/hostedContents/9/$value".into())
        );
    }
}
