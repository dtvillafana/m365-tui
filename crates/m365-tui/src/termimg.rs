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
    pixel_width: u32,
    pixel_height: u32,
    cell_width: u16,
    cell_height: u16,
    max_rows: u16,
}

impl ReadyImage {
    /// Preserve aspect ratio as the pane changes width. The terminal protocol
    /// then resizes the pixels to the matching render rectangle.
    pub fn rows_for_width(&self, columns: u16) -> u16 {
        scaled_rows(
            self.pixel_width,
            self.pixel_height,
            self.cell_width,
            self.cell_height,
            columns,
            self.max_rows,
        )
    }
}

fn scaled_rows(
    pixel_width: u32,
    pixel_height: u32,
    cell_width: u16,
    cell_height: u16,
    columns: u16,
    max_rows: u16,
) -> u16 {
    let available_width = u32::from(columns.max(1)) * u32::from(cell_width.max(1));
    let scaled_height = if available_width < pixel_width {
        pixel_height
            .saturating_mul(available_width)
            .div_ceil(pixel_width.max(1))
    } else {
        pixel_height
    };
    let rows = scaled_height.div_ceil(u32::from(cell_height.max(1))) as u16;
    rows.clamp(1, max_rows.max(1))
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
        let pixel_width = dyn_img.width();
        let pixel_height = dyn_img.height();
        let protocol = self.picker.new_resize_protocol(dyn_img);
        Ok(ReadyImage {
            protocol,
            pixel_width,
            pixel_height,
            cell_width: font.0,
            cell_height: font.1,
            max_rows,
        })
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

pub fn mail_cache_key(message_id: &str, src: &str) -> String {
    format!("mail:{message_id}:{src}")
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

    #[test]
    fn image_rows_shrink_with_the_available_width() {
        assert_eq!(scaled_rows(800, 400, 10, 20, 80, 30), 20);
        assert_eq!(scaled_rows(800, 400, 10, 20, 40, 30), 10);
        assert_eq!(scaled_rows(800, 400, 10, 20, 20, 30), 5);
    }

    #[test]
    fn image_rows_respect_the_height_cap() {
        assert_eq!(scaled_rows(800, 1200, 10, 20, 80, 8), 8);
    }
}
