//! Inline images for the Teams composer: `@path` tokens, Tab completion,
//! clipboard / data-URI paste, and the HTML Graph payload.

use std::fs;
use std::path::{Path, PathBuf};

use m365_core::hosted::HostedImage;
use m365_core::util::{base64_decode, html_escape, html_escape_attr};

/// Graph JSON bodies blow up once base64 is applied; keep a hard cap.
pub const MAX_IMAGE_BYTES: u64 = 3 * 1024 * 1024;

#[derive(Debug, Clone)]
pub struct InlineImage {
    pub name: String,
    pub mime: &'static str,
    pub bytes: Vec<u8>,
}

impl InlineImage {
    pub fn from_bytes(name: String, bytes: Vec<u8>) -> Result<Self, String> {
        let kind = image_kind(&bytes)
            .ok_or_else(|| "not an image (PNG, JPEG, GIF or WebP)".to_string())?;
        if bytes.len() as u64 > MAX_IMAGE_BYTES {
            return Err(format!(
                "{} is {}; Teams inline images are capped at {}",
                name,
                human_size(bytes.len() as u64),
                human_size(MAX_IMAGE_BYTES)
            ));
        }
        let name = if name.rsplit('.').next() == Some(kind.ext) {
            name
        } else {
            match Path::new(&name).file_stem().and_then(|s| s.to_str()) {
                Some(stem) if !stem.is_empty() => format!("{stem}.{}", kind.ext),
                _ => format!("image.{}", kind.ext),
            }
        };
        Ok(Self {
            name,
            mime: kind.mime,
            bytes,
        })
    }

    pub fn from_path(path: &Path) -> Result<Self, String> {
        let md = fs::metadata(path).map_err(|e| format!("{}: {e}", path.display()))?;
        if md.is_dir() {
            return Err(format!("{} is a directory, not an image", path.display()));
        }
        if !md.is_file() {
            return Err(format!("{} is not a file", path.display()));
        }
        if md.len() > MAX_IMAGE_BYTES {
            return Err(format!(
                "{} is {}; Teams inline images are capped at {}",
                path.display(),
                human_size(md.len()),
                human_size(MAX_IMAGE_BYTES)
            ));
        }
        let bytes = fs::read(path).map_err(|e| format!("{}: {e}", path.display()))?;
        let name = path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("image")
            .to_string();
        Self::from_bytes(name, bytes)
    }

    pub fn into_hosted(self) -> HostedImage {
        HostedImage {
            name: self.name,
            content_type: self.mime.to_string(),
            bytes: self.bytes,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct ImageKind {
    mime: &'static str,
    ext: &'static str,
}

pub fn image_extension(bytes: &[u8]) -> Option<&'static str> {
    image_kind(bytes).map(|k| k.ext)
}

fn image_kind(bytes: &[u8]) -> Option<ImageKind> {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) {
        Some(ImageKind {
            mime: "image/png",
            ext: "png",
        })
    } else if bytes.len() >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF {
        Some(ImageKind {
            mime: "image/jpeg",
            ext: "jpg",
        })
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        Some(ImageKind {
            mime: "image/gif",
            ext: "gif",
        })
    } else if bytes.len() >= 12 && bytes.starts_with(b"RIFF") && &bytes[8..12] == b"WEBP" {
        Some(ImageKind {
            mime: "image/webp",
            ext: "webp",
        })
    } else {
        None
    }
}

fn image_ext(name: &str) -> bool {
    matches!(
        Path::new(name)
            .extension()
            .and_then(|s| s.to_str())
            .map(|s| s.to_ascii_lowercase())
            .as_deref(),
        Some("png" | "jpg" | "jpeg" | "gif" | "webp")
    )
}

// ---------------------------------------------------------------------------
// @path tokens
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PathToken {
    /// Character index of the `@`.
    pub start: usize,
    /// Character index after the last path character.
    pub end: usize,
    /// Path with backslash escapes still present.
    pub raw: String,
    /// Path with escapes undone, tilde not expanded.
    pub unescaped: String,
}

pub fn looks_like_path(unescaped: &str) -> bool {
    unescaped.starts_with('/')
        || unescaped == "~"
        || unescaped.starts_with("~/")
        || unescaped == "."
        || unescaped == ".."
        || unescaped.starts_with("./")
        || unescaped.starts_with("../")
        || unescaped.contains('/')
}

pub fn unescape_path(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some(n) => out.push(n),
                None => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

pub fn escape_path(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        if c.is_whitespace() || c == '\\' {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

pub fn expand_tilde(p: &str) -> PathBuf {
    if p == "~" {
        return home_dir().unwrap_or_else(|| PathBuf::from("~"));
    }
    if let Some(rest) = p.strip_prefix("~/") {
        if let Some(home) = home_dir() {
            return if rest.is_empty() {
                home
            } else {
                home.join(rest)
            };
        }
    }
    PathBuf::from(p)
}

fn home_dir() -> Option<PathBuf> {
    std::env::var_os("HOME").map(PathBuf::from)
}

/// Every `@path` token in `chars` that looks like a filesystem path.
pub fn path_tokens(chars: &[char]) -> Vec<PathToken> {
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '@' && (i == 0 || chars[i - 1].is_whitespace()) {
            let (end, raw) = parse_path_body(chars, i + 1);
            let unescaped = unescape_path(&raw);
            if looks_like_path(&unescaped) {
                out.push(PathToken {
                    start: i,
                    end,
                    raw,
                    unescaped,
                });
            }
            i = end.max(i + 1);
        } else {
            i += 1;
        }
    }
    out
}

fn parse_path_body(chars: &[char], start: usize) -> (usize, String) {
    let mut raw = String::new();
    let mut i = start;
    while i < chars.len() {
        match chars[i] {
            '\\' => {
                raw.push('\\');
                i += 1;
                if i < chars.len() {
                    raw.push(chars[i]);
                    i += 1;
                }
            }
            c if c.is_whitespace() => break,
            c => {
                raw.push(c);
                i += 1;
            }
        }
    }
    (i, raw)
}

/// The path token the cursor is inside, if any.
pub fn token_at(chars: &[char], cursor: usize) -> Option<PathToken> {
    path_tokens(chars)
        .into_iter()
        .find(|t| t.start <= cursor && cursor <= t.end)
}

// ---------------------------------------------------------------------------
// Tab completion
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Completion {
    /// Full replacement including the leading `@`.
    pub replacement: String,
    pub status: String,
}

/// Complete the path token. `None` if there is nothing to do.
pub fn complete(token: &PathToken) -> Option<Completion> {
    let unescaped = &token.unescaped;
    let (dir_user, prefix) = split_dir_prefix(unescaped);
    let dir_path = expand_tilde(&dir_user);
    let entries = match fs::read_dir(&dir_path) {
        Ok(rd) => rd,
        Err(_) => {
            return Some(Completion {
                replacement: format!("@{}", token.raw),
                status: format!("no such directory {}", dir_path.display()),
            });
        }
    };

    let show_hidden = prefix.starts_with('.');
    let mut dirs = Vec::new();
    let mut files = Vec::new();
    for ent in entries.flatten() {
        let name = ent.file_name();
        let Some(name) = name.to_str() else { continue };
        if name == "." || name == ".." {
            continue;
        }
        if !show_hidden && name.starts_with('.') {
            continue;
        }
        if !name.starts_with(prefix) {
            continue;
        }
        let is_dir = ent.file_type().map(|t| t.is_dir()).unwrap_or(false);
        if is_dir {
            dirs.push(name.to_string());
        } else if image_ext(name) {
            files.push(name.to_string());
        }
    }
    dirs.sort();
    files.sort();
    let matches: Vec<(String, bool)> = dirs
        .into_iter()
        .map(|n| (n, true))
        .chain(files.into_iter().map(|n| (n, false)))
        .collect();

    if matches.is_empty() {
        return Some(Completion {
            replacement: format!("@{}", token.raw),
            status: "no image or directory matches".into(),
        });
    }

    let names: Vec<&str> = matches.iter().map(|(n, _)| n.as_str()).collect();
    let common = common_prefix(&names);
    let unique = matches.len() == 1;
    let completed_name = if unique {
        &matches[0].0
    } else if common.len() > prefix.len() {
        common.as_str()
    } else {
        prefix
    };
    let slash = unique && matches[0].1;
    let mut completed = format!("{dir_user}{completed_name}");
    if slash && !completed.ends_with('/') {
        completed.push('/');
    }
    let replacement = format!("@{}", escape_path(&completed));
    let status = if unique {
        if slash {
            format!("completed {}", matches[0].0)
        } else {
            format!("attached path {}", matches[0].0)
        }
    } else {
        list_candidates(&matches)
    };
    Some(Completion {
        replacement,
        status,
    })
}

fn split_dir_prefix(unescaped: &str) -> (String, &str) {
    if unescaped.ends_with('/') {
        (unescaped.to_string(), "")
    } else if let Some((dir, name)) = unescaped.rsplit_once('/') {
        (format!("{dir}/"), name)
    } else {
        (String::new(), unescaped)
    }
}

fn common_prefix(names: &[&str]) -> String {
    let Some(first) = names.first() else {
        return String::new();
    };
    let mut prefix: Vec<char> = first.chars().collect();
    for name in &names[1..] {
        let chars: Vec<char> = name.chars().collect();
        let n = prefix
            .iter()
            .zip(chars.iter())
            .take_while(|(a, b)| a == b)
            .count();
        prefix.truncate(n);
        if prefix.is_empty() {
            break;
        }
    }
    prefix.into_iter().collect()
}

fn list_candidates(matches: &[(String, bool)]) -> String {
    const MAX: usize = 8;
    let mut parts: Vec<String> = matches
        .iter()
        .take(MAX)
        .map(
            |(n, is_dir)| {
                if *is_dir {
                    format!("{n}/")
                } else {
                    n.clone()
                }
            },
        )
        .collect();
    if matches.len() > MAX {
        parts.push(format!("… {} more", matches.len() - MAX));
    }
    parts.join("  ")
}

// ---------------------------------------------------------------------------
// Data URIs
// ---------------------------------------------------------------------------

/// `None` if `s` is not a data URI. `Some(Err)` if it looks like one but
/// isn't a usable image.
pub fn from_data_uri(s: &str) -> Option<Result<InlineImage, String>> {
    let s = s.trim();
    let rest = s.strip_prefix("data:")?;
    let (meta, data) = rest.split_once(',')?;
    if !meta.starts_with("image/") {
        return Some(Err("data URI is not an image".into()));
    }
    if !meta.split(';').any(|p| p.eq_ignore_ascii_case("base64")) {
        return Some(Err("image data URI must be base64".into()));
    }
    let bytes = match base64_decode(data) {
        Ok(b) => b,
        Err(e) => return Some(Err(format!("invalid data URI ({e})"))),
    };
    Some(InlineImage::from_bytes("paste".into(), bytes))
}

// ---------------------------------------------------------------------------
// Compose → Graph
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub enum Prepared {
    Text(String),
    Html {
        html: String,
        images: Vec<HostedImage>,
    },
}

/// Build the outgoing Teams body from composer text and staged clipboard
/// images. `Ok(None)` means there is nothing to send.
pub fn prepare_message(text: &str, staged: &[InlineImage]) -> Result<Option<Prepared>, String> {
    let chars: Vec<char> = text.chars().collect();
    let tokens = path_tokens(&chars);
    if tokens.is_empty() && staged.is_empty() {
        let trimmed = text.trim();
        if trimmed.is_empty() {
            return Ok(None);
        }
        return Ok(Some(Prepared::Text(trimmed.to_string())));
    }

    let mut html = String::new();
    let mut images: Vec<HostedImage> = Vec::new();
    let mut i = 0;
    for token in &tokens {
        if token.start > i {
            html.push_str(&html_escape(
                &chars[i..token.start].iter().collect::<String>(),
            ));
        }
        let path = expand_tilde(&token.unescaped);
        let img = InlineImage::from_path(&path)?;
        images.push(img.into_hosted());
        let id = images.len();
        let last = images.last().unwrap();
        html.push_str(&format!(
            "<img alt=\"{}\" src=\"../hostedContents/{id}/$value\">",
            html_escape_attr(&last.name)
        ));
        i = token.end;
    }
    if i < chars.len() {
        html.push_str(&html_escape(&chars[i..].iter().collect::<String>()));
    }
    for img in staged {
        images.push(img.clone().into_hosted());
        let id = images.len();
        html.push_str(&format!(
            "<img alt=\"{}\" src=\"../hostedContents/{id}/$value\">",
            html_escape_attr(&img.name)
        ));
    }

    let total: u64 = images.iter().map(|i| i.bytes.len() as u64).sum();
    if total > MAX_IMAGE_BYTES {
        return Err(format!(
            "images total {}; capped at {} per message",
            human_size(total),
            human_size(MAX_IMAGE_BYTES)
        ));
    }
    if html.trim().is_empty() && images.is_empty() {
        return Ok(None);
    }
    Ok(Some(Prepared::Html { html, images }))
}

pub fn unique_paste_name(existing: &[InlineImage], ext: &str) -> String {
    let base = format!("clipboard.{ext}");
    if existing.iter().all(|i| i.name != base) {
        return base;
    }
    for n in 2..1000 {
        let name = format!("clipboard-{n}.{ext}");
        if existing.iter().all(|i| i.name != name) {
            return name;
        }
    }
    format!("clipboard-new.{ext}")
}

pub fn human_size(bytes: u64) -> String {
    if bytes >= 1024 * 1024 {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    } else if bytes >= 1024 {
        format!("{:.0} KB", bytes as f64 / 1024.0)
    } else {
        format!("{bytes} B")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 1×1 PNG, valid enough for magic-byte detection and round-trips.
    const TINY_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0x2C, 0x0D, 0x76, 0x8B, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    fn chars(s: &str) -> Vec<char> {
        s.chars().collect()
    }

    #[test]
    fn detects_magic_bytes() {
        assert_eq!(image_kind(TINY_PNG).unwrap().mime, "image/png");
        assert_eq!(image_kind(&[0xFF, 0xD8, 0xFF, 0xE0]).unwrap().ext, "jpg");
        assert_eq!(image_kind(b"GIF89a....").unwrap().mime, "image/gif");
        let mut webp = b"RIFF".to_vec();
        webp.extend_from_slice(&[0, 0, 0, 0]);
        webp.extend_from_slice(b"WEBP");
        assert_eq!(image_kind(&webp).unwrap().mime, "image/webp");
        assert!(image_kind(b"%PDF").is_none());
    }

    #[test]
    fn unescapes_backslash_spaces() {
        assert_eq!(unescape_path(r"/tmp/My\ Image.png"), "/tmp/My Image.png");
        assert_eq!(unescape_path(r"/tmp/foo\\bar.png"), r"/tmp/foo\bar.png");
        assert_eq!(escape_path("/tmp/My Image.png"), r"/tmp/My\ Image.png");
    }

    #[test]
    fn path_tokens_need_a_path_shape() {
        let t = path_tokens(&chars("hello @alice there"));
        assert!(t.is_empty(), "mentions are not paths");
        let t = path_tokens(&chars("see @~/Pictures/a.png please"));
        assert_eq!(t.len(), 1);
        assert_eq!(t[0].unescaped, "~/Pictures/a.png");
        let t = path_tokens(&chars("@./shot.png"));
        assert_eq!(t[0].unescaped, "./shot.png");
        let t = path_tokens(&chars(r"x @/tmp/My\ Image.png y"));
        assert_eq!(t[0].unescaped, "/tmp/My Image.png");
        let t = path_tokens(&chars("a@b.com"));
        assert!(t.is_empty());
        let t = path_tokens(&chars("@dir/foo.png"));
        assert_eq!(t[0].unescaped, "dir/foo.png");
    }

    #[test]
    fn token_at_cursor_includes_the_end() {
        let c = chars("@./a.png");
        assert!(token_at(&c, 0).is_some());
        assert!(token_at(&c, c.len()).is_some());
        assert!(token_at(&chars("hello"), 3).is_none());
    }

    #[test]
    fn tilde_expands_to_home() {
        std::env::set_var("HOME", "/home/tester");
        assert_eq!(expand_tilde("~/a.png"), PathBuf::from("/home/tester/a.png"));
        assert_eq!(expand_tilde("~"), PathBuf::from("/home/tester"));
        assert_eq!(expand_tilde("/abs"), PathBuf::from("/abs"));
    }

    #[test]
    fn data_uri_decodes_png() {
        let b64 = m365_core::util::base64_encode(TINY_PNG);
        let uri = format!("data:image/png;base64,{b64}");
        let img = from_data_uri(&uri).unwrap().unwrap();
        assert_eq!(img.mime, "image/png");
        assert_eq!(img.bytes, TINY_PNG);
        assert!(from_data_uri("hello").is_none());
        assert!(from_data_uri("data:image/png,notbase64").unwrap().is_err());
        assert!(from_data_uri("data:text/plain;base64,Zg==")
            .unwrap()
            .is_err());
    }

    #[test]
    fn rejects_non_image_bytes() {
        assert!(InlineImage::from_bytes("x.pdf".into(), b"%PDF".to_vec()).is_err());
    }

    #[test]
    fn prepare_plain_text_skips_html() {
        match prepare_message("hello", &[]).unwrap() {
            Some(Prepared::Text(t)) => assert_eq!(t, "hello"),
            _ => panic!("expected plain text"),
        }
        assert!(prepare_message("   ", &[]).unwrap().is_none());
    }

    #[test]
    fn prepare_replaces_path_tokens_in_place() {
        let dir = std::env::temp_dir().join(format!("m365-img-prepare-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        let file = dir.join("shot.png");
        fs::write(&file, TINY_PNG).unwrap();
        let text = format!("see @{}/shot.png please", dir.display());
        match prepare_message(&text, &[]).unwrap() {
            Some(Prepared::Html { html, images }) => {
                assert_eq!(images.len(), 1);
                assert_eq!(images[0].name, "shot.png");
                assert!(html.contains("see "));
                assert!(html.contains("please"));
                assert!(html.contains("../hostedContents/1/$value"));
                assert!(html.contains("alt=\"shot.png\""));
                assert!(!html.contains("shot.png please") || html.contains("<img"));
            }
            _ => panic!("expected html"),
        }
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn prepare_rejects_non_images() {
        let dir = std::env::temp_dir().join(format!("m365-img-pdf-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        let file = dir.join("doc.pdf");
        fs::write(&file, b"%PDF").unwrap();
        let text = format!("@{}/doc.pdf", dir.display());
        let err = prepare_message(&text, &[]).unwrap_err();
        assert!(err.contains("not an image"), "{err}");
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn prepare_appends_staged_and_allows_image_only() {
        let img = InlineImage::from_bytes("clipboard.png".into(), TINY_PNG.to_vec()).unwrap();
        match prepare_message("", &[img]).unwrap() {
            Some(Prepared::Html { html, images }) => {
                assert_eq!(images.len(), 1);
                assert!(html.contains("../hostedContents/1/$value"));
            }
            _ => panic!("expected html"),
        }
    }

    #[test]
    fn complete_unique_file_and_common_prefix() {
        let dir = std::env::temp_dir().join(format!("m365-img-comp-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(dir.join("Pictures")).unwrap();
        fs::write(dir.join("Picture.png"), TINY_PNG).unwrap();
        fs::write(dir.join("notes.txt"), b"no").unwrap();

        let token = PathToken {
            start: 0,
            end: 1,
            raw: format!("{}/Pic", dir.display()),
            unescaped: format!("{}/Pic", dir.display()),
        };
        let c = complete(&token).unwrap();
        assert!(c.replacement.contains("Picture"), "{}", c.replacement);
        assert!(
            c.status.contains("Pictures/") && c.status.contains("Picture.png"),
            "{}",
            c.status
        );

        let token = PathToken {
            start: 0,
            end: 1,
            raw: format!("{}/Picture.png", dir.display()),
            unescaped: format!("{}/Picture.png", dir.display()),
        };
        let c = complete(&token).unwrap();
        assert!(c.replacement.ends_with("Picture.png"), "{}", c.replacement);

        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn complete_escapes_spaces() {
        let dir = std::env::temp_dir().join(format!("m365-img-space-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        fs::write(dir.join("My Image.png"), TINY_PNG).unwrap();
        let token = PathToken {
            start: 0,
            end: 1,
            raw: format!("{}/My", dir.display()),
            unescaped: format!("{}/My", dir.display()),
        };
        let c = complete(&token).unwrap();
        assert!(
            c.replacement.contains(r"My\ Image.png"),
            "{}",
            c.replacement
        );
        let _ = fs::remove_dir_all(&dir);
    }
}
