//! Copying text to the system clipboard, and reading image bytes back.
//!
//! Prefers a native helper (`wl-copy`, `xclip`, `xsel`) when one is on `PATH`,
//! since those work regardless of terminal support. Otherwise falls back to the
//! OSC 52 escape sequence, which most modern terminals honour and which also
//! works over SSH.
//!
//! Image paste (`Ctrl+V`) reads `image/*` from `wl-paste` or `xclip`. Bracketed
//! paste is UTF-8 text only, so raw image bytes never arrive through the
//! terminal — the helpers are the only way.

use std::io::{Read, Write};
use std::process::{Command, Stdio};

use anyhow::{Context, Result};

const IMAGE_TYPES: &[&str] = &[
    "image/png",
    "image/jpeg",
    "image/jpg",
    "image/gif",
    "image/webp",
];

#[derive(Debug, PartialEq, Eq)]
pub enum ImagePasteError {
    /// Neither `wl-paste` nor `xclip` is available.
    NoHelper,
    /// A helper is present but the clipboard has no image.
    NoImage,
    Failed(String),
}

/// Copy `text` to the clipboard. Returns the mechanism used, for the status line.
pub fn copy(text: &str) -> Result<&'static str> {
    if let Some(tool) = via_helper(text) {
        return Ok(tool);
    }
    via_osc52(text)?;
    Ok("OSC 52")
}

fn via_helper(text: &str) -> Option<&'static str> {
    const CANDIDATES: &[(&str, &[&str])] = &[
        ("wl-copy", &[]),
        ("xclip", &["-selection", "clipboard"]),
        ("xsel", &["--clipboard", "--input"]),
    ];

    for (bin, args) in CANDIDATES {
        let Ok(mut child) = Command::new(bin)
            .args(*args)
            .stdin(Stdio::piped())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
        else {
            continue; // not installed
        };
        // Write then drop stdin so the helper sees EOF.
        if let Some(mut stdin) = child.stdin.take() {
            if stdin.write_all(text.as_bytes()).is_err() {
                let _ = child.kill();
                continue;
            }
        }
        // wl-copy daemonizes to hold the selection, so this returns promptly.
        let _ = child.wait();
        return Some(bin);
    }
    None
}

/// Terminal-native clipboard write. Safe to emit while in raw mode.
fn via_osc52(text: &str) -> Result<()> {
    let mut out = std::io::stdout();
    write!(
        out,
        "\x1b]52;c;{}\x07",
        m365_core::util::base64_encode(text.as_bytes())
    )
    .context("writing OSC 52 sequence")?;
    out.flush().context("flushing OSC 52 sequence")?;
    Ok(())
}

/// Read an image from the system clipboard, if one is there.
///
/// `wl-paste` is tried first; if it is missing, errors, or has no image,
/// `xclip` is tried next. Failure is only reported after both have been tried.
pub fn image_bytes() -> Result<(String, Vec<u8>), ImagePasteError> {
    let wl = from_wl_paste();
    if let Probe::Image(mime, bytes) = wl {
        return Ok((mime, bytes));
    }
    combine_probes(wl, from_xclip())
}

#[derive(Debug)]
pub(crate) enum Probe {
    /// Binary not installed.
    Missing,
    /// Helper ran but the clipboard has no image (or an empty one).
    Empty,
    Image(String, Vec<u8>),
    Failed(String),
}

pub(crate) fn combine_probes(
    wl: Probe,
    xclip: Probe,
) -> Result<(String, Vec<u8>), ImagePasteError> {
    if let Probe::Image(mime, bytes) = wl {
        return Ok((mime, bytes));
    }
    if let Probe::Image(mime, bytes) = xclip {
        return Ok((mime, bytes));
    }
    let wl_missing = matches!(wl, Probe::Missing);
    let x_missing = matches!(xclip, Probe::Missing);
    if wl_missing && x_missing {
        return Err(ImagePasteError::NoHelper);
    }
    if matches!(wl, Probe::Empty) || matches!(xclip, Probe::Empty) {
        return Err(ImagePasteError::NoImage);
    }
    let mut parts = Vec::new();
    if let Probe::Failed(e) = wl {
        parts.push(format!("wl-paste: {e}"));
    }
    if let Probe::Failed(e) = xclip {
        parts.push(format!("xclip: {e}"));
    }
    if parts.is_empty() {
        Err(ImagePasteError::NoImage)
    } else {
        Err(ImagePasteError::Failed(parts.join("; ")))
    }
}

fn from_wl_paste() -> Probe {
    read_image("wl-paste", &["-l"], |mime| vec!["--type", mime])
}

fn from_xclip() -> Probe {
    read_image(
        "xclip",
        &["-selection", "clipboard", "-t", "TARGETS", "-o"],
        |mime| vec!["-selection", "clipboard", "-t", mime, "-o"],
    )
}

fn read_image(bin: &str, list_args: &[&str], get_args: impl Fn(&str) -> Vec<&str>) -> Probe {
    let listed = match run_stdout(bin, list_args) {
        None => return Probe::Missing,
        Some(Err(e)) => return Probe::Failed(e),
        Some(Ok(bytes)) => String::from_utf8_lossy(&bytes).into_owned(),
    };
    let Some(mime) = pick_mime(listed.lines()) else {
        return Probe::Empty;
    };
    match run_stdout(bin, &get_args(mime)) {
        Some(Ok(bytes)) if !bytes.is_empty() => Probe::Image(mime.to_string(), bytes),
        Some(Err(e)) => Probe::Failed(e),
        _ => Probe::Empty,
    }
}

fn pick_mime<'a>(lines: impl Iterator<Item = &'a str>) -> Option<&'static str> {
    let available: Vec<&str> = lines.map(str::trim).filter(|s| !s.is_empty()).collect();
    IMAGE_TYPES
        .iter()
        .copied()
        .find(|mime| available.iter().any(|a| a.eq_ignore_ascii_case(mime)))
}

/// `None` if the binary is not installed; `Some(Err)` if it ran and failed.
fn run_stdout(bin: &str, args: &[&str]) -> Option<Result<Vec<u8>, String>> {
    let mut child = Command::new(bin)
        .args(args)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;
    let mut buf = Vec::new();
    if let Some(mut stdout) = child.stdout.take() {
        if stdout.read_to_end(&mut buf).is_err() {
            let _ = child.kill();
            return Some(Err(format!("{bin} read failed")));
        }
    }
    match child.wait() {
        Ok(st) if st.success() => Some(Ok(buf)),
        Ok(_) => Some(Err(format!("{bin} exited with an error"))),
        Err(e) => Some(Err(format!("{bin}: {e}"))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xclip_is_used_when_wl_paste_fails() {
        let got = combine_probes(
            Probe::Failed("no wayland".into()),
            Probe::Image("image/png".into(), b"\x89PNG".to_vec()),
        )
        .unwrap();
        assert_eq!(got.0, "image/png");
        assert_eq!(got.1, b"\x89PNG");
    }

    #[test]
    fn xclip_is_used_when_wl_paste_is_missing() {
        let got = combine_probes(
            Probe::Missing,
            Probe::Image("image/jpeg".into(), vec![0xFF, 0xD8, 0xFF]),
        )
        .unwrap();
        assert_eq!(got.0, "image/jpeg");
    }

    #[test]
    fn xclip_is_used_when_wl_paste_has_no_image() {
        let got = combine_probes(
            Probe::Empty,
            Probe::Image("image/png".into(), b"png".to_vec()),
        )
        .unwrap();
        assert_eq!(got.1, b"png");
    }

    #[test]
    fn no_helper_only_when_both_are_missing() {
        assert_eq!(
            combine_probes(Probe::Missing, Probe::Missing).unwrap_err(),
            ImagePasteError::NoHelper
        );
    }

    #[test]
    fn no_image_when_a_helper_ran_and_found_none() {
        assert_eq!(
            combine_probes(Probe::Empty, Probe::Missing).unwrap_err(),
            ImagePasteError::NoImage
        );
        assert_eq!(
            combine_probes(Probe::Failed("err".into()), Probe::Empty).unwrap_err(),
            ImagePasteError::NoImage
        );
    }

    #[test]
    fn combines_failures_from_both_helpers() {
        let err = combine_probes(Probe::Failed("wayland".into()), Probe::Failed("x11".into()))
            .unwrap_err();
        match err {
            ImagePasteError::Failed(s) => {
                assert!(s.contains("wl-paste"), "{s}");
                assert!(s.contains("xclip"), "{s}");
            }
            other => panic!("expected Failed, got {other:?}"),
        }
    }

    #[test]
    fn wl_paste_success_does_not_need_xclip() {
        let got = combine_probes(
            Probe::Image("image/gif".into(), b"GIF89a".to_vec()),
            Probe::Failed("should not be used".into()),
        )
        .unwrap();
        assert_eq!(got.0, "image/gif");
    }
}
