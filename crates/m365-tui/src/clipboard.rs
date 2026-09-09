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

#[derive(Debug)]
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

enum Probe {
    /// Binary not installed.
    Missing,
    /// Helper ran but the clipboard has no image (or an empty one).
    Empty,
    Image(String, Vec<u8>),
    Failed(String),
}

/// Read an image from the system clipboard, if one is there.
pub fn image_bytes() -> Result<(String, Vec<u8>), ImagePasteError> {
    let mut saw_helper = false;
    for probe in [from_wl_paste, from_xclip] {
        match probe() {
            Probe::Missing => {}
            Probe::Empty => saw_helper = true,
            Probe::Image(mime, bytes) => return Ok((mime, bytes)),
            Probe::Failed(e) => return Err(ImagePasteError::Failed(e)),
        }
    }
    if saw_helper {
        Err(ImagePasteError::NoImage)
    } else {
        Err(ImagePasteError::NoHelper)
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
