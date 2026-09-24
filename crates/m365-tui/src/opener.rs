//! Opening links in the user's browser.

use std::process::{Command, Stdio};

use anyhow::{Context, Result};

fn validate_url(url: &str) -> Result<()> {
    // Only hand off things that look like links, never arbitrary shell input.
    if !(url.starts_with("http://") || url.starts_with("https://") || url.starts_with("mailto:")) {
        anyhow::bail!("refusing to open non-http(s) link");
    }
    Ok(())
}

/// Open `url` with a specific executable. The URL is passed directly as one
/// argument; no shell is involved.
pub fn open_url_with(url: &str, executable: &str) -> Result<()> {
    validate_url(url)?;
    let executable = executable.trim();
    if executable.is_empty() {
        anyhow::bail!("meeting opener is empty");
    }

    Command::new(executable)
        .arg(url)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .with_context(|| format!("could not launch {executable}"))?;
    Ok(())
}

/// Open `url` with the system handler. Spawns detached so the TUI never blocks
/// and the browser's own output can't scribble on the terminal.
pub fn open_url(url: &str) -> Result<()> {
    validate_url(url)?;

    const CANDIDATES: &[&str] = &["xdg-open", "open"];
    for bin in CANDIDATES {
        let spawned = Command::new(bin)
            .arg(url)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn();
        if spawned.is_ok() {
            return Ok(());
        }
    }
    Err(anyhow::anyhow!("no opener found (tried xdg-open, open)"))
        .context("could not launch a browser")
}

#[cfg(test)]
mod tests {
    use super::{open_url, open_url_with};

    #[test]
    fn rejects_non_web_schemes() {
        // Guards against a crafted href turning into a local command.
        assert!(open_url("file:///etc/passwd").is_err());
        assert!(open_url("javascript:alert(1)").is_err());
        assert!(open_url("; rm -rf /").is_err());
        assert!(open_url_with("file:///etc/passwd", "/bin/true").is_err());
        assert!(open_url_with("javascript:alert(1)", "/bin/true").is_err());
    }
}
