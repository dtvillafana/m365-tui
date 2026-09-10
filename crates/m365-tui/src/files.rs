//! Saving attachments to disk.
//!
//! Attachment names come from email, i.e. from anyone who can send you mail, so
//! they are treated as hostile: the name is reduced to a bare file name and can
//! never escape the download directory.

use std::path::{Path, PathBuf};

use anyhow::{Context, Result};

/// Where downloads go: the user's Download dir, else the temp dir.
pub fn download_dir() -> PathBuf {
    directories::UserDirs::new()
        .and_then(|u| u.download_dir().map(Path::to_path_buf))
        .unwrap_or_else(std::env::temp_dir)
}

/// Write `bytes` into the download directory under a safe form of `name`,
/// without overwriting an existing file. Returns the path written.
pub fn save(name: &str, bytes: &[u8]) -> Result<PathBuf> {
    let dir = download_dir();
    std::fs::create_dir_all(&dir).with_context(|| format!("creating {}", dir.display()))?;
    let path = unique_path(&dir, &safe_file_name(name));
    std::fs::write(&path, bytes).with_context(|| format!("writing {}", path.display()))?;
    Ok(path)
}

/// Reduce an arbitrary attachment name to a bare, safe file name.
/// Strips any directory components, so `../../.ssh/authorized_keys` becomes
/// `authorized_keys` and can only ever land inside the download directory.
pub fn safe_file_name(name: &str) -> String {
    let base = name.rsplit(['/', '\\']).next().unwrap_or_default();
    let cleaned: String = base
        .chars()
        .filter(|c| !c.is_control() && *c != '\0')
        .collect();
    // A name of only dots (".", "..") is not a usable file name.
    let cleaned = cleaned.trim();
    if cleaned.is_empty() || cleaned.chars().all(|c| c == '.') {
        return "attachment".to_string();
    }
    cleaned.to_string()
}

/// `report.pdf` -> `report (1).pdf` if the name is taken.
fn unique_path(dir: &Path, name: &str) -> PathBuf {
    let candidate = dir.join(name);
    if !candidate.exists() {
        return candidate;
    }
    let path = Path::new(name);
    let stem = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("attachment");
    let ext = path.extension().and_then(|s| s.to_str());
    for n in 1..1000 {
        let next = match ext {
            Some(e) => dir.join(format!("{stem} ({n}).{e}")),
            None => dir.join(format!("{stem} ({n})")),
        };
        if !next.exists() {
            return next;
        }
    }
    candidate
}

#[cfg(test)]
mod tests {
    use super::{safe_file_name, unique_path};

    #[test]
    fn strips_path_traversal() {
        assert_eq!(
            safe_file_name("../../.ssh/authorized_keys"),
            "authorized_keys"
        );
        assert_eq!(safe_file_name("/etc/passwd"), "passwd");
        assert_eq!(
            safe_file_name(r"..\..\windows\system32\evil.dll"),
            "evil.dll"
        );
        assert_eq!(safe_file_name(".."), "attachment");
        assert_eq!(safe_file_name("."), "attachment");
        assert_eq!(safe_file_name(""), "attachment");
        assert_eq!(safe_file_name("   "), "attachment");
    }

    #[test]
    fn keeps_ordinary_names() {
        assert_eq!(safe_file_name("report.pdf"), "report.pdf");
        assert_eq!(
            safe_file_name("Relatório final.docx"),
            "Relatório final.docx"
        );
    }

    #[test]
    fn drops_control_characters() {
        assert_eq!(
            safe_file_name("evil\n\r\u{1b}[2Jname.txt"),
            "evil[2Jname.txt"
        );
    }

    #[test]
    fn avoids_overwriting_existing_files() {
        let dir = std::env::temp_dir().join("m365-files-test");
        std::fs::create_dir_all(&dir).unwrap();
        let taken = dir.join("a.txt");
        std::fs::write(&taken, b"x").unwrap();

        let next = unique_path(&dir, "a.txt");
        assert_eq!(next.file_name().unwrap(), "a (1).txt");
        assert_ne!(next, taken, "must not clobber an existing download");

        std::fs::remove_file(&taken).ok();
    }
}
