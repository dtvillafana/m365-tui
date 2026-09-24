//! Persistent Teams conversation and image cache.

use std::path::{Path, PathBuf};

use m365_core::models::ChatMessage;
use serde_json::json;

const CONVERSATION_DIR: &str = "conversations";
const CACHE_VERSION: u32 = 1;
const MAX_MESSAGES: usize = 2000;
const MAX_BYTES: usize = 16 * 1024 * 1024;

fn fnv1a64(data: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325;
    for byte in data {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

pub fn conversation_key(chat_id: &str) -> String {
    format!("{:016x}{:016x}", fnv1a64(chat_id.as_bytes()), fnv1a64(b"chat"))
}

pub fn conversation_path(root: &Path, chat_id: &str) -> PathBuf {
    root.join(CONVERSATION_DIR)
        .join(format!("{}.json", conversation_key(chat_id)))
}

pub fn image_disk_key(key: &str) -> String {
    format!("{:016x}{:016x}", fnv1a64(key.as_bytes()), fnv1a64(b"image"))
}

pub fn image_path(root: &Path, key: &str) -> PathBuf {
    root.join("images").join(format!("{}.bin", image_disk_key(key)))
}

pub fn create_private_dir(path: &Path) -> std::io::Result<()> {
    std::fs::create_dir_all(path)?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(path, std::fs::Permissions::from_mode(0o700))?;
    }
    Ok(())
}

fn write_private_file(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    if let Some(parent) = path.parent() {
        create_private_dir(parent)?;
    }
    let tmp = path.with_extension(format!(
        "tmp-{}",
        std::process::id()
    ));
    std::fs::write(&tmp, bytes)?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(&tmp, std::fs::Permissions::from_mode(0o600))?;
    }
    std::fs::rename(&tmp, path)?;
    Ok(())
}

pub fn load_conversation(root: &Path, chat_id: &str) -> Option<Vec<ChatMessage>> {
    let path = conversation_path(root, chat_id);
    let bytes = std::fs::read(&path).ok()?;
    if bytes.len() > MAX_BYTES {
        let _ = std::fs::remove_file(&path);
        return None;
    }
    let value: serde_json::Value = serde_json::from_slice(&bytes).ok()?;
    if value.get("version").and_then(|v| v.as_u64()) != Some(CACHE_VERSION as u64) {
        let _ = std::fs::remove_file(&path);
        return None;
    }
    if value.get("chat_id").and_then(|v| v.as_str()) != Some(chat_id) {
        let _ = std::fs::remove_file(&path);
        return None;
    }
    let messages: Vec<ChatMessage> = serde_json::from_value(value.get("messages")?.clone()).ok()?;
    Some(messages)
}

pub fn store_conversation(root: &Path, chat_id: &str, messages: &[ChatMessage]) -> std::io::Result<()> {
    let mut newest_first: Vec<&ChatMessage> = messages.iter().collect();
    newest_first.reverse();
    newest_first.truncate(MAX_MESSAGES);
    let payload = json!({
        "version": CACHE_VERSION,
        "chat_id": chat_id,
        "messages": newest_first,
    });
    let bytes = serde_json::to_vec(&payload).map_err(std::io::Error::other)?;
    if bytes.len() > MAX_BYTES {
        return Ok(());
    }
    write_private_file(&conversation_path(root, chat_id), &bytes)
}

pub fn load_image_bytes(root: &Path, key: &str) -> Option<Vec<u8>> {
    std::fs::read(image_path(root, key)).ok()
}

pub fn store_image_bytes(root: &Path, key: &str, bytes: &[u8]) -> std::io::Result<()> {
    write_private_file(&image_path(root, key), bytes)
}

pub fn prune_image_cache(root: &Path, max_bytes: u64) {
    let dir = root.join("images");
    let Ok(entries) = std::fs::read_dir(&dir) else {
        return;
    };
    let mut files: Vec<(std::time::SystemTime, u64, PathBuf)> = entries
        .flatten()
        .filter_map(|entry| {
            let meta = entry.metadata().ok()?;
            if !meta.is_file() {
                return None;
            }
            Some((
                meta.modified().unwrap_or(std::time::SystemTime::UNIX_EPOCH),
                meta.len(),
                entry.path(),
            ))
        })
        .collect();
    files.sort_by_key(|(mtime, _, _)| *mtime);
    let mut total: u64 = files.iter().map(|(_, size, _)| *size).sum();
    for (_, size, path) in files {
        if total <= max_bytes {
            break;
        }
        if std::fs::remove_file(&path).is_ok() {
            total = total.saturating_sub(size);
        }
    }
}
