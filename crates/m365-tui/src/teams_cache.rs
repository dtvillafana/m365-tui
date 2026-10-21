//! Persistent bytes for inline images rendered by the local graphics backend.

use std::path::{Path, PathBuf};

fn fnv1a64(data: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325;
    for byte in data {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

pub fn image_disk_key(key: &str) -> String {
    format!("{:016x}{:016x}", fnv1a64(key.as_bytes()), fnv1a64(b"image"))
}

pub fn image_path(root: &Path, key: &str) -> PathBuf {
    root.join("images")
        .join(format!("{}.bin", image_disk_key(key)))
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

pub fn write_private_file(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    if let Some(parent) = path.parent() {
        create_private_dir(parent)?;
    }
    let tmp = path.with_extension(format!("tmp-{}", std::process::id()));
    std::fs::write(&tmp, bytes)?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(&tmp, std::fs::Permissions::from_mode(0o600))?;
    }
    std::fs::rename(&tmp, path)?;
    Ok(())
}

pub fn load_image_bytes(root: &Path, key: &str) -> Option<Vec<u8>> {
    std::fs::read(image_path(root, key)).ok()
}

pub fn store_image_bytes(root: &Path, key: &str, bytes: &[u8]) -> std::io::Result<()> {
    write_private_file(&image_path(root, key), bytes)
}
