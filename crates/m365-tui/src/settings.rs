use std::path::PathBuf;

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Settings {
    pub preview_mail_on_hover: bool,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            preview_mail_on_hover: true,
        }
    }
}

impl Settings {
    pub fn load() -> Self {
        settings_path()
            .and_then(|path| {
                let text = std::fs::read_to_string(&path)
                    .with_context(|| format!("reading {}", path.display()))?;
                serde_json::from_str(&text).with_context(|| format!("parsing {}", path.display()))
            })
            .unwrap_or_default()
    }

    pub fn save(&self) -> Result<()> {
        let path = settings_path()?;
        let text = serde_json::to_string_pretty(self)?;
        std::fs::write(&path, text).with_context(|| format!("writing {}", path.display()))
    }
}

fn settings_path() -> Result<PathBuf> {
    let dirs = directories::ProjectDirs::from("dev", "rootHytx", "m365-tui")
        .context("could not determine the m365-tui config directory")?;
    let dir = dirs.config_dir();
    std::fs::create_dir_all(dir).with_context(|| format!("creating {}", dir.display()))?;
    Ok(dir.join("settings.json"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_fields_keep_defaults() {
        let settings: Settings = serde_json::from_str("{}").unwrap();
        assert!(settings.preview_mail_on_hover);
    }
}
