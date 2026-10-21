//! Desktop activity detection for automatic primary presence.
//!
//! X11 is fully implemented through the XScreenSaver extension. Wayland is
//! represented explicitly in the backend hierarchy, but until the
//! ext-idle-notify-v1 client can be runtime-tested it deliberately uses
//! systemd-logind's IdleHint/LockedHint as a conservative fallback.

use std::process::Command;
use std::time::{Duration, Instant};

use m365_core::config::PresenceActivitySource;

const LOGIND_PROBE_INTERVAL: Duration = Duration::from_secs(5);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ActivityState {
    Active,
    Idle,
    Locked,
    Unknown,
}

impl ActivityState {
    fn label(self) -> &'static str {
        match self {
            Self::Active => "active",
            Self::Idle => "idle",
            Self::Locked => "locked",
            Self::Unknown => "unknown",
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ActivitySnapshot {
    pub state: ActivityState,
    pub idle_for: Option<Duration>,
    pub locked: Option<bool>,
    backend: String,
    degraded: bool,
    detail: Option<String>,
}

impl ActivitySnapshot {
    fn unknown(backend: impl Into<String>, degraded: bool, detail: Option<String>) -> Self {
        Self {
            state: ActivityState::Unknown,
            idle_for: None,
            locked: None,
            backend: backend.into(),
            degraded,
            detail,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ActivityDiagnostics {
    pub requested: &'static str,
    pub backend: String,
    pub session_type: String,
    pub state: &'static str,
    pub idle_for: Option<Duration>,
    pub locked: Option<bool>,
    pub degraded: bool,
    pub detail: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct LogindProbe {
    idle: Option<bool>,
    locked: Option<bool>,
    error: Option<String>,
}

enum Backend {
    #[cfg(target_os = "linux")]
    X11(Box<X11Backend>),
    Logind {
        label: String,
        degraded: bool,
        detail: Option<String>,
    },
    AppOnly {
        detail: Option<String>,
    },
}

pub struct ActivityMonitor {
    requested: PresenceActivitySource,
    session_type: String,
    backend: Backend,
    last_snapshot: ActivitySnapshot,
    last_logind_probe: Option<Instant>,
    logind_cache: LogindProbe,
}

impl ActivityMonitor {
    pub fn new(requested: PresenceActivitySource) -> Self {
        let session_type = std::env::var("XDG_SESSION_TYPE")
            .unwrap_or_else(|_| "unknown".into())
            .trim()
            .to_ascii_lowercase();

        let backend = choose_backend(requested, &session_type);
        let backend_label = backend_label(&backend);
        let degraded = backend_degraded(&backend);
        let detail = backend_detail(&backend);

        Self {
            requested,
            session_type,
            backend,
            last_snapshot: ActivitySnapshot::unknown(backend_label, degraded, detail),
            last_logind_probe: None,
            logind_cache: LogindProbe::default(),
        }
    }

    pub(crate) fn sample(&mut self, idle_threshold: Duration) -> ActivitySnapshot {
        let x11_result = match &mut self.backend {
            #[cfg(target_os = "linux")]
            Backend::X11(backend) => Some(backend.idle_time()),
            _ => None,
        };

        let snapshot = if let Some(result) = x11_result {
            let logind = self.logind_probe();
            match result {
                Ok(idle_for) => {
                    let locked = logind.locked;
                    ActivitySnapshot {
                        state: classify(idle_for, idle_threshold, locked),
                        idle_for: Some(idle_for),
                        locked,
                        backend: "X11 XScreenSaver".into(),
                        degraded: false,
                        detail: logind
                            .error
                            .map(|error| format!("lock hint unavailable: {error}")),
                    }
                }
                Err(error) => snapshot_from_logind(
                    &logind,
                    "logind fallback (X11 query failed)",
                    true,
                    Some(error),
                ),
            }
        } else {
            match &self.backend {
                Backend::Logind {
                    label,
                    degraded,
                    detail,
                } => {
                    let label = label.clone();
                    let degraded = *degraded;
                    let detail = detail.clone();
                    let probe = self.logind_probe();
                    snapshot_from_logind(&probe, &label, degraded, detail)
                }
                Backend::AppOnly { detail } => {
                    ActivitySnapshot::unknown("app-local only", false, detail.clone())
                }
                #[cfg(target_os = "linux")]
                Backend::X11(_) => unreachable!(),
            }
        };

        self.last_snapshot = snapshot.clone();
        snapshot
    }

    pub fn diagnostics(&self) -> ActivityDiagnostics {
        ActivityDiagnostics {
            requested: self.requested.as_str(),
            backend: self.last_snapshot.backend.clone(),
            session_type: self.session_type.clone(),
            state: self.last_snapshot.state.label(),
            idle_for: self.last_snapshot.idle_for,
            locked: self.last_snapshot.locked,
            degraded: self.last_snapshot.degraded,
            detail: self.last_snapshot.detail.clone(),
        }
    }

    fn logind_probe(&mut self) -> LogindProbe {
        if self
            .last_logind_probe
            .is_some_and(|at| at.elapsed() < LOGIND_PROBE_INTERVAL)
        {
            return self.logind_cache.clone();
        }

        self.last_logind_probe = Some(Instant::now());
        self.logind_cache = query_logind();
        self.logind_cache.clone()
    }
}

fn choose_backend(requested: PresenceActivitySource, session_type: &str) -> Backend {
    match requested {
        PresenceActivitySource::App => Backend::AppOnly { detail: None },
        PresenceActivitySource::Logind => Backend::Logind {
            label: "systemd-logind".into(),
            degraded: false,
            detail: None,
        },
        PresenceActivitySource::Wayland => wayland_backend(),
        PresenceActivitySource::X11 => x11_backend_or_fallback("explicit X11 backend unavailable"),
        PresenceActivitySource::Desktop | PresenceActivitySource::Auto => match session_type {
            "x11" => x11_backend_or_fallback("X11 backend unavailable"),
            "wayland" => wayland_backend(),
            _ => Backend::Logind {
                label: "systemd-logind fallback".into(),
                degraded: true,
                detail: Some(format!(
                    "unknown graphical session type {session_type:?}; using logind"
                )),
            },
        },
    }
}

fn wayland_backend() -> Backend {
    Backend::Logind {
        label: "Wayland ext-idle-notify v2 prototype -> logind".into(),
        degraded: true,
        detail: Some(
            "ext-idle-notify runtime backend is scaffolded but intentionally not enabled until runtime-tested"
                .into(),
        ),
    }
}

#[cfg(target_os = "linux")]
fn x11_backend_or_fallback(reason: &str) -> Backend {
    match X11Backend::new() {
        Ok(backend) => Backend::X11(Box::new(backend)),
        Err(error) => Backend::Logind {
            label: "systemd-logind fallback".into(),
            degraded: true,
            detail: Some(format!("{reason}: {error}")),
        },
    }
}

#[cfg(not(target_os = "linux"))]
fn x11_backend_or_fallback(reason: &str) -> Backend {
    Backend::AppOnly {
        detail: Some(format!("{reason}: X11 backend is Linux-only")),
    }
}

fn backend_label(backend: &Backend) -> String {
    match backend {
        #[cfg(target_os = "linux")]
        Backend::X11(_) => "X11 XScreenSaver".into(),
        Backend::Logind { label, .. } => label.clone(),
        Backend::AppOnly { .. } => "app-local only".into(),
    }
}

fn backend_degraded(backend: &Backend) -> bool {
    match backend {
        #[cfg(target_os = "linux")]
        Backend::X11(_) => false,
        Backend::Logind { degraded, .. } => *degraded,
        Backend::AppOnly { .. } => false,
    }
}

fn backend_detail(backend: &Backend) -> Option<String> {
    match backend {
        #[cfg(target_os = "linux")]
        Backend::X11(_) => None,
        Backend::Logind { detail, .. } | Backend::AppOnly { detail } => detail.clone(),
    }
}

fn classify(idle_for: Duration, idle_threshold: Duration, locked: Option<bool>) -> ActivityState {
    if locked == Some(true) {
        ActivityState::Locked
    } else if idle_threshold.is_zero() || idle_for < idle_threshold {
        ActivityState::Active
    } else {
        ActivityState::Idle
    }
}

fn snapshot_from_logind(
    probe: &LogindProbe,
    backend: &str,
    degraded: bool,
    detail: Option<String>,
) -> ActivitySnapshot {
    let mut detail_parts = Vec::new();
    if let Some(detail) = detail {
        detail_parts.push(detail);
    }
    if let Some(error) = probe.error.as_ref() {
        detail_parts.push(error.clone());
    }

    let state = if probe.locked == Some(true) {
        ActivityState::Locked
    } else {
        match probe.idle {
            Some(false) => ActivityState::Active,
            Some(true) => ActivityState::Idle,
            None => ActivityState::Unknown,
        }
    };

    ActivitySnapshot {
        state,
        idle_for: None,
        locked: probe.locked,
        backend: backend.to_string(),
        degraded,
        detail: (!detail_parts.is_empty()).then(|| detail_parts.join("; ")),
    }
}

fn query_logind() -> LogindProbe {
    let session = std::env::var("XDG_SESSION_ID")
        .ok()
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "self".into());

    let output = match Command::new("loginctl")
        .args([
            "show-session",
            &session,
            "--property=IdleHint",
            "--property=LockedHint",
            "--no-pager",
        ])
        .output()
    {
        Ok(output) => output,
        Err(error) => {
            return LogindProbe {
                error: Some(format!("loginctl unavailable: {error}")),
                ..LogindProbe::default()
            }
        }
    };

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return LogindProbe {
            error: Some(if stderr.is_empty() {
                format!("loginctl exited with {}", output.status)
            } else {
                format!("loginctl: {stderr}")
            }),
            ..LogindProbe::default()
        };
    }

    parse_logind(&String::from_utf8_lossy(&output.stdout))
}

fn parse_logind(text: &str) -> LogindProbe {
    let mut probe = LogindProbe::default();

    for line in text.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        let parsed = match value.trim().to_ascii_lowercase().as_str() {
            "yes" | "true" | "1" => Some(true),
            "no" | "false" | "0" => Some(false),
            _ => None,
        };
        match key.trim() {
            "IdleHint" => probe.idle = parsed,
            "LockedHint" => probe.locked = parsed,
            _ => {}
        }
    }

    if probe.idle.is_none() && probe.locked.is_none() {
        probe.error = Some("loginctl returned no IdleHint/LockedHint".into());
    }
    probe
}

#[cfg(target_os = "linux")]
struct X11Backend {
    connection: x11rb::rust_connection::RustConnection,
    root: u32,
}

#[cfg(target_os = "linux")]
impl X11Backend {
    fn new() -> Result<Self, String> {
        use x11rb::connection::Connection;
        use x11rb::protocol::screensaver::ConnectionExt;

        let (connection, screen_num) = x11rb::rust_connection::RustConnection::connect(None)
            .map_err(|error| error.to_string())?;
        let root = connection
            .setup()
            .roots
            .get(screen_num)
            .ok_or_else(|| "X11 screen index is out of range".to_string())?
            .root;

        connection
            .screensaver_query_version(1, 0)
            .map_err(|error| error.to_string())?
            .reply()
            .map_err(|error| error.to_string())?;

        Ok(Self { connection, root })
    }

    fn idle_time(&self) -> Result<Duration, String> {
        use x11rb::protocol::screensaver::ConnectionExt;

        let info = self
            .connection
            .screensaver_query_info(self.root)
            .map_err(|error| error.to_string())?
            .reply()
            .map_err(|error| error.to_string())?;

        Ok(Duration::from_millis(u64::from(info.ms_since_user_input)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn idle_classification_honours_threshold_and_lock() {
        let threshold = Duration::from_secs(300);
        assert_eq!(
            classify(Duration::from_secs(12), threshold, Some(false)),
            ActivityState::Active
        );
        assert_eq!(
            classify(Duration::from_secs(300), threshold, Some(false)),
            ActivityState::Idle
        );
        assert_eq!(
            classify(Duration::from_secs(1), threshold, Some(true)),
            ActivityState::Locked
        );
        assert_eq!(
            classify(Duration::from_secs(999), Duration::ZERO, Some(false)),
            ActivityState::Active
        );
    }

    #[test]
    fn parses_logind_hints() {
        let probe = parse_logind("IdleHint=yes\nLockedHint=no\n");
        assert_eq!(probe.idle, Some(true));
        assert_eq!(probe.locked, Some(false));
        assert!(probe.error.is_none());
    }
}
