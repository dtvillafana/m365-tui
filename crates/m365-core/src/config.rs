//! Runtime configuration, loaded from the environment (and an optional `.env`).

use std::path::PathBuf;

use anyhow::{Context, Result};

/// Microsoft Graph base URL (v1.0 endpoint).
pub const GRAPH_BASE: &str = "https://graph.microsoft.com/v1.0";

/// Number of recently-active Teams chats accelerated by default.
pub const TEAMS_HOT_CHATS_DEFAULT: usize = 6;
/// Hard safety ceiling for accelerated Teams chats.
pub const TEAMS_HOT_CHATS_MAX: usize = 16;
/// Default request budget for optional/background Teams message GETs.
pub const TEAMS_POLL_BUDGET_DEFAULT_RPS: f64 = 6.0;
/// Runtime/configuration floor used after Graph throttling.
pub const TEAMS_POLL_BUDGET_MIN_RPS: f64 = 0.5;
/// Hard safety ceiling: 40% of the documented 20 rps per-app/per-tenant chat-message GET limit.
pub const TEAMS_POLL_BUDGET_MAX_RPS: f64 = 8.0;
/// Reject combinations that would make an accelerated chat slower than this.
pub const TEAMS_POLL_MAX_EFFECTIVE_INTERVAL_SECS: f64 = 5.0;

/// The Graph endpoint to talk to, overridable with `M365_GRAPH_BASE`.
///
/// Pointing this at a local mock lets the app run on fabricated data — useful
/// for recording a demo without a real mailbox on screen, and for exercising the
/// UI offline. Unset in normal use.
pub fn graph_base() -> String {
    std::env::var("M365_GRAPH_BASE")
        .ok()
        .map(|s| s.trim().trim_end_matches('/').to_string())
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| GRAPH_BASE.to_string())
}

/// The delegated scopes the app requests. `offline_access` is required for
/// refresh tokens; `openid`/`profile` give us the signed-in user's identity.
///
/// Keep this list stable: changing it invalidates existing consent, and in
/// tenants that disallow user consent every sign-in then needs a fresh admin
/// approval. New optional capabilities should be opt-in (see
/// [`PRESENCE_WRITE_SCOPE`]) rather than added here.
pub const DEFAULT_SCOPES: &[&str] = &[
    "openid",
    "profile",
    "offline_access",
    "User.Read",
    "People.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Chat.ReadWrite",
    "ChannelMessage.Send",
    "ChannelMessage.Read.All",
    "Presence.Read.All",
];

/// Needed only to *set* your own presence (`setUserPreferredPresence`).
/// Opt in with `M365_PRESENCE_WRITE=1`; requires the permission to be added to
/// the app registration and consented.
pub const PRESENCE_WRITE_SCOPE: &str = "Presence.ReadWrite";

/// Needed to list the teams and channels you belong to (`/me/joinedTeams`).
/// Chats work without it. Opt in with `M365_TEAMS_CHANNELS=1`.
pub const TEAMS_READ_SCOPE: &str = "Team.ReadBasic.All";

/// Directory search for starting chats outside the recent list.
pub const PEOPLE_SEARCH_SCOPE: &str = "User.ReadBasic.All";

/// Needed to download regular Teams file attachments from SharePoint/OneDrive.
/// Opt in with `M365_TEAMS_FILE_IMAGES=1`.
pub const FILES_READ_SCOPE: &str = "Files.Read.All";

/// Needed to read profile photos of other organization users.
/// Opt in with `M365_PROFILE_PHOTO=1`; kept out of DEFAULT_SCOPES so existing
/// consent is not disturbed for users who do not want contact avatars.
pub const PROFILE_PHOTO_SCOPE: &str = "User.ReadBasic.All";

/// Experimental full directory profile lookup for Teams contacts.
///
/// Opt in with `M365_DIRECTORY_PROFILE=1`. This adds delegated `User.Read.All`,
/// which requires administrator consent. The feature is currently untested.
pub const DIRECTORY_PROFILE_SCOPE: &str = "User.Read.All";

/// Needed to read and update Outlook automatic-reply (Out of Office) settings.
/// Kept opt-in because changing delegated scopes can trigger a new consent flow.
pub const MAILBOX_SETTINGS_WRITE_SCOPE: &str = "MailboxSettings.ReadWrite";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresenceActivitySource {
    Auto,
    Desktop,
    X11,
    Wayland,
    Logind,
    App,
}

impl PresenceActivitySource {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Desktop => "desktop",
            Self::X11 => "x11",
            Self::Wayland => "wayland",
            Self::Logind => "logind",
            Self::App => "app",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TeamsSystemEvents {
    Useful,
    All,
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NtfyMode {
    Never,
    Always,
    Away,
    AlwaysWd,
    AwayWd,
}

impl NtfyMode {
    pub fn enabled(self) -> bool {
        self != Self::Never
    }

    pub fn needs_presence(self) -> bool {
        matches!(self, Self::Away | Self::AwayWd)
    }

    pub fn needs_work_plan(self) -> bool {
        matches!(self, Self::AlwaysWd | Self::AwayWd)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalendarNotify {
    All,
    Internal,
    External,
    None,
}

impl CalendarNotify {
    pub fn enabled(self) -> bool {
        self != Self::None
    }

    pub fn internal(self) -> bool {
        matches!(self, Self::All | Self::Internal)
    }

    pub fn external(self) -> bool {
        matches!(self, Self::All | Self::External)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CacheDirSource {
    None,
    New,
    DeprecatedOnly,
    BothSame,
    BothDifferent {
        new_value: String,
        deprecated_value: String,
    },
}

impl CacheDirSource {
    pub fn warning(&self) -> Option<String> {
        match self {
            Self::None | Self::New => None,
            Self::DeprecatedOnly => Some(
                "Warning: M365_TEAMS_IMAGE_CACHE_DIR is deprecated.\nPlease use M365_CACHE_DIR instead."
                    .to_string(),
            ),
            Self::BothSame => Some(
                "Warning: M365_TEAMS_IMAGE_CACHE_DIR is deprecated and duplicates M365_CACHE_DIR.\nPlease remove or comment out M365_TEAMS_IMAGE_CACHE_DIR."
                    .to_string(),
            ),
            Self::BothDifferent {
                new_value,
                deprecated_value,
            } => Some(format!(
                "Warning: both M365_CACHE_DIR and deprecated M365_TEAMS_IMAGE_CACHE_DIR are set.\nUsing M365_CACHE_DIR={new_value:?}; ignoring M365_TEAMS_IMAGE_CACHE_DIR={deprecated_value:?}.\nPlease remove or comment out M365_TEAMS_IMAGE_CACHE_DIR."
            )),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CacheMaxSource {
    Default,
    New,
    DeprecatedOnly,
    BothSame,
    BothDifferent {
        new_value: String,
        deprecated_value: String,
    },
}

impl CacheMaxSource {
    pub fn warning(&self) -> Option<String> {
        match self {
            Self::Default | Self::New => None,
            Self::DeprecatedOnly => Some(
                "Warning: M365_TEAMS_IMAGE_CACHE_MAX_MB is deprecated.\nPlease use M365_CACHE_MAX_MB instead."
                    .to_string(),
            ),
            Self::BothSame => Some(
                "Warning: M365_TEAMS_IMAGE_CACHE_MAX_MB is deprecated and duplicates M365_CACHE_MAX_MB.\nPlease remove or comment out M365_TEAMS_IMAGE_CACHE_MAX_MB."
                    .to_string(),
            ),
            Self::BothDifferent {
                new_value,
                deprecated_value,
            } => Some(format!(
                "Warning: both M365_CACHE_MAX_MB and deprecated M365_TEAMS_IMAGE_CACHE_MAX_MB are set.\nUsing M365_CACHE_MAX_MB={new_value:?}; ignoring M365_TEAMS_IMAGE_CACHE_MAX_MB={deprecated_value:?}.\nPlease remove or comment out M365_TEAMS_IMAGE_CACHE_MAX_MB."
            )),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WebLinkOpenerSource {
    None,
    New,
    DeprecatedOnly,
    BothSame,
    BothDifferent {
        new_value: String,
        deprecated_value: String,
    },
}

impl WebLinkOpenerSource {
    pub fn warning(&self) -> Option<String> {
        match self {
            Self::None | Self::New => None,
            Self::DeprecatedOnly => Some(
                "Warning: M365_MEETING_OPENER is deprecated.\nPlease use M365_WEBLINK_OPENER instead."
                    .to_string(),
            ),
            Self::BothSame => Some(
                "Warning: M365_MEETING_OPENER is deprecated and duplicates M365_WEBLINK_OPENER.\nPlease remove or comment out M365_MEETING_OPENER."
                    .to_string(),
            ),
            Self::BothDifferent {
                new_value,
                deprecated_value,
            } => Some(format!(
                "Warning: both M365_WEBLINK_OPENER and deprecated M365_MEETING_OPENER are set.\nUsing M365_WEBLINK_OPENER={new_value:?}; ignoring M365_MEETING_OPENER={deprecated_value:?}.\nPlease remove or comment out M365_MEETING_OPENER."
            )),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Config {
    /// Entra application (client) ID of the registered public client.
    pub client_id: String,
    /// Directory tenant ID (or `organizations` / `common`).
    pub tenant_id: String,
    /// Delegated scopes to request.
    pub scopes: Vec<String>,
    /// Public HTTPS base of the tunnel that fronts the webhook service, e.g.
    /// `https://m365.example.com`. `None` disables push (poll-only mode).
    pub tunnel_base_url: Option<String>,
    /// Redis connection URL used to receive change events from the webhook.
    pub redis_url: String,
    /// Path of the on-disk token cache (0600).
    pub token_cache_path: PathBuf,
    /// Shared secret echoed in subscription `clientState` and verified by the
    /// webhook. Generated on first run if absent.
    pub client_state: String,
    /// Desktop notifications for direct messages and `@mentions`.
    pub notifications: bool,
    /// Forward notification events to ntfy.
    pub ntfy: NtfyMode,
    /// ntfy server root URL, e.g. https://ntfy.example.com.
    pub ntfy_server: Option<String>,
    /// ntfy topic to publish to.
    pub ntfy_topic: Option<String>,
    /// Optional ntfy Bearer token for protected topics.
    pub ntfy_token: Option<String>,
    /// Calendar reminder delivery mode: all, internal, external, or none.
    pub calendar_notify: CalendarNotify,
    /// Seconds an unread message must remain open before it is marked read.
    /// Zero marks it read immediately after the body is displayed.
    pub read_msg_timeout: u64,
    /// Show presence indicators for contacts in one-to-one Teams chats.
    pub presence_read: bool,
    /// Keep an Available application presence session active while the TUI runs.
    pub presence_primary: bool,
    /// Minutes of local inactivity before an automatic primary session becomes Away.
    /// Zero disables the automatic Away transition.
    pub presence_available_timeout_min: u64,
    /// Local activity source used by automatic primary presence.
    pub presence_activity_source: PresenceActivitySource,
    /// Restore the pre-lock application presence after confirmed unlock activity.
    /// Enabled by default; M365_PRESENCE_LOCK_RESTORE=0 disables it.
    pub presence_lock_restore: bool,
    /// Download regular Teams image file attachments from SharePoint/OneDrive.
    pub teams_file_images: bool,
    /// Experimental direct Entra directory profile lookup for Teams contacts.
    /// Off by default; requires delegated User.Read.All and administrator consent.
    pub directory_profile: bool,
    /// Read and update Outlook/Teams Out of Office automatic replies.
    /// Off by default; adds delegated MailboxSettings.ReadWrite with default scopes.
    pub out_of_office: bool,
    /// Optional common root for persistent application cache/state.
    /// M365_CACHE_DIR wins over the deprecated M365_TEAMS_IMAGE_CACHE_DIR.
    /// If neither is set, Teams/UI persistence remains disabled.
    pub cache_dir: Option<PathBuf>,
    /// Records which environment variable selected the common cache root.
    pub cache_dir_source: CacheDirSource,
    /// Maximum size in MiB of evictable persistent cache (Teams images + conversations).
    pub cache_max_mb: u64,
    /// Records which environment variable selected the common cache limit.
    pub cache_max_source: CacheMaxSource,
    /// Markdown-style formatting for outgoing Teams messages. Enabled by default.
    pub teams_markdown: bool,
    /// One-to-one peer IDs whose self-authored Graph messages should be
    /// treated as incoming automation events (for example Workflows "Post as User").
    pub teams_automation_peers: std::collections::HashSet<String>,
    /// Exact .env file loaded by dotenvy, when configuration came from one.
    pub dotenv_path: Option<PathBuf>,
    /// True when M365_TEAMS_AUTOMATION_PEERS already existed in the process
    /// environment before dotenv loading.
    pub teams_automation_peers_from_process_env: bool,
    /// Give Teams messages sent by applications/bots high notification priority.
    pub teams_application_high_priority: bool,
    /// Automatically prefill persistent Teams conversation caches in the background.
    /// Enabled by default; set M365_TEAMS_CACHE_WARMUP=0 to disable.
    pub teams_cache_warmup: bool,
    /// Maximum number of recently-active chats kept in accelerated polling.
    /// Default 6, hard maximum 16.
    pub teams_hot_chats: usize,
    /// Shared rate budget for optional/background Teams message GETs.
    /// This covers hot polling, unread counting and cold-cache warm-up.
    pub teams_poll_budget_rps: f64,
    /// Which Teams system events are visible in conversations.
    /// Defaults to Useful; accepts useful, all, or none.
    pub teams_system_events: TeamsSystemEvents,
    /// Optional executable used to open web links from the TUI.
    /// M365_WEBLINK_OPENER wins over deprecated M365_MEETING_OPENER.
    /// Unset uses the operating system handler (`xdg-open` / `open`).
    pub web_link_opener: Option<String>,
    /// Records which environment variable selected the web-link opener.
    pub web_link_opener_source: WebLinkOpenerSource,
}

impl Config {
    /// Load configuration from the process environment. Call
    /// [`Config::load_dotenv`] first if you want `.env` support.
    pub fn from_env() -> Result<Self> {
        let client_id = env_required("M365_CLIENT_ID")?;
        let tenant_id = std::env::var("M365_TENANT_ID").unwrap_or_else(|_| "organizations".into());
        let presence_primary = env_flag("M365_PRESENCE_PRIMARY");
        let teams_file_images = env_flag("M365_TEAMS_FILE_IMAGES");
        let directory_profile = env_flag("M365_DIRECTORY_PROFILE");
        let out_of_office = env_flag("M365_OUT_OF_OFFICE");
        let new_cache_dir = std::env::var("M365_CACHE_DIR").ok();
        let deprecated_cache_dir = std::env::var("M365_TEAMS_IMAGE_CACHE_DIR").ok();
        let (cache_dir, cache_dir_source) =
            resolve_cache_dir_values(new_cache_dir.as_deref(), deprecated_cache_dir.as_deref());
        let new_cache_max = std::env::var("M365_CACHE_MAX_MB").ok();
        let deprecated_cache_max = std::env::var("M365_TEAMS_IMAGE_CACHE_MAX_MB").ok();
        let (cache_max_mb, cache_max_source) =
            resolve_cache_max_values(new_cache_max.as_deref(), deprecated_cache_max.as_deref())?;
        let teams_markdown = env_flag_default_on("M365_TEAMS_MARKDOWN");
        let teams_automation_peers = parse_teams_automation_peers(
            std::env::var("M365_TEAMS_AUTOMATION_PEERS").ok().as_deref(),
        );
        let teams_application_high_priority = env_flag("M365_TEAMS_APPLICATION_HIGH_PRIORITY");
        let new_web_link_opener = std::env::var("M365_WEBLINK_OPENER").ok();
        let deprecated_meeting_opener = std::env::var("M365_MEETING_OPENER").ok();
        let (web_link_opener, web_link_opener_source) = resolve_web_link_opener_values(
            new_web_link_opener.as_deref(),
            deprecated_meeting_opener.as_deref(),
        );
        let teams_cache_warmup = env_flag_default_on("M365_TEAMS_CACHE_WARMUP");
        let teams_system_events =
            parse_teams_system_events(std::env::var("M365_TEAMS_SYSTEM_EVENTS").ok().as_deref())?;
        let teams_hot_chats = match std::env::var("M365_TEAMS_HOT_CHATS") {
            Ok(value) if !value.trim().is_empty() => value
                .trim()
                .parse::<usize>()
                .context("M365_TEAMS_HOT_CHATS must be an integer")?,
            _ => TEAMS_HOT_CHATS_DEFAULT,
        };
        let teams_poll_budget_rps = match std::env::var("M365_TEAMS_POLL_BUDGET_RPS") {
            Ok(value) if !value.trim().is_empty() => value
                .trim()
                .parse::<f64>()
                .context("M365_TEAMS_POLL_BUDGET_RPS must be a number")?,
            _ => TEAMS_POLL_BUDGET_DEFAULT_RPS,
        };
        validate_teams_polling(teams_hot_chats, teams_poll_budget_rps)?;
        let presence_available_timeout_min =
            match std::env::var("M365_PRESENCE_AVAILABLE_TIMEOUT_MIN") {
                Ok(value) if !value.trim().is_empty() => value.trim().parse::<u64>().context(
                    "M365_PRESENCE_AVAILABLE_TIMEOUT_MIN must be an integer number of minutes",
                )?,
                _ => 5,
            };
        let presence_activity_source = parse_presence_activity_source(
            std::env::var("M365_PRESENCE_ACTIVITY_SOURCE")
                .ok()
                .as_deref(),
        )?;
        let presence_lock_restore = parse_presence_lock_restore(
            std::env::var("M365_PRESENCE_LOCK_RESTORE").ok().as_deref(),
        )?;

        let scopes = match std::env::var("M365_SCOPES") {
            Ok(s) if !s.trim().is_empty() => s.split_whitespace().map(|s| s.to_string()).collect(),
            _ => {
                let mut s: Vec<String> = DEFAULT_SCOPES.iter().map(|s| s.to_string()).collect();
                // Presence *writing* is opt-in: adding a scope invalidates any
                // existing consent grant, which is disruptive in tenants that
                // require admin approval.
                if env_flag("M365_PRESENCE_WRITE") || presence_primary {
                    s.push(PRESENCE_WRITE_SCOPE.to_string());
                }
                if env_flag("M365_TEAMS_CHANNELS") {
                    s.push(TEAMS_READ_SCOPE.to_string());
                }
                if env_flag("M365_PEOPLE_SEARCH") {
                    s.push(PEOPLE_SEARCH_SCOPE.to_string());
                }
                if teams_file_images {
                    s.push(FILES_READ_SCOPE.to_string());
                }
                if env_flag("M365_PROFILE_PHOTO") {
                    s.push(PROFILE_PHOTO_SCOPE.to_string());
                }
                if directory_profile {
                    s.push(DIRECTORY_PROFILE_SCOPE.to_string());
                }
                if out_of_office {
                    s.push(MAILBOX_SETTINGS_WRITE_SCOPE.to_string());
                }
                s
            }
        };

        let tunnel_base_url = std::env::var("M365_TUNNEL_BASE_URL")
            .ok()
            .filter(|s| !s.trim().is_empty())
            .map(|s| s.trim_end_matches('/').to_string());

        let redis_url =
            std::env::var("M365_REDIS_URL").unwrap_or_else(|_| "redis://127.0.0.1:6379".into());

        let token_cache_path = match std::env::var("M365_TOKEN_CACHE") {
            Ok(p) if !p.trim().is_empty() => PathBuf::from(p),
            _ => default_cache_dir()?.join("token-cache.json"),
        };

        let client_state = std::env::var("M365_CLIENT_STATE")
            .ok()
            .filter(|s| !s.trim().is_empty())
            .unwrap_or_else(|| uuid::Uuid::new_v4().to_string());

        // On unless explicitly disabled.
        let notifications = !matches!(
            std::env::var("M365_NOTIFY").as_deref(),
            Ok("0") | Ok("false") | Ok("no") | Ok("off")
        );
        let ntfy = parse_ntfy_mode(std::env::var("M365_NTFY").ok().as_deref())?;
        let ntfy_server = std::env::var("M365_NTFY_SERVER")
            .ok()
            .map(|value| value.trim().trim_end_matches('/').to_string())
            .filter(|value| !value.is_empty());
        let ntfy_topic = std::env::var("M365_NTFY_TOPIC")
            .ok()
            .map(|value| value.trim().trim_matches('/').to_string())
            .filter(|value| !value.is_empty());
        let ntfy_token = std::env::var("M365_NTFY_TOKEN")
            .ok()
            .map(|value| value.trim().to_string())
            .filter(|value| !value.is_empty());

        if ntfy.enabled() {
            let server = ntfy_server
                .as_deref()
                .context("M365_NTFY_SERVER is required when M365_NTFY is enabled")?;
            anyhow::ensure!(
                server.starts_with("http://") || server.starts_with("https://"),
                "M365_NTFY_SERVER must start with http:// or https://"
            );
            anyhow::ensure!(
                ntfy_topic.is_some(),
                "M365_NTFY_TOPIC is required when M365_NTFY is enabled"
            );
        }

        let calendar_notify =
            parse_calendar_notify(std::env::var("M365_CALENDAR_NOTIFY").ok().as_deref())?;
        let read_msg_timeout = match std::env::var("M365_READ_MSG_TIMEOUT") {
            Ok(value) if !value.trim().is_empty() => value
                .trim()
                .parse::<u64>()
                .context("M365_READ_MSG_TIMEOUT must be an integer number of seconds")?,
            _ => 0,
        };
        let presence_read = env_flag("M365_PRESENCE_READ");

        Ok(Self {
            client_id,
            tenant_id,
            scopes,
            tunnel_base_url,
            redis_url,
            token_cache_path,
            client_state,
            notifications,
            ntfy,
            ntfy_server,
            ntfy_topic,
            ntfy_token,
            calendar_notify,
            read_msg_timeout,
            presence_read,
            presence_primary,
            presence_available_timeout_min,
            presence_activity_source,
            presence_lock_restore,
            teams_file_images,
            directory_profile,
            out_of_office,
            cache_dir,
            cache_dir_source,
            cache_max_mb,
            cache_max_source,
            teams_markdown,
            teams_automation_peers,
            dotenv_path: None,
            teams_automation_peers_from_process_env: false,
            teams_application_high_priority,
            teams_cache_warmup,
            teams_hot_chats,
            teams_poll_budget_rps,
            teams_system_events,
            web_link_opener,
            web_link_opener_source,
        })
    }

    /// Best-effort load of a `.env` file from the current directory or nearest
    /// parent. Missing file is not an error.
    pub fn load_dotenv() -> Option<PathBuf> {
        dotenvy::dotenv().ok()
    }

    pub fn scope_string(&self) -> String {
        self.scopes.join(" ")
    }

    /// Deferred cache-variable compatibility warning.
    ///
    /// Frontends should print this only after leaving any alternate-screen TUI.
    pub fn cache_warning(&self) -> Option<String> {
        let warnings: Vec<String> = [
            self.cache_dir_source.warning(),
            self.cache_max_source.warning(),
        ]
        .into_iter()
        .flatten()
        .collect();
        (!warnings.is_empty()).then(|| warnings.join("\n\n"))
    }

    /// Deferred compatibility warnings for deprecated environment variables.
    ///
    /// Frontends should print this only after leaving any alternate-screen TUI.
    pub fn compatibility_warning(&self) -> Option<String> {
        let warnings: Vec<String> = [self.cache_warning(), self.web_link_opener_source.warning()]
            .into_iter()
            .flatten()
            .collect();
        (!warnings.is_empty()).then(|| warnings.join("\n\n"))
    }

    /// Whether the token we request can set presence.
    pub fn can_write_presence(&self) -> bool {
        self.has_scope(PRESENCE_WRITE_SCOPE)
    }

    /// Whether the token we request can enumerate teams and channels.
    pub fn can_read_teams(&self) -> bool {
        self.has_scope(TEAMS_READ_SCOPE)
    }

    /// Whether the current token request includes read access to files.
    pub fn can_read_files(&self) -> bool {
        self.has_scope(FILES_READ_SCOPE)
    }

    /// Whether the requested token can read other users' basic profile/photo.
    pub fn can_read_profile_photos(&self) -> bool {
        self.has_scope(PROFILE_PHOTO_SCOPE) || self.has_scope(PEOPLE_SEARCH_SCOPE)
    }

    /// Whether the current token request includes mailbox-settings write access.
    pub fn can_write_out_of_office(&self) -> bool {
        self.has_scope(MAILBOX_SETTINGS_WRITE_SCOPE)
    }

    pub fn can_search_people(&self) -> bool {
        self.has_scope(PEOPLE_SEARCH_SCOPE)
    }

    fn has_scope(&self, scope: &str) -> bool {
        self.scopes.iter().any(|s| s.eq_ignore_ascii_case(scope))
    }

    pub fn devicecode_endpoint(&self) -> String {
        format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/devicecode",
            self.tenant_id
        )
    }

    pub fn token_endpoint(&self) -> String {
        format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        )
    }

    /// URL the TUI points its subscriptions at (`None` if no tunnel configured).
    pub fn notification_url(&self) -> Option<String> {
        self.tunnel_base_url
            .as_ref()
            .map(|b| format!("{b}/notifications"))
    }

    pub fn lifecycle_url(&self) -> Option<String> {
        self.tunnel_base_url
            .as_ref()
            .map(|b| format!("{b}/lifecycle"))
    }
}

fn parse_teams_automation_peers(raw: Option<&str>) -> std::collections::HashSet<String> {
    raw.into_iter()
        .flat_map(|value| value.split(|ch: char| ch == ',' || ch.is_ascii_whitespace()))
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
        .collect()
}

fn resolve_cache_dir_values(
    new_raw: Option<&str>,
    deprecated_raw: Option<&str>,
) -> (Option<PathBuf>, CacheDirSource) {
    let new_raw = new_raw.filter(|value| !value.trim().is_empty());
    let deprecated_raw = deprecated_raw.filter(|value| !value.trim().is_empty());

    match (new_raw, deprecated_raw) {
        (None, None) => (None, CacheDirSource::None),
        (Some(new_value), None) => (Some(PathBuf::from(new_value.trim())), CacheDirSource::New),
        (None, Some(deprecated_value)) => (
            Some(PathBuf::from(deprecated_value.trim())),
            CacheDirSource::DeprecatedOnly,
        ),
        (Some(new_value), Some(deprecated_value)) if new_value == deprecated_value => (
            Some(PathBuf::from(new_value.trim())),
            CacheDirSource::BothSame,
        ),
        (Some(new_value), Some(deprecated_value)) => (
            Some(PathBuf::from(new_value.trim())),
            CacheDirSource::BothDifferent {
                new_value: new_value.to_string(),
                deprecated_value: deprecated_value.to_string(),
            },
        ),
    }
}

fn resolve_web_link_opener_values(
    new_raw: Option<&str>,
    deprecated_raw: Option<&str>,
) -> (Option<String>, WebLinkOpenerSource) {
    let new_value = new_raw.map(str::trim).filter(|value| !value.is_empty());
    let deprecated_value = deprecated_raw
        .map(str::trim)
        .filter(|value| !value.is_empty());

    match (new_value, deprecated_value) {
        (None, None) => (None, WebLinkOpenerSource::None),
        (Some(new_value), None) => (Some(new_value.to_string()), WebLinkOpenerSource::New),
        (None, Some(deprecated_value)) => (
            Some(deprecated_value.to_string()),
            WebLinkOpenerSource::DeprecatedOnly,
        ),
        (Some(new_value), Some(deprecated_value)) if new_value == deprecated_value => {
            (Some(new_value.to_string()), WebLinkOpenerSource::BothSame)
        }
        (Some(new_value), Some(deprecated_value)) => (
            Some(new_value.to_string()),
            WebLinkOpenerSource::BothDifferent {
                new_value: new_value.to_string(),
                deprecated_value: deprecated_value.to_string(),
            },
        ),
    }
}

fn resolve_cache_max_values(
    new_raw: Option<&str>,
    deprecated_raw: Option<&str>,
) -> Result<(u64, CacheMaxSource)> {
    const DEFAULT_MB: u64 = 256;

    let new_raw = new_raw.filter(|value| !value.trim().is_empty());
    let deprecated_raw = deprecated_raw.filter(|value| !value.trim().is_empty());

    let parse = |name: &str, raw: &str| -> Result<u64> {
        let value = raw
            .trim()
            .parse::<u64>()
            .with_context(|| format!("{name} must be a positive integer number of MiB"))?;
        anyhow::ensure!(value > 0, "{name} must be greater than zero");
        Ok(value)
    };

    match (new_raw, deprecated_raw) {
        (None, None) => Ok((DEFAULT_MB, CacheMaxSource::Default)),
        (Some(new_value), None) => {
            Ok((parse("M365_CACHE_MAX_MB", new_value)?, CacheMaxSource::New))
        }
        (None, Some(deprecated_value)) => Ok((
            parse("M365_TEAMS_IMAGE_CACHE_MAX_MB", deprecated_value)?,
            CacheMaxSource::DeprecatedOnly,
        )),
        (Some(new_value), Some(deprecated_value)) if new_value == deprecated_value => Ok((
            parse("M365_CACHE_MAX_MB", new_value)?,
            CacheMaxSource::BothSame,
        )),
        (Some(new_value), Some(deprecated_value)) => Ok((
            parse("M365_CACHE_MAX_MB", new_value)?,
            CacheMaxSource::BothDifferent {
                new_value: new_value.to_string(),
                deprecated_value: deprecated_value.to_string(),
            },
        )),
    }
}

fn validate_teams_polling(hot_chats: usize, budget_rps: f64) -> Result<()> {
    anyhow::ensure!(
        (1..=TEAMS_HOT_CHATS_MAX).contains(&hot_chats),
        "M365_TEAMS_HOT_CHATS must be between 1 and {TEAMS_HOT_CHATS_MAX} (got {hot_chats})"
    );
    anyhow::ensure!(
        budget_rps.is_finite()
            && (TEAMS_POLL_BUDGET_MIN_RPS..=TEAMS_POLL_BUDGET_MAX_RPS).contains(&budget_rps),
        "M365_TEAMS_POLL_BUDGET_RPS must be between {TEAMS_POLL_BUDGET_MIN_RPS} and {TEAMS_POLL_BUDGET_MAX_RPS} (got {budget_rps})"
    );

    let effective = hot_chats as f64 / budget_rps;
    anyhow::ensure!(
        effective <= TEAMS_POLL_MAX_EFFECTIVE_INTERVAL_SECS,
        "M365_TEAMS_HOT_CHATS={hot_chats} with M365_TEAMS_POLL_BUDGET_RPS={budget_rps} implies about {effective:.2}s per accelerated chat at full load; maximum allowed is {TEAMS_POLL_MAX_EFFECTIVE_INTERVAL_SECS:.0}s"
    );
    Ok(())
}

fn parse_ntfy_mode(value: Option<&str>) -> Result<NtfyMode> {
    match value.map(str::trim).filter(|value| !value.is_empty()) {
        None => Ok(NtfyMode::Never),
        Some(value) if value.eq_ignore_ascii_case("never") => Ok(NtfyMode::Never),
        Some(value) if value.eq_ignore_ascii_case("always") => Ok(NtfyMode::Always),
        Some(value) if value.eq_ignore_ascii_case("away") => Ok(NtfyMode::Away),
        Some(value) if value.eq_ignore_ascii_case("alwayswd") => Ok(NtfyMode::AlwaysWd),
        Some(value) if value.eq_ignore_ascii_case("awaywd") => Ok(NtfyMode::AwayWd),
        Some(value) => anyhow::bail!(
            "M365_NTFY must be one of: never, always, away, alwayswd, awaywd (got {value:?})"
        ),
    }
}

fn parse_calendar_notify(value: Option<&str>) -> Result<CalendarNotify> {
    match value.map(str::trim).filter(|value| !value.is_empty()) {
        None => Ok(CalendarNotify::All),
        Some(value) if value.eq_ignore_ascii_case("all") => Ok(CalendarNotify::All),
        Some(value) if value.eq_ignore_ascii_case("internal") => Ok(CalendarNotify::Internal),
        Some(value) if value.eq_ignore_ascii_case("external") => Ok(CalendarNotify::External),
        Some(value) if value.eq_ignore_ascii_case("none") => Ok(CalendarNotify::None),
        Some(value) => anyhow::bail!(
            "M365_CALENDAR_NOTIFY must be one of: all, internal, external, none (got {value:?})"
        ),
    }
}

fn parse_presence_activity_source(value: Option<&str>) -> Result<PresenceActivitySource> {
    match value.map(str::trim).filter(|value| !value.is_empty()) {
        None => Ok(PresenceActivitySource::Auto),
        Some(value) if value.eq_ignore_ascii_case("auto") => Ok(PresenceActivitySource::Auto),
        Some(value) if value.eq_ignore_ascii_case("desktop") => {
            Ok(PresenceActivitySource::Desktop)
        }
        Some(value) if value.eq_ignore_ascii_case("x11") => Ok(PresenceActivitySource::X11),
        Some(value) if value.eq_ignore_ascii_case("wayland") => {
            Ok(PresenceActivitySource::Wayland)
        }
        Some(value) if value.eq_ignore_ascii_case("logind") => {
            Ok(PresenceActivitySource::Logind)
        }
        Some(value) if value.eq_ignore_ascii_case("app") => Ok(PresenceActivitySource::App),
        Some(value) => anyhow::bail!(
            "M365_PRESENCE_ACTIVITY_SOURCE must be one of: auto, desktop, x11, wayland, logind, app (got {value:?})"
        ),
    }
}

fn parse_presence_lock_restore(value: Option<&str>) -> Result<bool> {
    match value.map(str::trim).filter(|value| !value.is_empty()) {
        None | Some("1") => Ok(true),
        Some("0") => Ok(false),
        Some(value) => anyhow::bail!("M365_PRESENCE_LOCK_RESTORE must be 0 or 1 (got {value:?})"),
    }
}

fn parse_teams_system_events(value: Option<&str>) -> Result<TeamsSystemEvents> {
    match value.map(str::trim).filter(|value| !value.is_empty()) {
        None => Ok(TeamsSystemEvents::Useful),
        Some(value) if value.eq_ignore_ascii_case("useful") => Ok(TeamsSystemEvents::Useful),
        Some(value) if value.eq_ignore_ascii_case("all") => Ok(TeamsSystemEvents::All),
        Some(value) if value.eq_ignore_ascii_case("none") => Ok(TeamsSystemEvents::None),
        Some(value) => anyhow::bail!(
            "M365_TEAMS_SYSTEM_EVENTS must be one of: useful, all, none (got {value:?})"
        ),
    }
}

/// Enabled by default. Only `0`, `false`, `no`, or `off` disable it
/// (case-insensitive).
fn env_flag_default_on(key: &str) -> bool {
    std::env::var(key)
        .map(|value| {
            !matches!(
                value.trim().to_ascii_lowercase().as_str(),
                "0" | "false" | "no" | "off"
            )
        })
        .unwrap_or(true)
}

/// True for `1`, `true`, `yes`, `on` (case-insensitive).
fn env_flag(key: &str) -> bool {
    std::env::var(key)
        .map(|v| {
            let v = v.trim().to_ascii_lowercase();
            v == "1" || v == "true" || v == "yes" || v == "on"
        })
        .unwrap_or(false)
}

fn env_required(key: &str) -> Result<String> {
    std::env::var(key)
        .with_context(|| format!("required environment variable {key} is not set"))
        .map(|s| s.trim().to_string())
        .and_then(|s| {
            if s.is_empty() {
                anyhow::bail!("environment variable {key} is empty");
            }
            Ok(s)
        })
}

fn default_cache_dir() -> Result<PathBuf> {
    let dirs = directories::ProjectDirs::from("dev", "rootHytx", "m365-tui")
        .context("could not determine a config directory for this platform")?;
    let dir = dirs.config_dir().to_path_buf();
    std::fs::create_dir_all(&dir)
        .with_context(|| format!("creating config dir {}", dir.display()))?;
    Ok(dir)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample(tunnel: Option<&str>) -> Config {
        Config {
            client_id: "cid".into(),
            tenant_id: "organizations".into(),
            scopes: vec!["User.Read".into(), "Mail.Send".into()],
            tunnel_base_url: tunnel.map(|s| s.to_string()),
            redis_url: "redis://127.0.0.1:6379".into(),
            token_cache_path: PathBuf::from("/tmp/x.json"),
            client_state: "secret".into(),
            notifications: true,
            ntfy: NtfyMode::Never,
            ntfy_server: None,
            ntfy_topic: None,
            ntfy_token: None,
            calendar_notify: CalendarNotify::All,
            read_msg_timeout: 0,
            presence_read: false,
            presence_primary: false,
            presence_available_timeout_min: 5,
            presence_activity_source: PresenceActivitySource::Auto,
            presence_lock_restore: true,
            teams_file_images: false,
            directory_profile: false,
            out_of_office: false,
            cache_dir: None,
            cache_dir_source: CacheDirSource::None,
            cache_max_mb: 256,
            cache_max_source: CacheMaxSource::Default,
            teams_markdown: true,
            teams_automation_peers: std::collections::HashSet::new(),
            dotenv_path: None,
            teams_automation_peers_from_process_env: false,
            teams_application_high_priority: false,
            teams_cache_warmup: true,
            teams_hot_chats: TEAMS_HOT_CHATS_DEFAULT,
            teams_poll_budget_rps: TEAMS_POLL_BUDGET_DEFAULT_RPS,
            teams_system_events: TeamsSystemEvents::Useful,
            web_link_opener: None,
            web_link_opener_source: WebLinkOpenerSource::None,
        }
    }

    #[test]
    fn parses_teams_automation_peers() {
        let peers = parse_teams_automation_peers(Some(" peer-a,peer-b  peer-a\npeer-c "));
        assert_eq!(peers.len(), 3);
        assert!(peers.contains("peer-a"));
        assert!(peers.contains("peer-b"));
        assert!(peers.contains("peer-c"));
        assert!(parse_teams_automation_peers(None).is_empty());
    }

    #[test]
    fn resolves_web_link_opener_priority_and_legacy_fallback() {
        assert_eq!(
            resolve_web_link_opener_values(None, None),
            (None, WebLinkOpenerSource::None)
        );
        assert_eq!(
            resolve_web_link_opener_values(Some(" /usr/bin/firefox "), None),
            (
                Some("/usr/bin/firefox".to_string()),
                WebLinkOpenerSource::New
            )
        );
        assert_eq!(
            resolve_web_link_opener_values(None, Some(" /usr/bin/firefox ")),
            (
                Some("/usr/bin/firefox".to_string()),
                WebLinkOpenerSource::DeprecatedOnly
            )
        );
        assert_eq!(
            resolve_web_link_opener_values(Some(" /usr/bin/firefox "), Some("/usr/bin/firefox")),
            (
                Some("/usr/bin/firefox".to_string()),
                WebLinkOpenerSource::BothSame
            )
        );
        assert_eq!(
            resolve_web_link_opener_values(Some("/usr/bin/firefox"), Some("/usr/bin/chromium")),
            (
                Some("/usr/bin/firefox".to_string()),
                WebLinkOpenerSource::BothDifferent {
                    new_value: "/usr/bin/firefox".to_string(),
                    deprecated_value: "/usr/bin/chromium".to_string(),
                }
            )
        );
        assert_eq!(
            resolve_web_link_opener_values(Some("  "), Some(" /usr/bin/firefox ")),
            (
                Some("/usr/bin/firefox".to_string()),
                WebLinkOpenerSource::DeprecatedOnly
            ),
            "an empty new variable must not suppress the deprecated fallback"
        );
    }

    #[test]
    fn web_link_opener_warnings_cover_deprecated_configurations() {
        assert!(WebLinkOpenerSource::None.warning().is_none());
        assert!(WebLinkOpenerSource::New.warning().is_none());

        let legacy = WebLinkOpenerSource::DeprecatedOnly.warning().unwrap();
        assert!(legacy.contains("M365_MEETING_OPENER is deprecated"));
        assert!(legacy.contains("M365_WEBLINK_OPENER"));

        let same = WebLinkOpenerSource::BothSame.warning().unwrap();
        assert!(same.contains("duplicates M365_WEBLINK_OPENER"));

        let different = WebLinkOpenerSource::BothDifferent {
            new_value: "/usr/bin/firefox".to_string(),
            deprecated_value: "/usr/bin/chromium".to_string(),
        }
        .warning()
        .unwrap();
        assert!(different.contains("Using M365_WEBLINK_OPENER=\"/usr/bin/firefox\""));
        assert!(different.contains("ignoring M365_MEETING_OPENER=\"/usr/bin/chromium\""));
    }

    #[test]
    fn resolves_cache_dir_priority_and_legacy_fallback() {
        assert_eq!(
            resolve_cache_dir_values(None, None),
            (None, CacheDirSource::None)
        );
        assert_eq!(
            resolve_cache_dir_values(Some(" /new/cache "), None),
            (Some(PathBuf::from("/new/cache")), CacheDirSource::New)
        );
        assert_eq!(
            resolve_cache_dir_values(None, Some(" /legacy/cache ")),
            (
                Some(PathBuf::from("/legacy/cache")),
                CacheDirSource::DeprecatedOnly
            )
        );
        assert_eq!(
            resolve_cache_dir_values(Some("/same"), Some("/same")),
            (Some(PathBuf::from("/same")), CacheDirSource::BothSame)
        );
        assert_eq!(
            resolve_cache_dir_values(Some("/new"), Some("/old")),
            (
                Some(PathBuf::from("/new")),
                CacheDirSource::BothDifferent {
                    new_value: "/new".to_string(),
                    deprecated_value: "/old".to_string(),
                }
            )
        );
        assert_eq!(
            resolve_cache_dir_values(Some(" /same "), Some("/same")),
            (
                Some(PathBuf::from("/same")),
                CacheDirSource::BothDifferent {
                    new_value: " /same ".to_string(),
                    deprecated_value: "/same".to_string(),
                }
            ),
            "same/different classification intentionally compares exact environment strings"
        );
    }

    #[test]
    fn cache_dir_warnings_cover_deprecated_configurations() {
        assert!(CacheDirSource::None.warning().is_none());
        assert!(CacheDirSource::New.warning().is_none());

        let legacy = CacheDirSource::DeprecatedOnly.warning().unwrap();
        assert!(legacy.contains("deprecated"));
        assert!(legacy.contains("M365_CACHE_DIR"));

        let same = CacheDirSource::BothSame.warning().unwrap();
        assert!(same.contains("duplicates M365_CACHE_DIR"));

        let different = CacheDirSource::BothDifferent {
            new_value: "/new".to_string(),
            deprecated_value: "/old".to_string(),
        }
        .warning()
        .unwrap();
        assert!(different.contains("Using M365_CACHE_DIR=\"/new\""));
        assert!(different.contains("ignoring M365_TEAMS_IMAGE_CACHE_DIR=\"/old\""));
    }

    #[test]
    fn resolves_cache_max_priority_and_legacy_fallback() {
        assert_eq!(
            resolve_cache_max_values(None, None).unwrap(),
            (256, CacheMaxSource::Default)
        );
        assert_eq!(
            resolve_cache_max_values(Some(" 512 "), None).unwrap(),
            (512, CacheMaxSource::New)
        );
        assert_eq!(
            resolve_cache_max_values(None, Some(" 128 ")).unwrap(),
            (128, CacheMaxSource::DeprecatedOnly)
        );
        assert_eq!(
            resolve_cache_max_values(Some("64"), Some("64")).unwrap(),
            (64, CacheMaxSource::BothSame)
        );
        assert_eq!(
            resolve_cache_max_values(Some("32"), Some("64")).unwrap(),
            (
                32,
                CacheMaxSource::BothDifferent {
                    new_value: "32".into(),
                    deprecated_value: "64".into(),
                }
            )
        );
        assert!(resolve_cache_max_values(Some("0"), None).is_err());
        assert!(resolve_cache_max_values(None, Some("bogus")).is_err());
        assert_eq!(
            resolve_cache_max_values(Some("32"), Some("bogus"))
                .unwrap()
                .0,
            32,
            "deprecated value is ignored when the new value wins"
        );
    }

    #[test]
    fn cache_max_warnings_cover_deprecated_configurations() {
        assert!(CacheMaxSource::Default.warning().is_none());
        assert!(CacheMaxSource::New.warning().is_none());
        assert!(CacheMaxSource::DeprecatedOnly
            .warning()
            .unwrap()
            .contains("M365_CACHE_MAX_MB"));
        assert!(CacheMaxSource::BothSame
            .warning()
            .unwrap()
            .contains("duplicates M365_CACHE_MAX_MB"));
        assert!(CacheMaxSource::BothDifferent {
            new_value: "32".into(),
            deprecated_value: "64".into(),
        }
        .warning()
        .unwrap()
        .contains("ignoring M365_TEAMS_IMAGE_CACHE_MAX_MB=\"64\""));
    }

    #[test]
    fn parses_presence_activity_source() {
        assert_eq!(
            parse_presence_activity_source(None).unwrap(),
            PresenceActivitySource::Auto
        );
        assert_eq!(
            parse_presence_activity_source(Some(" X11 ")).unwrap(),
            PresenceActivitySource::X11
        );
        assert_eq!(
            parse_presence_activity_source(Some("wayland")).unwrap(),
            PresenceActivitySource::Wayland
        );
        assert_eq!(
            parse_presence_activity_source(Some("logind")).unwrap(),
            PresenceActivitySource::Logind
        );
        assert_eq!(
            parse_presence_activity_source(Some("app")).unwrap(),
            PresenceActivitySource::App
        );
        assert!(parse_presence_activity_source(Some("keyboard")).is_err());
    }

    #[test]
    fn parses_presence_lock_restore() {
        assert!(parse_presence_lock_restore(None).unwrap());
        assert!(parse_presence_lock_restore(Some("1")).unwrap());
        assert!(!parse_presence_lock_restore(Some("0")).unwrap());
        assert!(parse_presence_lock_restore(Some("true")).is_err());
    }

    #[test]
    fn parses_ntfy_mode() {
        assert_eq!(parse_ntfy_mode(None).unwrap(), NtfyMode::Never);
        assert_eq!(parse_ntfy_mode(Some(" ALWAYS ")).unwrap(), NtfyMode::Always);
        assert_eq!(parse_ntfy_mode(Some("away")).unwrap(), NtfyMode::Away);
        assert_eq!(
            parse_ntfy_mode(Some("alwayswd")).unwrap(),
            NtfyMode::AlwaysWd
        );
        assert_eq!(parse_ntfy_mode(Some("AWAYWD")).unwrap(), NtfyMode::AwayWd);
        assert_eq!(parse_ntfy_mode(Some("never")).unwrap(), NtfyMode::Never);
        assert!(parse_ntfy_mode(Some("sometimes")).is_err());

        assert!(!NtfyMode::Never.enabled());
        assert!(NtfyMode::Always.enabled());
        assert!(NtfyMode::Away.enabled());
        assert!(NtfyMode::AlwaysWd.enabled());
        assert!(NtfyMode::AwayWd.enabled());

        assert!(!NtfyMode::Always.needs_presence());
        assert!(NtfyMode::Away.needs_presence());
        assert!(!NtfyMode::AlwaysWd.needs_presence());
        assert!(NtfyMode::AwayWd.needs_presence());

        assert!(!NtfyMode::Always.needs_work_plan());
        assert!(!NtfyMode::Away.needs_work_plan());
        assert!(NtfyMode::AlwaysWd.needs_work_plan());
        assert!(NtfyMode::AwayWd.needs_work_plan());
    }

    #[test]
    fn parses_calendar_notification_mode() {
        assert_eq!(parse_calendar_notify(None).unwrap(), CalendarNotify::All);
        assert_eq!(
            parse_calendar_notify(Some(" internal ")).unwrap(),
            CalendarNotify::Internal
        );
        assert_eq!(
            parse_calendar_notify(Some("EXTERNAL")).unwrap(),
            CalendarNotify::External
        );
        assert_eq!(
            parse_calendar_notify(Some("none")).unwrap(),
            CalendarNotify::None
        );
        assert!(parse_calendar_notify(Some("desktop")).is_err());

        assert!(CalendarNotify::All.internal());
        assert!(CalendarNotify::All.external());
        assert!(CalendarNotify::Internal.internal());
        assert!(!CalendarNotify::Internal.external());
        assert!(!CalendarNotify::External.internal());
        assert!(CalendarNotify::External.external());
        assert!(!CalendarNotify::None.enabled());
    }

    #[test]
    fn parses_teams_system_event_visibility() {
        assert_eq!(
            parse_teams_system_events(None).unwrap(),
            TeamsSystemEvents::Useful
        );
        assert_eq!(
            parse_teams_system_events(Some(" useful ")).unwrap(),
            TeamsSystemEvents::Useful
        );
        assert_eq!(
            parse_teams_system_events(Some("ALL")).unwrap(),
            TeamsSystemEvents::All
        );
        assert_eq!(
            parse_teams_system_events(Some("none")).unwrap(),
            TeamsSystemEvents::None
        );
        assert!(parse_teams_system_events(Some("usefull")).is_err());
    }

    #[test]
    fn validates_teams_polling_limits_and_combinations() {
        assert!(validate_teams_polling(6, 6.0).is_ok());
        assert!(validate_teams_polling(16, 6.0).is_ok());
        assert!(validate_teams_polling(16, 3.2).is_ok());

        assert!(validate_teams_polling(0, 6.0).is_err());
        assert!(validate_teams_polling(17, 6.0).is_err());
        assert!(validate_teams_polling(6, 0.4).is_err());
        assert!(validate_teams_polling(6, 8.1).is_err());
        assert!(validate_teams_polling(16, 3.0).is_err());
        assert!(validate_teams_polling(6, f64::NAN).is_err());
    }

    #[test]
    fn builds_endpoints_and_urls() {
        let c = sample(Some("https://m365.example.com"));
        assert_eq!(c.scope_string(), "User.Read Mail.Send");
        assert_eq!(
            c.token_endpoint(),
            "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
        );
        assert_eq!(
            c.notification_url().as_deref(),
            Some("https://m365.example.com/notifications")
        );
        assert_eq!(
            c.lifecycle_url().as_deref(),
            Some("https://m365.example.com/lifecycle")
        );
    }

    #[test]
    fn graph_base_defaults_to_the_real_endpoint() {
        // Serialised with the override test below: both touch process env.
        std::env::remove_var("M365_GRAPH_BASE");
        assert_eq!(graph_base(), GRAPH_BASE);

        std::env::set_var("M365_GRAPH_BASE", "http://127.0.0.1:8765/v1.0/");
        assert_eq!(graph_base(), "http://127.0.0.1:8765/v1.0");

        // Empty means "not set", so an unfilled .env entry can't break the app.
        std::env::set_var("M365_GRAPH_BASE", "  ");
        assert_eq!(graph_base(), GRAPH_BASE);
        std::env::remove_var("M365_GRAPH_BASE");
    }

    #[test]
    fn no_tunnel_means_no_notification_url() {
        let c = sample(None);
        assert!(c.notification_url().is_none());
    }
}
