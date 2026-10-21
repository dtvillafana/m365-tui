//! Application state and the async orchestration behind it.
//!
//! The UI thread never blocks on the network: key handlers spawn tokio tasks
//! that fetch from Graph and send an [`AppMessage`] back over an mpsc channel,
//! which the main loop applies to the state before the next redraw.

use std::future::Future;
use std::io::IsTerminal;
use std::time::Duration;

use anyhow::Context;

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers, MouseEvent, MouseEventKind};
use m365_core::config::{NtfyMode, TeamsSystemEvents, TEAMS_POLL_BUDGET_MIN_RPS};
use m365_core::events::{ChangeEvent, ChangeKind};
use m365_core::models::{
    Attachment, Chat, ChatMessage, Event as CalEvent, IdentitySet, MailFolder, MailMessage,
    MailboxSettingsSnapshot, Person, Presence, Recipient, SystemEventClass, Team, User,
};
use m365_core::{
    calendar, channels, chats, mail, mailbox, people, teams_presence, work_plan, OutgoingBody,
    Session,
};
#[cfg(test)]
use ratatui::text::Line;
use ratatui::text::Text;
use tokio::sync::mpsc;

use crate::content;
use crate::diagnostics::{
    ContactDiagnosticsRemote, ContactDiagnosticsState, DiagnosticsRemote, DiagnosticsState,
    PresenceProbe,
};
use crate::editor::TextInput;
use crate::navigation;
use crate::notify::NotificationPriority;
use crate::out_of_office;
use crate::presence_state::{
    PresenceStateDiagnostics, PresenceStateManager, PresenceWrite, SessionTarget, WriteReason,
    AVAILABLE, AWAY,
};

/// Messages sent from background tasks to the UI loop.
#[derive(Debug)]
#[allow(clippy::large_enum_variant)] // Mail payloads stay inline and are immediately applied.
pub enum AppMessage {
    Status(String),
    Error(String),
    Whoami(User),
    Folders(Vec<MailFolder>),
    Messages {
        items: Vec<MailMessage>,
        next: Option<String>,
        mode: ListUpdate,
    },
    MessageBody(MailMessage),
    AutoMarkReadDue {
        id: String,
        generation: u64,
    },
    Calendar(Vec<CalEvent>),
    CalendarReminders(Vec<CalEvent>),
    Chats(Vec<Chat>),
    ChatVisibilityChanged {
        chat_id: String,
        hidden: bool,
    },
    StayArchivedRehideFinished {
        chat_id: String,
        error: Option<String>,
    },
    ContactPresences(Vec<Presence>),
    ContactProfileLoaded {
        chat_id: String,
        person: Option<Person>,
        avatar: Option<image::DynamicImage>,
    },
    ChatUnreadCount {
        chat_id: String,
        count: usize,
    },
    ChatUnreadRefreshDone,
    ChatMessages {
        chat_id: String,
        messages: Vec<ChatMessage>,
        next: Option<String>,
        mode: ListUpdate,
    },
    ChatMessagesFailed {
        chat_id: String,
        mode: ListUpdate,
        error: String,
    },
    ChatMessageSent {
        chat_id: String,
        message_id: String,
        status: String,
    },
    ChatPinsLoaded {
        chat_id: String,
        messages: Vec<ChatMessage>,
        status: Option<String>,
    },
    ChatMarkedUnread {
        chat_id: String,
    },
    ChatTopicUpdated {
        chat_id: String,
        topic: String,
        status: String,
    },
    ChatCacheWarmed {
        chat_id: String,
        messages: Vec<ChatMessage>,
    },
    /// Result of one rate-limited accelerated/background chat poll.
    HotChatPollFinished {
        chat_id: String,
        messages: Vec<ChatMessage>,
        next: Option<String>,
        error: Option<String>,
    },
    HotChatPollSkipped {
        chat_id: String,
    },
    Teams(Vec<Team>),
    Channels {
        team_id: String,
        channels: Vec<m365_core::models::Channel>,
    },
    ChannelMessages {
        team_id: String,
        channel_id: String,
        messages: Vec<ChatMessage>,
        next: Option<String>,
        mode: ListUpdate,
    },
    /// Kitty image previews for the currently selected Teams message.
    TeamsImages {
        key: TeamsImageKey,
        images: Vec<image::DynamicImage>,
    },
    /// A send/action completed; optional status text and refresh hint.
    Done(String),
    /// Result of a cross-navigation request to open a chat by email.
    OpenChat(Option<String>),
    /// The signed-in user's presence (status), for the tab-bar indicator.
    /// `requested` is set when this followed an explicit change, so we can tell
    /// the user if the effective presence didn't match what they asked for.
    Presence {
        presence: Presence,
        requested: Option<String>,
    },
    /// Silent own-presence refresh used by ntfy `away` forwarding.
    PresenceRefresh(Presence),
    /// Completion of a Teams presence status-message update.
    PresenceStatusMessageUpdated {
        presence: Presence,
        cleared: bool,
    },
    /// Completion of one serialized write to this app's presence session.
    PresenceSessionWriteFinished {
        write: PresenceWrite,
        requested: Option<String>,
        presence: Option<Presence>,
        success: bool,
        error: Option<String>,
        retry_preference: Option<PresencePreferenceAction>,
    },
    /// Silent Microsoft 365 work-plan result used by `alwayswd` / `awaywd`.
    NtfyWorkPlan(bool),
    /// Fresh mailbox settings for /ooo or /ooo set.
    OutOfOfficeLoaded {
        settings: MailboxSettingsSnapshot,
        edit: bool,
    },
    /// Completion of an OOO write.
    OutOfOfficeSaved(String),
    /// Fresh read-only data for the F6 diagnostics overlay.
    DiagnosticsLoaded(DiagnosticsRemote),
    /// Fresh read-only data for the selected-contact F7 diagnostics overlay.
    ContactDiagnosticsLoaded(ContactDiagnosticsRemote),
    /// Newest inbox messages, fetched purely to drive notifications.
    InboxPeek(Vec<MailMessage>),
    /// Attachments of the open mail message.
    Attachments {
        message_id: String,
        items: Vec<Attachment>,
    },
    /// A downloaded attachment, ready to write to disk.
    Downloaded {
        name: String,
        bytes: Vec<u8>,
    },
    /// Push-notification health, for the status bar.
    /// Decoded image attachments for the currently open Outlook message.
    Push(PushState),
    /// Lightweight timer: refresh memory usage and expire stale status text.
    Tick,
    /// One-second local scheduler tick for accelerated Teams chats.
    HotPollTick,
    /// Periodic tick: refresh the current view from the server.
    Poll,

    MailThread {
        conversation_id: String,
        items: Vec<MailMessage>,
        truncated: bool,
    },
    PeopleSearch {
        query: String,
        result: Result<Vec<User>, String>,
    },
    MailSent(String),
    TeamsSent(String),
    TeamsSendFailed(String),
    HostedImage {
        key: String,
        bytes: Vec<u8>,
    },
    HostedImageFailed {
        key: String,
    },
    OpenMailDue {
        generation: u64,
    },
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Screen {
    Outlook,
    Teams,
    Calendar,
}

/// Which kind of reply the user asked for.
enum ReplyMode {
    Reply,
    ReplyAll,
    Forward,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum OutlookFocus {
    Folders,
    Messages,
    Reading,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TeamsMode {
    Chats,
    Channels,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TeamsFocus {
    List,
    Messages,
    Composer,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum TeamsSlashCommand {
    Topic(String),
    Rename(String),
    Edit(String),
    Pin,
    Unpin,
    Pins,
    Unread,
    Delete,
    OutOfOffice,
    OutOfOfficeSet,
    OutOfOfficeClear,
    Help,
    HelpAll,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChatPollTier {
    Hot,
    Warm,
    Cool,
    Normal,
}

impl ChatPollTier {
    fn interval(self) -> Option<std::time::Duration> {
        match self {
            Self::Hot => Some(std::time::Duration::from_secs(2)),
            Self::Warm => Some(std::time::Duration::from_secs(5)),
            Self::Cool => Some(std::time::Duration::from_secs(10)),
            Self::Normal => None,
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub struct TeamsPollDiagnostics {
    pub hot: usize,
    pub warm: usize,
    pub cool: usize,
    pub normal: usize,
    pub hot_pending: usize,
    pub foreground_in_flight: usize,
    pub runtime_budget_rps: f64,
    pub reserved_chats: usize,
}

fn teams_message_id_is_newer(new_id: &str, old_id: &str) -> bool {
    match (new_id.parse::<u64>(), old_id.parse::<u64>()) {
        (Ok(new_id), Ok(old_id)) => new_id > old_id,
        _ => new_id > old_id,
    }
}

fn chat_poll_tier_from_age(age: std::time::Duration) -> ChatPollTier {
    let seconds = age.as_secs();
    if seconds <= TEAMS_HOT_AGE_SECONDS {
        ChatPollTier::Hot
    } else if seconds <= TEAMS_WARM_AGE_SECONDS {
        ChatPollTier::Warm
    } else if seconds <= TEAMS_COOL_AGE_SECONDS {
        ChatPollTier::Cool
    } else {
        ChatPollTier::Normal
    }
}

/// A transient full-screen/modal overlay.
#[allow(clippy::large_enum_variant)] // Only one overlay exists; boxing complicates input handling.
pub enum Overlay {
    Notice(String),
    Help,
    TeamsCommands,
    Diagnostics,
    ContactDiagnostics,
    Palette {
        query: String,
        sel: usize,
    },
    Search {
        query: String,
    },
    Compose(Compose),
    Calendar,
    CalendarEvent,
    ContactProfile,
    /// Emoji reaction picker for the selected Teams message.
    React {
        row: usize,
        col: usize,
    },
    /// Pinned messages for the currently open chat.
    PinnedMessages {
        chat_id: String,
        messages: Vec<ChatMessage>,
        sel: usize,
    },
    /// Explicit confirmation before renaming a group chat.
    ConfirmRename {
        chat_id: String,
        name: String,
    },
    /// Teams chat pinning is single-slot: a new pin replaces the old one.
    ConfirmReplacePin {
        chat_id: String,
        message_id: String,
    },
    /// Explicit confirmation before soft-deleting a chat message.
    ConfirmDelete {
        chat_id: String,
        message_id: String,
        user_id: String,
    },
    /// Explicit confirmation before persisting an F7 automation peer.
    ConfirmAutomationPeer {
        peer_id: String,
        dotenv_path: String,
    },
    /// Presence (status) picker for the signed-in user.
    Presence,
    /// Single-line editor for the Teams presence status message.
    PresenceMessage(TextInput),
    /// Read-only Outlook/Teams Out of Office state.
    OutOfOffice(Box<out_of_office::View>),
    /// Scheduled Out of Office editor.
    OutOfOfficeEdit(Box<out_of_office::Editor>),
    /// Confirmation before disabling automatic replies.
    ConfirmOutOfOfficeClear,
    /// Temporary pause for ntfy forwarding.
    NtfySnooze {
        sel: usize,
    },
    /// Numbered links in the focused message, to open in a browser.
    Links,
    /// Attachments of the open mail message, to save to disk.
    Attachments,

    FolderSearch {
        query: String,
        sel: usize,
    },
    NewChat {
        query: String,
        results: Vec<User>,
        sel: usize,
        loading: bool,
    },
    ViewImage {
        keys: Vec<String>,
        sel: usize,
    },
    Settings {
        sel: usize,
    },
    MailImages {
        scroll: u16,
        fullscreen: Option<usize>,
    },
    MoveMail {
        sel: usize,
    },
}

/// Whether Graph push notifications are working.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PushState {
    /// No tunnel configured — polling only.
    Off,
    /// Tunnel configured, subscriptions being created.
    Connecting,
    /// Subscriptions live; changes arrive in seconds.
    Live,
    /// Graph rejected the subscriptions — still polling, but not instant.
    Failed(String),
}

/// Ticks (see `TICK_SECONDS` in main) a status message survives before being
/// cleared. Long enough to read, short enough not to linger.
const STATUS_TICKS_TO_LIVE: u32 = 5;

/// How many items to fetch per page — mail messages and Teams messages alike.
/// Scrolling to the end of a list pulls the next page of this size.
pub const PAGE_SIZE: u32 = 50;
const DEFAULT_CALENDAR_DAYS: i64 = 30;
const CALENDAR_RANGES: &[i64] = &[7, 14, 30, 60, 90, 180, 365];
const CALENDAR_MONTH_MIN_WIDTH: u16 = 68;
const CALENDAR_MONTH_MIN_HEIGHT: u16 = 23;

/// Wait this long after the selection settles before fetching a mail body or
/// thread, so moving through the list does not fire a Graph call per row.
const MAIL_FETCH_DELAY: Duration = Duration::from_secs(2);

/// How an incoming page of items updates the list already on screen.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListUpdate {
    /// Fresh context (folder switch, conversation opened, search): replace all.
    Replace,
    /// Older items fetched by scrolling: append after what's loaded.
    Append,
    /// Periodic refresh: fold the newest page in without dropping the older
    /// pages the user has already scrolled back through.
    Merge,
}

/// Reactions used by the Tele-style 4x7 picker.
///
/// Layout copied from tele's internal/ui/components/reactionpicker.go so the
/// same physical navigation and familiar reaction set are available here.
pub const REACTION_COLS: usize = 7;
pub const REACTIONS: &[&str] = &[
    "🤝",
    "🙏",
    "👍",
    "👌",
    "✍️",
    "😁",
    "😈",
    "💯",
    "🖕",
    "❤️",
    "🫡",
    "👏",
    "👨‍💻",
    "👀",
    "😎",
    "🏆",
    "🤔",
    "🤷",
    "😢",
    "🥰",
    "🦄",
    "🤡",
    "🔥",
    "💩",
    "😱",
    "👎",
    "😡",
    "🤬",
];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresencePreferenceAction {
    Set {
        availability: &'static str,
        activity: &'static str,
        label: &'static str,
    },
    Clear,
}

/// A status the user can pick.
pub struct PresenceOption {
    pub label: &'static str,
    /// What `setUserPreferredPresence` records — the sticky preference a Teams
    /// client will surface.
    pub preferred: (&'static str, &'static str),
    /// What this app publishes as its own presence *session*, which is what
    /// makes the status visible with no Teams client running. `None` means no
    /// session, i.e. appear offline. The session API accepts a narrower
    /// vocabulary than the preference API, hence the remapping.
    pub session: Option<(&'static str, &'static str)>,
}

/// Presence options, keyed 1-6.
pub const PRESENCE_OPTIONS: &[PresenceOption] = &[
    PresenceOption {
        label: "Available",
        preferred: ("Available", "Available"),
        session: Some(("Available", "Available")),
    },
    PresenceOption {
        label: "Busy",
        // The session API only offers Busy/InACall or Busy/InAConferenceCall.
        preferred: ("Busy", "Busy"),
        session: Some(("Busy", "InACall")),
    },
    PresenceOption {
        label: "Do not disturb",
        preferred: ("DoNotDisturb", "DoNotDisturb"),
        session: Some(("DoNotDisturb", "Presenting")),
    },
    PresenceOption {
        label: "Be right back",
        // No BeRightBack in the session vocabulary; Away is the closest.
        preferred: ("BeRightBack", "BeRightBack"),
        session: Some(("Away", "Away")),
    },
    PresenceOption {
        label: "Away",
        preferred: ("Away", "Away"),
        session: Some(("Away", "Away")),
    },
    PresenceOption {
        label: "Appear offline",
        preferred: ("Offline", "OffWork"),
        session: None,
    },
];

/// How long each presence session is asserted for.
pub const PRESENCE_SESSION_LEASE: &str = "PT1H";

/// Graph applies a separate five-minute activity timeout to Available sessions.
/// Refresh with a one-minute safety margin so Available cannot fade to
/// AvailableInactive/Away while m365-tui still considers the session active.
pub const PRESENCE_AVAILABLE_RENEW_AFTER: std::time::Duration =
    std::time::Duration::from_secs(4 * 60);

/// Non-Available sessions do not have Graph's five-minute Available activity
/// timeout, so keep the existing low-frequency lease renewal for them.
pub const PRESENCE_RENEW_AFTER: std::time::Duration = std::time::Duration::from_secs(30 * 60);

fn presence_renew_after(target: SessionTarget) -> std::time::Duration {
    if target == SessionTarget::Present(AVAILABLE) {
        PRESENCE_AVAILABLE_RENEW_AFTER
    } else {
        PRESENCE_RENEW_AFTER
    }
}

pub(crate) const COMPOSE_TO: usize = 0;
pub(crate) const COMPOSE_CC: usize = 1;
pub(crate) const COMPOSE_BCC: usize = 2;
pub(crate) const COMPOSE_SUBJECT: usize = 3;
pub(crate) const COMPOSE_BODY: usize = 4;
pub(crate) const COMPOSE_ATTACH: usize = 5;

#[allow(clippy::enum_variant_names)] // the "Mail" suffix distinguishes from Teams compose
pub enum ComposeKind {
    NewMail,
    ReplyMail { id: String },
    ReplyAllMail { id: String },
    ForwardMail { id: String },
}

impl ComposeKind {
    /// Which compose fields are shown/editable:
    /// To, Cc, Bcc, Subject, Body, Attach.
    pub fn fields(&self) -> &'static [usize] {
        match self {
            ComposeKind::NewMail => &[
                COMPOSE_TO,
                COMPOSE_CC,
                COMPOSE_BCC,
                COMPOSE_SUBJECT,
                COMPOSE_BODY,
                COMPOSE_ATTACH,
            ],
            ComposeKind::ReplyMail { .. } | ComposeKind::ReplyAllMail { .. } => {
                &[COMPOSE_BODY, COMPOSE_ATTACH]
            }
            ComposeKind::ForwardMail { .. } => &[
                COMPOSE_TO,
                COMPOSE_CC,
                COMPOSE_BCC,
                COMPOSE_BODY,
                COMPOSE_ATTACH,
            ],
        }
    }

    pub fn title(&self) -> &'static str {
        match self {
            ComposeKind::NewMail => "Compose mail",
            ComposeKind::ReplyMail { .. } => "Reply",
            ComposeKind::ReplyAllMail { .. } => "Reply all",
            ComposeKind::ForwardMail { .. } => "Forward",
        }
    }
}

pub struct Compose {
    pub kind: ComposeKind,
    pub to: TextInput,
    pub cc: TextInput,
    pub bcc: TextInput,
    pub subject: TextInput,
    pub body: TextInput,
    /// Path being typed in the Attach field.
    pub attach: TextInput,
    /// Files staged for sending, as (path, size).
    pub attachments: Vec<(std::path::PathBuf, u64)>,
    /// One of the `COMPOSE_*` field constants.
    pub field: usize,
    pub suggestion_sel: usize,
    pub edit_recipients: bool,
    /// Waiting for the second key of a Ctrl+X chord.
    pub ctrl_x: bool,
}

impl Compose {
    pub fn new(kind: ComposeKind, field: usize) -> Self {
        Self {
            kind,
            to: TextInput::new(),
            cc: TextInput::new(),
            bcc: TextInput::new(),
            subject: TextInput::new(),
            body: TextInput::new(),
            attach: TextInput::new(),
            attachments: Vec::new(),
            field,
            suggestion_sel: 0,
            edit_recipients: false,
            ctrl_x: false,
        }
    }

    /// Stage the path currently typed in the Attach field.
    /// Returns a message for the status line.
    fn stage_attachment(&mut self) -> String {
        let raw = self.attach.text();
        let raw = raw.trim();
        if raw.is_empty() {
            return String::new();
        }
        let path = expand_tilde(&crate::images::unescape_path(raw));
        match std::fs::metadata(&path) {
            Ok(md) if md.is_file() => {
                let size = md.len();
                self.attachments.push((path.clone(), size));
                self.attach.clear();
                format!("attached {} ({})", path.display(), human_size(size))
            }
            Ok(_) => format!("{} is not a file", path.display()),
            Err(e) => format!("cannot attach {}: {e}", path.display()),
        }
    }

    fn active_mut(&mut self) -> &mut TextInput {
        match self.field {
            COMPOSE_TO => &mut self.to,
            COMPOSE_CC => &mut self.cc,
            COMPOSE_BCC => &mut self.bcc,
            COMPOSE_SUBJECT => &mut self.subject,
            COMPOSE_ATTACH => &mut self.attach,
            _ => &mut self.body,
        }
    }

    fn active_recipient(&self) -> Option<&TextInput> {
        match self.field {
            COMPOSE_TO => Some(&self.to),
            COMPOSE_CC => Some(&self.cc),
            COMPOSE_BCC => Some(&self.bcc),
            _ => None,
        }
    }

    pub(crate) fn fields(&self) -> &'static [usize] {
        match &self.kind {
            ComposeKind::ReplyMail { .. } | ComposeKind::ReplyAllMail { .. }
                if self.edit_recipients =>
            {
                &[
                    COMPOSE_TO,
                    COMPOSE_CC,
                    COMPOSE_BCC,
                    COMPOSE_BODY,
                    COMPOSE_ATTACH,
                ]
            }
            _ => self.kind.fields(),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum CalendarView {
    Agenda,
    Month,
}

pub struct CalendarState {
    pub events: Vec<CalEvent>,
    pub selected: usize,
    pub days: i64,
    pub view: CalendarView,
    pub month_offset: i32,
}

impl Default for CalendarState {
    fn default() -> Self {
        Self {
            events: Vec::new(),
            selected: 0,
            days: DEFAULT_CALENDAR_DAYS,
            view: CalendarView::Agenda,
            month_offset: 0,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct MailContact {
    pub name: Option<String>,
    pub address: String,
}

pub struct MailImage {
    pub state: std::cell::RefCell<ratatui_image::protocol::StatefulProtocol>,
}

pub struct ContactProfileState {
    pub chat_id: String,
    pub user_id: Option<String>,
    pub display_name: String,
    pub fallback_email: Option<String>,
    pub account: String,
    pub person: Option<Person>,
    pub avatar: Option<MailImage>,
    pub loading: bool,
}

#[derive(Default)]
pub struct OutlookState {
    pub folders: Vec<MailFolder>,
    pub folder_sel: usize,
    /// Width of the folders panel. `None` until the first folder list arrives,
    /// at which point it is fitted to the rendered folder labels.
    pub folder_width: Option<u16>,
    pub folder_height: Option<u16>,
    pub messages: Vec<MailMessage>,
    /// Indices into `messages` for the rows currently visible in the list.
    /// Thread mode contains only the newest loaded message per conversation.
    pub message_rows: Vec<usize>,
    pub threaded: bool,
    pub msg_sel: usize,
    /// `@odata.nextLink` for the current folder listing (Some = more to load).
    pub messages_next: Option<String>,
    /// Guards against firing multiple "load more" requests at once.
    pub loading_more: bool,
    pub reading: Option<MailMessage>,
    /// Cached styled body of the open message (HTML parsed once, not per frame).
    pub reading_body: Option<Text<'static>>,
    /// Full mailbox-wide thread and its cached rendered bodies when thread mode
    /// is active. Empty for an individually-opened message.
    pub reading_thread: Vec<MailMessage>,
    pub reading_thread_bodies: Vec<Text<'static>>,
    /// Links referenced by the open message, numbered `[1]`, `[2]`, ...
    pub reading_links: Vec<String>,
    /// Attachments of the open message (fetched when it has any).
    pub reading_attachments: Vec<Attachment>,
    pub reading_images: Vec<MailBodyImage>,
    message_cache: std::collections::HashMap<String, MailMessage>,
    thread_cache: std::collections::HashMap<String, (Vec<MailMessage>, bool)>,
    attachment_cache: std::collections::HashMap<String, Vec<Attachment>>,
    attachment_loading: std::collections::HashSet<String>,
    pending_reading: Option<ReadingKey>,
    /// Incremented to cancel an in-flight mail-fetch debounce timer.
    mail_fetch_generation: u64,
    /// Line scroll offset of the reading pane.
    pub reading_scroll: u16,
}

struct UiState {
    screen: Screen,
    calendar_days: i64,
    calendar_view: CalendarView,
    teams_chat_id: Option<String>,
    teams_archive_expanded: bool,
}

impl Default for UiState {
    fn default() -> Self {
        Self {
            screen: Screen::Outlook,
            calendar_days: DEFAULT_CALENDAR_DAYS,
            calendar_view: CalendarView::Agenda,
            teams_chat_id: None,
            teams_archive_expanded: false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TeamsImageSource {
    Chat(String),
    Channel { team_id: String, channel_id: String },
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TeamsImageKey {
    source: TeamsImageSource,
    message_id: String,
}

pub struct TeamsState {
    pub mode: TeamsMode,
    pub chats: Vec<Chat>,
    /// Selection in the virtual chat list:
    /// active chats, one Archive drawer row, then archived chats when expanded.
    pub chat_sel: usize,
    pub archive_expanded: bool,
    /// Chats that remain in the local Archive drawer even when Teams
    /// automatically unhides them after a new message.
    pub stay_archived: std::collections::HashSet<String>,
    /// Automatic hideForUser calls currently queued/running for Stay archived.
    stay_archived_rehide_in_flight: std::collections::HashSet<String>,
    /// Presence by directory user id for one-to-one chat contacts.
    pub contact_presences: std::collections::HashMap<String, Presence>,
    /// Last known display name by one-to-one chat id. Federated chat member
    /// objects can omit displayName, so keep a name learned from messages.
    pub contact_names: std::collections::HashMap<String, String>,
    /// Last known peer directory user id by one-to-one chat id.
    pub contact_user_ids: std::collections::HashMap<String, String>,
    /// One-to-one chats whose peer is known to belong to another tenant.
    pub external_chats: std::collections::HashSet<String>,
    /// Signed-in tenant id learned from config, chat membership or messages.
    own_tenant_id: Option<String>,
    /// Server-derived unread message count by one-to-one chat id.
    pub chat_unread_counts: std::collections::HashMap<String, usize>,
    /// Latest message id already shown to the user per one-to-one chat.
    /// Shields local read state from a temporarily stale Graph viewpoint.
    pub locally_read_through: std::collections::HashMap<String, String>,
    pub teams: Vec<Team>,
    pub team_sel: usize,
    pub channels: Vec<m365_core::models::Channel>,
    pub channel_sel: usize,
    pub messages: Vec<ChatMessage>,
    /// Cached styled body per message, index-aligned with `messages`.
    pub messages_rendered: Vec<Text<'static>>,
    /// Links per message, index-aligned with `messages`.
    pub messages_links: Vec<Vec<String>>,
    /// Inline images per message, index-aligned with `messages`.
    pub messages_images: Vec<Vec<content::BodyImage>>,
    /// Selected message index in the conversation pane (drives scroll + react).
    pub msg_sel: usize,
    /// Kitty previews for the currently selected Teams message.
    pub selected_images: Vec<MailImage>,
    /// Key currently represented by `selected_images`.
    selected_image_key: Option<TeamsImageKey>,
    /// Decoded Teams thumbnails cached for the lifetime of this process/session.
    /// Empty vectors are cached too, so messages without usable images are not
    /// repeatedly inspected or downloaded.
    image_cache: std::collections::HashMap<TeamsImageKey, Vec<image::DynamicImage>>,
    /// Image requests already running in the background.
    image_in_flight: std::collections::HashSet<TeamsImageKey>,
    /// `@odata.nextLink` for older messages in the open conversation.
    pub messages_next: Option<String>,
    /// Guards against firing multiple "load older" requests at once.
    pub loading_more: bool,
    /// Messages that arrived while the user was reading further back.
    pub unseen: usize,
    /// Index of the message being replied to, while composing a reply.
    pub replying_to: Option<usize>,
    /// Chat/message IDs being edited through the composer.
    pub editing_message: Option<(String, String)>,
    pub open_chat_id: Option<String>,
    pub preview_chat_id: Option<String>,
    last_chat_id: Option<String>,
    chat_open_pending_read: Option<String>,
    pub open_channel: Option<(String, String)>,
    pub composer: TextInput,
    /// Clipboard / data-URI images staged for the next send.
    pub images: Vec<crate::images::InlineImage>,
    /// Pixel previews for `images`, when the terminal can draw them.
    pub composer_previews: Vec<Option<crate::termimg::ReadyImage>>,
    /// True while a send is in flight, so Enter cannot fire twice.
    pub sending: bool,
    /// Composer snapshot to restore if the in-flight send fails.
    pending_restore: Option<(String, Vec<crate::images::InlineImage>)>,
    pub focus: TeamsFocus,
}

impl Default for TeamsState {
    fn default() -> Self {
        Self {
            mode: TeamsMode::Chats,
            chats: Vec::new(),
            chat_sel: 0,
            archive_expanded: false,
            stay_archived: std::collections::HashSet::new(),
            stay_archived_rehide_in_flight: std::collections::HashSet::new(),
            contact_presences: std::collections::HashMap::new(),
            contact_names: std::collections::HashMap::new(),
            contact_user_ids: std::collections::HashMap::new(),
            external_chats: std::collections::HashSet::new(),
            own_tenant_id: None,
            chat_unread_counts: std::collections::HashMap::new(),
            locally_read_through: std::collections::HashMap::new(),
            teams: Vec::new(),
            team_sel: 0,
            channels: Vec::new(),
            channel_sel: 0,
            messages: Vec::new(),
            messages_rendered: Vec::new(),
            messages_links: Vec::new(),
            messages_images: Vec::new(),
            msg_sel: 0,
            selected_images: Vec::new(),
            selected_image_key: None,
            image_cache: std::collections::HashMap::new(),
            image_in_flight: std::collections::HashSet::new(),
            messages_next: None,
            loading_more: false,
            unseen: 0,
            replying_to: None,
            editing_message: None,
            open_chat_id: None,
            preview_chat_id: None,
            last_chat_id: None,
            chat_open_pending_read: None,
            open_channel: None,
            composer: TextInput::new(),
            images: Vec::new(),
            composer_previews: Vec::new(),
            sending: false,
            pending_restore: None,
            focus: TeamsFocus::List,
        }
    }
}

pub fn teams_chat_is_hidden(chat: &Chat) -> bool {
    chat.viewpoint
        .as_ref()
        .and_then(|viewpoint| viewpoint.is_hidden)
        .unwrap_or(false)
}

fn teams_sender_is_incoming(from: Option<&IdentitySet>, me_id: Option<&str>) -> bool {
    let Some(from) = from else {
        return false;
    };

    if let Some(user) = from.user.as_ref() {
        return match (user.id.as_deref(), me_id) {
            (Some(sender_id), Some(me_id)) => sender_id != me_id,
            _ => true,
        };
    }

    from.application.is_some()
}

fn teams_message_is_incoming(message: &ChatMessage, me_id: Option<&str>) -> bool {
    if message.deleted_date_time.is_some() || message.is_system_event() {
        return false;
    }

    teams_sender_is_incoming(message.from.as_ref(), me_id)
}

fn teams_chat_is_archived(chat: &Chat, stay_archived: &std::collections::HashSet<String>) -> bool {
    teams_chat_is_hidden(chat) || stay_archived.contains(&chat.id)
}

fn teams_active_chat_count(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
) -> usize {
    chats
        .iter()
        .filter(|chat| !teams_chat_is_archived(chat, stay_archived))
        .count()
}

fn teams_archived_chat_count(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
) -> usize {
    chats
        .iter()
        .filter(|chat| teams_chat_is_archived(chat, stay_archived))
        .count()
}

fn teams_chat_row_count(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
    archive_expanded: bool,
) -> usize {
    teams_active_chat_count(chats, stay_archived)
        + 1
        + if archive_expanded {
            teams_archived_chat_count(chats, stay_archived)
        } else {
            0
        }
}

fn teams_chat_index_for_row(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
    archive_expanded: bool,
    row: usize,
) -> Option<usize> {
    let active_count = teams_active_chat_count(chats, stay_archived);

    if row < active_count {
        return chats
            .iter()
            .enumerate()
            .filter(|(_, chat)| !teams_chat_is_archived(chat, stay_archived))
            .nth(row)
            .map(|(index, _)| index);
    }

    if row == active_count || !archive_expanded {
        return None;
    }

    chats
        .iter()
        .enumerate()
        .filter(|(_, chat)| teams_chat_is_archived(chat, stay_archived))
        .nth(row.saturating_sub(active_count + 1))
        .map(|(index, _)| index)
}

fn teams_chat_row_for_id(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
    archive_expanded: bool,
    chat_id: &str,
) -> Option<usize> {
    let chat = chats.iter().find(|chat| chat.id == chat_id)?;
    if teams_chat_is_archived(chat, stay_archived) {
        if !archive_expanded {
            return None;
        }
        let hidden_position = chats
            .iter()
            .filter(|candidate| teams_chat_is_archived(candidate, stay_archived))
            .position(|candidate| candidate.id == chat_id)?;
        Some(teams_active_chat_count(chats, stay_archived) + 1 + hidden_position)
    } else {
        chats
            .iter()
            .filter(|candidate| !teams_chat_is_archived(candidate, stay_archived))
            .position(|candidate| candidate.id == chat_id)
    }
}

fn teams_stay_archived_rehide_candidates(
    chats: &[Chat],
    stay_archived: &std::collections::HashSet<String>,
    in_flight: &std::collections::HashSet<String>,
) -> Vec<String> {
    chats
        .iter()
        .filter(|chat| {
            stay_archived.contains(&chat.id)
                && !teams_chat_is_hidden(chat)
                && !in_flight.contains(&chat.id)
        })
        .map(|chat| chat.id.clone())
        .collect()
}

struct TeamsHotChatState {
    last_activity: std::time::Instant,
    last_poll_started: Option<std::time::Instant>,
    latest_message_id: Option<String>,
    latest_message_at: Option<chrono::DateTime<chrono::Utc>>,
    hot_pending: bool,
    foreground_in_flight: bool,
}

#[derive(Clone)]
struct TeamsBackgroundLimiter {
    configured_rps: f64,
    state: std::sync::Arc<std::sync::Mutex<TeamsBackgroundBudgetState>>,
    permits: std::sync::Arc<tokio::sync::Semaphore>,
}

struct TeamsBackgroundBudgetState {
    runtime_rps: f64,
    tokens: f64,
    last_refill: std::time::Instant,
    last_recovery: std::time::Instant,
    throttle_generation: u64,
    reserved_chats: std::collections::HashSet<String>,
    last_chat_request: std::collections::HashMap<String, std::time::Instant>,
}

struct TeamsChatReservation {
    state: std::sync::Arc<std::sync::Mutex<TeamsBackgroundBudgetState>>,
    chat_id: String,
}

impl Drop for TeamsChatReservation {
    fn drop(&mut self) {
        if let Ok(mut state) = self.state.lock() {
            state.reserved_chats.remove(&self.chat_id);
        }
    }
}

struct TeamsBackgroundPermit {
    _chat: TeamsChatReservation,
    _global: tokio::sync::OwnedSemaphorePermit,
}

impl TeamsBackgroundLimiter {
    fn new(configured_rps: f64, throttle_generation: u64) -> Self {
        let now = std::time::Instant::now();
        Self {
            configured_rps,
            state: std::sync::Arc::new(std::sync::Mutex::new(TeamsBackgroundBudgetState {
                runtime_rps: configured_rps,
                tokens: TEAMS_BACKGROUND_CONCURRENCY as f64,
                last_refill: now,
                last_recovery: now,
                throttle_generation,
                reserved_chats: std::collections::HashSet::new(),
                last_chat_request: std::collections::HashMap::new(),
            })),
            permits: std::sync::Arc::new(tokio::sync::Semaphore::new(TEAMS_BACKGROUND_CONCURRENCY)),
        }
    }

    fn diagnostics_snapshot(&self) -> (f64, usize) {
        let state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        (state.runtime_rps, state.reserved_chats.len())
    }

    async fn reserve_chat(&self, chat_id: &str) -> TeamsChatReservation {
        loop {
            let reserved = {
                let mut state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                if state.reserved_chats.contains(chat_id) {
                    false
                } else {
                    state.reserved_chats.insert(chat_id.to_string());
                    true
                }
            };

            if reserved {
                return TeamsChatReservation {
                    state: self.state.clone(),
                    chat_id: chat_id.to_string(),
                };
            }
            tokio::time::sleep(std::time::Duration::from_millis(25)).await;
        }
    }

    async fn wait_per_chat_interval(&self, chat_id: &str, min_interval: std::time::Duration) {
        loop {
            let now = std::time::Instant::now();
            let wait = {
                let state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                state.last_chat_request.get(chat_id).and_then(|last| {
                    let elapsed = now.duration_since(*last);
                    (elapsed < min_interval).then(|| min_interval - elapsed)
                })
            };
            match wait {
                Some(wait) => tokio::time::sleep(wait).await,
                None => return,
            }
        }
    }

    fn note_chat_request(&self, chat_id: &str) {
        let mut state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        state
            .last_chat_request
            .insert(chat_id.to_string(), std::time::Instant::now());
    }

    fn chat_requested_since(&self, chat_id: &str, since: std::time::Instant) -> bool {
        let state = self
            .state
            .lock()
            .expect("Teams background limiter mutex poisoned");
        state
            .last_chat_request
            .get(chat_id)
            .is_some_and(|last| *last > since)
    }

    async fn acquire_foreground_chat(&self, chat_id: &str) -> TeamsChatReservation {
        let reservation = self.reserve_chat(chat_id).await;
        self.wait_per_chat_interval(
            chat_id,
            std::time::Duration::from_secs(TEAMS_FOREGROUND_PER_CHAT_MIN_SECONDS),
        )
        .await;
        self.note_chat_request(chat_id);
        reservation
    }

    async fn acquire_inner(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
        supersede_after: Option<std::time::Instant>,
    ) -> Option<TeamsBackgroundPermit> {
        // Background work earns a global slot/token before reserving a chat.
        // Foreground open/F5 can therefore cut ahead while background work is
        // merely queued for budget.
        let global = self
            .permits
            .clone()
            .acquire_owned()
            .await
            .expect("Teams background limiter semaphore closed");

        loop {
            let now = std::time::Instant::now();
            let generation = graph.throttle_generation();

            let wait = {
                let mut state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");

                if generation != state.throttle_generation {
                    state.throttle_generation = generation;
                    state.runtime_rps = (state.runtime_rps * 0.5).max(TEAMS_POLL_BUDGET_MIN_RPS);
                    state.tokens = state.tokens.min(1.0);
                    state.last_recovery = now;
                    tracing::warn!(
                        "Graph throttling observed; Teams background budget reduced to {:.2} rps",
                        state.runtime_rps
                    );
                } else if state.runtime_rps < self.configured_rps {
                    let steps = now.duration_since(state.last_recovery).as_secs()
                        / TEAMS_BACKGROUND_RECOVERY_SECONDS;
                    if steps > 0 {
                        state.runtime_rps = (state.runtime_rps
                            + steps as f64 * TEAMS_BACKGROUND_RECOVERY_STEP_RPS)
                            .min(self.configured_rps);
                        state.last_recovery += std::time::Duration::from_secs(
                            steps * TEAMS_BACKGROUND_RECOVERY_SECONDS,
                        );
                        tracing::debug!(
                            "Teams background budget recovered to {:.2} rps",
                            state.runtime_rps
                        );
                    }
                }

                let elapsed = now.duration_since(state.last_refill).as_secs_f64();
                let runtime_rps = state.runtime_rps;
                state.tokens =
                    (state.tokens + elapsed * runtime_rps).min(TEAMS_BACKGROUND_CONCURRENCY as f64);
                state.last_refill = now;

                if state.tokens >= 1.0 {
                    state.tokens -= 1.0;
                    None
                } else {
                    let seconds = ((1.0 - state.tokens) / runtime_rps).max(0.01);
                    Some(std::time::Duration::from_secs_f64(seconds))
                }
            };

            match wait {
                None => break,
                Some(wait) => tokio::time::sleep(wait).await,
            }
        }

        if supersede_after.is_some_and(|since| self.chat_requested_since(chat_id, since)) {
            return None;
        }

        let min_interval = std::time::Duration::from_secs(TEAMS_BACKGROUND_PER_CHAT_MIN_SECONDS);
        let reservation = loop {
            self.wait_per_chat_interval(chat_id, min_interval).await;
            let reservation = self.reserve_chat(chat_id).await;
            let now = std::time::Instant::now();
            let remaining = {
                let state = self
                    .state
                    .lock()
                    .expect("Teams background limiter mutex poisoned");
                state.last_chat_request.get(chat_id).and_then(|last| {
                    let elapsed = now.duration_since(*last);
                    (elapsed < min_interval).then(|| min_interval - elapsed)
                })
            };
            if let Some(wait) = remaining {
                drop(reservation);
                tokio::time::sleep(wait).await;
                continue;
            }
            break reservation;
        };

        if supersede_after.is_some_and(|since| self.chat_requested_since(chat_id, since)) {
            drop(reservation);
            return None;
        }

        self.note_chat_request(chat_id);
        Some(TeamsBackgroundPermit {
            _chat: reservation,
            _global: global,
        })
    }

    async fn acquire(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
    ) -> TeamsBackgroundPermit {
        self.acquire_inner(graph, chat_id, None)
            .await
            .expect("non-supersedable Teams background request disappeared")
    }

    async fn acquire_hot(
        &self,
        graph: &m365_core::GraphClient,
        chat_id: &str,
        scheduled_at: std::time::Instant,
    ) -> Option<TeamsBackgroundPermit> {
        self.acquire_inner(graph, chat_id, Some(scheduled_at)).await
    }
}

pub struct App {
    pub session: Session,
    pub tx: mpsc::Sender<AppMessage>,
    pub screen: Screen,
    pub outlook: OutlookState,
    /// Present only when terminal probing detects the Kitty graphics protocol.
    image_picker: Option<ratatui_image::picker::Picker>,
    pub outlook_focus: OutlookFocus,
    pub teams: TeamsState,
    pub calendar: CalendarState,
    /// Small independent near-term feed used for 15/5-minute reminders.
    calendar_reminder_events: Vec<CalEvent>,
    /// Event/start/threshold keys already announced during this process.
    calendar_reminder_fired: std::collections::HashSet<String>,
    /// Last time the near-term reminder feed was queued for refresh.
    calendar_reminder_last_refresh: Option<std::time::Instant>,
    pub overlay: Option<Overlay>,
    notice_until: Option<std::time::Instant>,
    pub status: String,
    pub me: Option<User>,
    /// The user's current presence (shown in the tab bar).
    pub my_presence: Option<Presence>,
    /// Latest Microsoft 365 work-plan answer for `alwayswd` / `awaywd`.
    /// `None` means no successful refresh has completed yet.
    ntfy_working_now: Option<bool>,
    /// Temporary ntfy suppression deadline, persisted across restarts.
    ntfy_snooze_until: Option<std::time::SystemTime>,
    /// Last confirmed successful Graph folder refresh.
    poll_last_success_at: std::time::Instant,
    /// A scheduled poll has started and its folder-health response is still pending.
    poll_health_pending: bool,
    /// Current network retry interval after consecutive late poll-health responses.
    poll_retry_interval_secs: u64,
    /// Earliest instant at which a scheduled poll may hit Graph again.
    poll_backoff_until: Option<std::time::Instant>,
    /// Whether instant push is working (shown in the status bar).
    pub push: PushState,
    /// Message ids already notified about, so a poll can't repeat them.
    notified: std::collections::HashSet<String>,
    /// Newest message id seen per chat, to spot arrivals in chats that aren't
    /// open. `None` until the first chat list lands — the first sync must not
    /// fire a burst of notifications for history.
    chat_seen: Option<std::collections::HashMap<String, String>>,
    /// Activity/rate state for accelerated Teams chats.
    teams_hot_chat_state: std::collections::HashMap<String, TeamsHotChatState>,
    /// Shared rate/concurrency limiter for optional Teams message GETs.
    teams_background_limiter: TeamsBackgroundLimiter,
    /// Prevent overlapping unread-count sweeps when Graph is slow/throttled.
    chat_unread_refresh_running: bool,
    /// Runtime automation peer set; starts from config and can be extended by F7.
    teams_automation_peers: std::collections::HashSet<String>,
    /// IDs returned by sends performed by this m365-tui instance.
    own_sent_chat_message_ids: std::collections::HashSet<String>,
    /// Inbox message ids already seen, same baseline rule as `chat_seen`.
    mail_seen: Option<std::collections::HashSet<String>>,
    /// State and transitions for the application presence session published by
    /// m365-tui. Graph's effective user presence remains separate in `my_presence`.
    presence_state: PresenceStateManager,
    /// Desktop-session activity monitor (X11 now; Wayland prototype/fallback).
    presence_activity: crate::activity::ActivityMonitor,
    /// Preferred-presence operation that must run immediately before the next
    /// serialized app-session write. Newer user choices replace older queued ones.
    presence_preference_pending: Option<PresencePreferenceAction>,
    /// Resident memory in KiB, refreshed on each tick.
    pub rss_kb: Option<u64>,
    /// Copy of `status` as of the last tick, plus how many ticks it has been
    /// unchanged — transient messages are cleared so nothing gets stuck.
    last_status: String,
    status_ticks: u32,
    /// Width the focused text area was last laid out at, so Up/Down move by the
    /// rows actually on screen. Set by the renderer, read by the key handler.
    pub text_width_hint: std::cell::Cell<usize>,
    /// Largest useful reading-pane scroll offset, set by the renderer once it
    /// knows the wrapped height of the open message.
    pub reading_max_scroll: std::cell::Cell<u16>,
    /// Largest useful mail-image gallery scroll offset, set by the renderer.
    pub images_max_scroll: std::cell::Cell<u16>,
    /// Largest useful copy-mode scroll offset, set by the renderer.
    pub copy_max_scroll: std::cell::Cell<u16>,
    /// Current vertical scroll offset of the Help overlay.
    pub help_scroll: u16,
    /// Largest useful Help scroll offset, set by the renderer after wrapping.
    pub help_max_scroll: std::cell::Cell<u16>,
    /// Read-only F6 diagnostics state.
    pub diagnostics: DiagnosticsState,
    /// Current vertical scroll offset of the diagnostics overlay.
    pub diagnostics_scroll: u16,
    /// Largest useful diagnostics scroll offset, set by the renderer.
    pub diagnostics_max_scroll: std::cell::Cell<u16>,
    /// Read-only F7 diagnostics for the selected Teams 1:1 contact.
    pub contact_diagnostics: ContactDiagnosticsState,
    /// Current vertical scroll offset of the contact diagnostics overlay.
    pub contact_diagnostics_scroll: u16,
    /// Largest useful contact diagnostics scroll offset, set by the renderer.
    pub contact_diagnostics_max_scroll: std::cell::Cell<u16>,
    /// Raw selected chat id used only for fresh read-only F7 probes; never rendered/exported.
    contact_diagnostics_chat_id: Option<String>,
    /// Raw Graph user id used only for the live F7 probe; never rendered/exported.
    contact_diagnostics_user_id: Option<String>,
    /// Raw Teams user MRI candidate retained only for read-only UPS probing.
    /// The value is never rendered or exported.
    contact_diagnostics_mri: Option<String>,
    /// Raw lookup address used only by the read-only Middle Tier diagnostics.
    /// The value is never rendered, logged, or exported.
    contact_diagnostics_lookup_address: Option<String>,
    /// Contact profile currently shown by Teams `g`.
    pub contact_profile: Option<ContactProfileState>,
    /// Vertical scroll offset for the contact profile.
    pub contact_profile_scroll: u16,
    /// Maximum useful vertical contact-profile scroll offset.
    pub contact_profile_max_scroll: std::cell::Cell<u16>,
    /// Generation token used to invalidate delayed automatic read timers.
    read_timer_generation: u64,
    /// A new Teams chat message arrived since Teams was last opened.
    pub teams_unread: bool,
    /// Whether the automatic Teams conversation-cache warm-up has been queued.
    chat_cache_warmup_started: bool,
    /// Borderless full-width view for clean terminal text selection.
    pub copy_mode: bool,
    pub copy_scroll: u16,
    /// When true, Outlook and Teams panes stack top-to-bottom instead of
    /// left-to-right.
    pub panes_vertical: bool,
    /// Leave the TUI and open `$EDITOR` on the compose subject and body.
    pub pending_external_edit: bool,
    pub should_quit: bool,
    pub settings: crate::settings::Settings,
    /// Terminal graphics protocol, if the terminal can draw pixels.
    pub graphics: Option<crate::termimg::Graphics>,
    pub image_cache: std::collections::HashMap<String, crate::termimg::ReadyImage>,
    pub image_loading: std::collections::HashSet<String>,
    pub image_failed: std::collections::HashSet<String>,
    /// Contacts learned only from mail already loaded during this session.
    mail_contacts: std::collections::BTreeMap<String, MailContact>,
}

/// Palette command identifiers.
const PALETTE_COMMANDS: &[(&str, &str)] = &[
    ("outlook", "Switch to Outlook"),
    ("teams", "Switch to Teams"),
    ("calendar", "Switch to Calendar"),
    ("compose", "Compose new mail"),
    ("reply", "Reply to the selected mail"),
    ("reply-all", "Reply-all to the selected mail"),
    ("forward", "Forward the selected mail"),
    ("mark-read", "Toggle the selected mail's read/unread state"),
    ("move-mail", "Move the selected mail to another folder"),
    ("trash", "Move the selected mail to Deleted Items"),
    ("layout", "Toggle horizontal/vertical pane layout"),
    ("threads", "Toggle mail threads/individual messages"),
    ("chat-sender", "Teams: chat with selected email's sender"),
    ("new-chat", "Teams: find a person and start a chat"),
    ("refresh", "Refresh current view"),
    ("diagnostics", "Show diagnostics"),
    ("help", "Show help"),
    ("settings", "Open settings"),
    ("quit", "Quit"),
];

const MAX_MAIL_IMAGES: usize = 4;
const MAX_MAIL_IMAGE_BYTES: usize = 8 * 1024 * 1024;
const MAX_MAIL_IMAGE_WIDTH: u32 = 1600;
/// Normal network poll cadence. The timer still ticks at this rate while
/// App-level backoff decides whether another Graph poll is actually started.
pub const POLL_SECONDS: u64 = 20;
/// Local scheduler cadence for accelerated Teams polling.
pub const HOT_POLL_TICK_SECONDS: u64 = 1;
/// Activity ages defining the visual/polling tiers.
const TEAMS_HOT_AGE_SECONDS: u64 = 2 * 60;
const TEAMS_WARM_AGE_SECONDS: u64 = 5 * 60;
const TEAMS_COOL_AGE_SECONDS: u64 = 15 * 60;
/// Background Teams GETs are deliberately prevented from occupying all Graph work.
const TEAMS_BACKGROUND_CONCURRENCY: usize = 2;
/// Keep background GETs for one chat comfortably below the per-chat ceiling.
const TEAMS_BACKGROUND_PER_CHAT_MIN_SECONDS: u64 = 2;
/// Foreground refreshes may be quicker, but still avoid back-to-back per-chat GETs.
const TEAMS_FOREGROUND_PER_CHAT_MIN_SECONDS: u64 = 1;
/// After any observed Graph 429, recover the Teams background budget slowly.
const TEAMS_BACKGROUND_RECOVERY_SECONDS: u64 = 30;
const TEAMS_BACKGROUND_RECOVERY_STEP_RPS: f64 = 0.5;
/// A successful Graph folder refresh older than this is visibly stale.
pub const POLL_STALE_SECONDS: u64 = 2 * 60;
/// Maximum poll retry interval after repeated late/failed poll health checks.
const POLL_BACKOFF_MAX_SECONDS: u64 = 5 * 60;
/// Choices shown by the F4 ntfy snooze picker.
pub const NTFY_SNOOZE_HOURS: &[u64] = &[1, 2, 4, 8, 12, 24];
const NTFY_SNOOZE_STATE_FILE: &str = "ntfy-snooze";
const UI_STATE_FILE: &str = "ui-state";
const STAY_ARCHIVED_STATE_FILE: &str = "stay-archived.json";
const STAY_ARCHIVED_STATE_VERSION: u64 = 1;
const MAX_MAIL_IMAGE_HEIGHT: u32 = 1200;
const MAX_TEAMS_IMAGES: usize = MAX_MAIL_IMAGES;
const MAX_TEAMS_IMAGE_BYTES: usize = MAX_MAIL_IMAGE_BYTES;
const MAX_TEAMS_IMAGE_WIDTH: u32 = MAX_MAIL_IMAGE_WIDTH;
const MAX_TEAMS_IMAGE_HEIGHT: u32 = MAX_MAIL_IMAGE_HEIGHT;

fn is_supported_teams_image_name(name: &str) -> bool {
    name.rsplit_once('.')
        .map(|(_, extension)| {
            extension.eq_ignore_ascii_case("png")
                || extension.eq_ignore_ascii_case("jpg")
                || extension.eq_ignore_ascii_case("jpeg")
                || extension.eq_ignore_ascii_case("gif")
        })
        .unwrap_or(false)
}

const TEAMS_DISK_CACHE_KEY_HEX_LEN: usize = 32;
const TEAMS_CONVERSATION_CACHE_DIR: &str = "conversations";
const TEAMS_CONVERSATION_CACHE_VERSION: u64 = 2;
const MAX_TEAMS_CONVERSATION_CACHE_BYTES: u64 = 16 * 1024 * 1024;
const MAX_TEAMS_CACHED_MESSAGES: usize = 2000;
static TEAMS_CACHE_TEMP_SEQ: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

fn teams_conversation_cache_key(chat_id: &str) -> String {
    fn update(mut hash: u64, bytes: &[u8]) -> u64 {
        const PRIME: u64 = 0x0000_0100_0000_01b3;
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(PRIME);
        }
        hash
    }

    let first = update(0xcbf2_9ce4_8422_2325, chat_id.as_bytes());
    let second = update(0x8422_2325_cbf2_9ce4, chat_id.as_bytes());
    format!("{first:016x}{second:016x}")
}

fn teams_conversation_cache_path(
    cache_root: &std::path::Path,
    chat_id: &str,
) -> std::path::PathBuf {
    cache_root
        .join(TEAMS_CONVERSATION_CACHE_DIR)
        .join(format!("{}.json", teams_conversation_cache_key(chat_id)))
}

fn load_teams_conversation_cache(
    cache_root: &std::path::Path,
    chat_id: &str,
) -> std::io::Result<Option<(Vec<ChatMessage>, Option<String>)>> {
    let path = teams_conversation_cache_path(cache_root, chat_id);
    let metadata = match std::fs::metadata(&path) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(error),
    };

    if !metadata.is_file() || metadata.len() > MAX_TEAMS_CONVERSATION_CACHE_BYTES {
        let _ = std::fs::remove_file(&path);
        return Ok(None);
    }

    let bytes = std::fs::read(&path)?;
    let value: serde_json::Value = match serde_json::from_slice(&bytes) {
        Ok(value) => value,
        Err(_) => {
            let _ = std::fs::remove_file(&path);
            return Ok(None);
        }
    };

    if value.get("version").and_then(serde_json::Value::as_u64)
        != Some(TEAMS_CONVERSATION_CACHE_VERSION)
        || value.get("chat_id").and_then(serde_json::Value::as_str) != Some(chat_id)
    {
        let _ = std::fs::remove_file(&path);
        return Ok(None);
    }

    // Graph continuation links are intentionally not restored from disk.
    let next = None;

    let Some(messages_value) = value.get("messages").cloned() else {
        let _ = std::fs::remove_file(&path);
        return Ok(None);
    };
    let messages: Vec<ChatMessage> =
        match serde_json::from_value::<Vec<ChatMessage>>(messages_value) {
            Ok(messages) if messages.len() <= MAX_TEAMS_CACHED_MESSAGES => messages,
            _ => {
                let _ = std::fs::remove_file(&path);
                return Ok(None);
            }
        };

    Ok(Some((messages, next)))
}

fn store_teams_conversation_cache(
    cache_root: &std::path::Path,
    chat_id: &str,
    messages: &[ChatMessage],
    max_bytes: u64,
) -> std::io::Result<()> {
    let cache_dir = cache_root.join(TEAMS_CONVERSATION_CACHE_DIR);
    let final_path = teams_conversation_cache_path(cache_root, chat_id);

    if messages.is_empty() {
        match std::fs::remove_file(&final_path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => return Err(error),
        }
        return Ok(());
    }

    create_private_cache_dir(&cache_dir)?;

    // Keep the newest bounded slice instead of dropping the cache completely
    // after a user has paged far back in a long-running conversation.
    let start = messages.len().saturating_sub(MAX_TEAMS_CACHED_MESSAGES);
    let newest_first: Vec<&ChatMessage> = messages[start..].iter().rev().collect();
    let value = serde_json::json!({
        "version": TEAMS_CONVERSATION_CACHE_VERSION,
        "chat_id": chat_id,
        "messages": newest_first,
    });
    let bytes = serde_json::to_vec(&value).map_err(std::io::Error::other)?;

    if bytes.len() as u64 > MAX_TEAMS_CONVERSATION_CACHE_BYTES {
        let _ = std::fs::remove_file(final_path);
        return Ok(());
    }

    let temp_seq = TEAMS_CACHE_TEMP_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let temp_path = cache_dir.join(format!(
        "{}.json.tmp-{}-{temp_seq}",
        teams_conversation_cache_key(chat_id),
        std::process::id()
    ));
    write_private_cache_file(&temp_path, &bytes)?;

    if final_path.exists() {
        std::fs::remove_file(&final_path)?;
    }
    std::fs::rename(temp_path, final_path)?;
    prune_persistent_cache(cache_root, max_bytes)
}

fn teams_disk_cache_key(key: &TeamsImageKey) -> String {
    fn update(mut hash: u64, bytes: &[u8]) -> u64 {
        const PRIME: u64 = 0x0000_0100_0000_01b3;
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(PRIME);
        }
        hash
    }

    fn feed(hash: u64, part: &str) -> u64 {
        let hash = update(hash, part.as_bytes());
        update(hash, &[0])
    }

    // Two deterministic FNV-1a streams give a compact 128-bit cache filename
    // without introducing a hashing crate or exposing chat/file names.
    let mut first = 0xcbf2_9ce4_8422_2325;
    let mut second = 0x8422_2325_cbf2_9ce4;

    match &key.source {
        TeamsImageSource::Chat(chat_id) => {
            first = feed(first, "chat");
            first = feed(first, chat_id);
            second = feed(second, "chat");
            second = feed(second, chat_id);
        }
        TeamsImageSource::Channel {
            team_id,
            channel_id,
        } => {
            first = feed(first, "channel");
            first = feed(first, team_id);
            first = feed(first, channel_id);
            second = feed(second, "channel");
            second = feed(second, channel_id);
            second = feed(second, team_id);
        }
    }

    first = feed(first, &key.message_id);
    second = feed(second, &key.message_id);
    format!("{first:016x}{second:016x}")
}

#[cfg(unix)]
fn create_private_cache_dir(path: &std::path::Path) -> std::io::Result<()> {
    use std::os::unix::fs::{DirBuilderExt, PermissionsExt};

    let mut builder = std::fs::DirBuilder::new();
    builder.recursive(true);
    builder.mode(0o700);
    builder.create(path)?;

    let mut permissions = std::fs::metadata(path)?.permissions();
    permissions.set_mode(0o700);
    std::fs::set_permissions(path, permissions)
}

#[cfg(not(unix))]
fn create_private_cache_dir(path: &std::path::Path) -> std::io::Result<()> {
    std::fs::create_dir_all(path)
}

#[cfg(unix)]
fn write_private_cache_file(path: &std::path::Path, bytes: &[u8]) -> std::io::Result<()> {
    use std::io::Write;
    use std::os::unix::fs::OpenOptionsExt;

    let mut file = std::fs::OpenOptions::new()
        .create(true)
        .truncate(true)
        .write(true)
        .mode(0o600)
        .open(path)?;
    file.write_all(bytes)?;
    file.flush()
}

#[cfg(not(unix))]
fn write_private_cache_file(path: &std::path::Path, bytes: &[u8]) -> std::io::Result<()> {
    use std::io::Write;

    let mut file = std::fs::OpenOptions::new()
        .create(true)
        .truncate(true)
        .write(true)
        .open(path)?;
    file.write_all(bytes)?;
    file.flush()
}

fn calendar_month_view_available() -> bool {
    crossterm::terminal::size()
        .map(|(width, height)| {
            width >= CALENDAR_MONTH_MIN_WIDTH && height >= CALENDAR_MONTH_MIN_HEIGHT
        })
        .unwrap_or(false)
}

fn load_ui_state(cache_dir: Option<&std::path::Path>) -> UiState {
    let Some(cache_dir) = cache_dir else {
        return UiState::default();
    };

    let path = cache_dir.join(UI_STATE_FILE);
    let text = match std::fs::read_to_string(&path) {
        Ok(text) => text,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            return UiState::default();
        }
        Err(error) => {
            tracing::warn!("could not read UI state {}: {error}", path.display());
            return UiState::default();
        }
    };

    let mut state = UiState::default();
    for line in text.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        match key.trim() {
            "screen" => {
                state.screen = match value.trim() {
                    "outlook" => Screen::Outlook,
                    "teams" => Screen::Teams,
                    "calendar" => Screen::Calendar,
                    _ => state.screen,
                };
            }
            "calendar_days" => {
                if let Ok(days) = value.trim().parse::<i64>() {
                    if CALENDAR_RANGES.contains(&days) {
                        state.calendar_days = days;
                    }
                }
            }
            "calendar_view" => {
                state.calendar_view = match value.trim() {
                    "agenda" => CalendarView::Agenda,
                    "month" => CalendarView::Month,
                    _ => state.calendar_view,
                };
            }
            "teams_chat_id" => {
                let value = value.trim();
                state.teams_chat_id = (!value.is_empty()).then(|| value.to_string());
            }
            "teams_archive_expanded" => {
                state.teams_archive_expanded = value.trim().eq_ignore_ascii_case("true");
            }
            _ => {}
        }
    }

    state
}

fn store_ui_state(
    cache_dir: &std::path::Path,
    screen: Screen,
    calendar_days: i64,
    calendar_view: CalendarView,
    teams_chat_id: Option<&str>,
    teams_archive_expanded: bool,
) -> std::io::Result<()> {
    create_private_cache_dir(cache_dir)?;

    let screen = match screen {
        Screen::Outlook => "outlook",
        Screen::Teams => "teams",
        Screen::Calendar => "calendar",
    };
    let calendar_view = match calendar_view {
        CalendarView::Agenda => "agenda",
        CalendarView::Month => "month",
    };
    let mut data = format!(
        "screen={screen}\ncalendar_days={calendar_days}\ncalendar_view={calendar_view}\nteams_archive_expanded={teams_archive_expanded}\n"
    );
    if let Some(chat_id) = teams_chat_id.filter(|value| !value.is_empty()) {
        data.push_str("teams_chat_id=");
        data.push_str(chat_id);
        data.push('\n');
    }
    let final_path = cache_dir.join(UI_STATE_FILE);
    let temp_path = cache_dir.join(format!("{UI_STATE_FILE}.tmp-{}", std::process::id()));

    write_private_cache_file(&temp_path, data.as_bytes())?;
    if final_path.exists() {
        std::fs::remove_file(&final_path)?;
    }
    std::fs::rename(temp_path, final_path)
}

fn load_stay_archived(cache_dir: Option<&std::path::Path>) -> std::collections::HashSet<String> {
    let Some(cache_dir) = cache_dir else {
        return std::collections::HashSet::new();
    };

    let path = cache_dir.join(STAY_ARCHIVED_STATE_FILE);
    let text = match std::fs::read_to_string(&path) {
        Ok(text) => text,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            return std::collections::HashSet::new();
        }
        Err(error) => {
            tracing::warn!(
                "could not read Stay archived state {}: {error}",
                path.display()
            );
            return std::collections::HashSet::new();
        }
    };

    let value: serde_json::Value = match serde_json::from_str(&text) {
        Ok(value) => value,
        Err(error) => {
            tracing::warn!(
                "could not parse Stay archived state {}: {error}",
                path.display()
            );
            return std::collections::HashSet::new();
        }
    };

    if value.get("version").and_then(serde_json::Value::as_u64) != Some(STAY_ARCHIVED_STATE_VERSION)
    {
        tracing::warn!(
            "unsupported Stay archived state version in {}",
            path.display()
        );
        return std::collections::HashSet::new();
    }

    value
        .get("chats")
        .and_then(serde_json::Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(serde_json::Value::as_str)
        .map(str::trim)
        .filter(|chat_id| !chat_id.is_empty())
        .map(str::to_string)
        .collect()
}

fn store_stay_archived(
    cache_dir: &std::path::Path,
    stay_archived: &std::collections::HashSet<String>,
) -> std::io::Result<()> {
    create_private_cache_dir(cache_dir)?;

    let mut chats: Vec<&str> = stay_archived.iter().map(String::as_str).collect();
    chats.sort_unstable();

    let value = serde_json::json!({
        "version": STAY_ARCHIVED_STATE_VERSION,
        "chats": chats,
    });
    let bytes = serde_json::to_vec_pretty(&value).map_err(std::io::Error::other)?;

    let final_path = cache_dir.join(STAY_ARCHIVED_STATE_FILE);
    let temp_path = cache_dir.join(format!(
        "{STAY_ARCHIVED_STATE_FILE}.tmp-{}",
        std::process::id()
    ));

    write_private_cache_file(&temp_path, &bytes)?;
    if final_path.exists() {
        std::fs::remove_file(&final_path)?;
    }
    std::fs::rename(temp_path, final_path)
}

fn ntfy_snooze_state_path(token_cache_path: &std::path::Path) -> std::path::PathBuf {
    // Keep persistent application state beside the already-established token
    // cache. This also honours a custom M365_TOKEN_CACHE path automatically.
    token_cache_path.with_file_name(NTFY_SNOOZE_STATE_FILE)
}

fn legacy_ntfy_snooze_state_path() -> Option<std::path::PathBuf> {
    let root = std::env::var_os("XDG_STATE_HOME")
        .filter(|value| !value.is_empty())
        .map(std::path::PathBuf::from)
        .or_else(|| {
            std::env::var_os("HOME")
                .filter(|value| !value.is_empty())
                .map(std::path::PathBuf::from)
                .map(|home| home.join(".local").join("state"))
        })?;

    Some(root.join("m365-tui").join(NTFY_SNOOZE_STATE_FILE))
}

fn ntfy_snooze_remaining_at(
    until: Option<std::time::SystemTime>,
    now: std::time::SystemTime,
) -> Option<std::time::Duration> {
    let remaining = until?.duration_since(now).ok()?;
    (!remaining.is_zero()).then_some(remaining)
}

fn remove_file_if_present(path: &std::path::Path) -> std::io::Result<()> {
    match std::fs::remove_file(path) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error),
    }
}

fn load_ntfy_snooze_file(
    path: &std::path::Path,
    now: std::time::SystemTime,
) -> Option<std::time::SystemTime> {
    let text = match std::fs::read_to_string(path) {
        Ok(text) => text,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return None,
        Err(error) => {
            tracing::warn!(
                "could not read ntfy snooze state {}: {error}",
                path.display()
            );
            return None;
        }
    };

    let epoch_seconds = match text.trim().parse::<u64>() {
        Ok(value) => value,
        Err(error) => {
            tracing::warn!(
                "invalid ntfy snooze state {}: {error}; discarding it",
                path.display()
            );
            let _ = remove_file_if_present(path);
            return None;
        }
    };

    let Some(until) = std::time::SystemTime::UNIX_EPOCH
        .checked_add(std::time::Duration::from_secs(epoch_seconds))
    else {
        tracing::warn!(
            "invalid ntfy snooze deadline in {}; discarding it",
            path.display()
        );
        let _ = remove_file_if_present(path);
        return None;
    };

    if ntfy_snooze_remaining_at(Some(until), now).is_none() {
        let _ = remove_file_if_present(path);
        return None;
    }

    Some(until)
}

fn store_ntfy_snooze_until(
    token_cache_path: &std::path::Path,
    until: std::time::SystemTime,
) -> std::io::Result<()> {
    let path = ntfy_snooze_state_path(token_cache_path);
    let Some(parent) = path.parent() else {
        return Ok(());
    };
    create_private_cache_dir(parent)?;

    let epoch_seconds = until
        .duration_since(std::time::SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let data = format!(
        "{epoch_seconds}
"
    );

    // Write through a private temporary file, then replace the final state.
    // An application crash therefore cannot leave a partially written deadline.
    let temp_path = path.with_extension(format!("tmp-{}", std::process::id()));
    write_private_cache_file(&temp_path, data.as_bytes())?;

    #[cfg(unix)]
    {
        std::fs::rename(&temp_path, &path)?;
    }

    #[cfg(not(unix))]
    {
        if path.exists() {
            remove_file_if_present(&path)?;
        }
        std::fs::rename(&temp_path, &path)?;
    }

    Ok(())
}

fn clear_ntfy_snooze_state(token_cache_path: &std::path::Path) -> std::io::Result<()> {
    let primary = ntfy_snooze_state_path(token_cache_path);
    remove_file_if_present(&primary)?;

    // Also clear the short-lived legacy location so Resume now cannot cause a
    // previously persisted snooze to reappear after the next restart.
    if let Some(legacy) = legacy_ntfy_snooze_state_path() {
        if legacy != primary {
            remove_file_if_present(&legacy)?;
        }
    }

    Ok(())
}

fn load_ntfy_snooze_until(token_cache_path: &std::path::Path) -> Option<std::time::SystemTime> {
    let now = std::time::SystemTime::now();
    let primary = ntfy_snooze_state_path(token_cache_path);

    if let Some(until) = load_ntfy_snooze_file(&primary, now) {
        if let Some(remaining) = ntfy_snooze_remaining_at(Some(until), now) {
            tracing::info!(
                "restored ntfy snooze from {} ({}m remaining)",
                primary.display(),
                remaining.as_secs().div_ceil(60)
            );
        }
        return Some(until);
    }

    let legacy = legacy_ntfy_snooze_state_path()?;
    if legacy == primary {
        return None;
    }

    let until = load_ntfy_snooze_file(&legacy, now)?;

    match store_ntfy_snooze_until(token_cache_path, until) {
        Ok(()) => {
            if let Err(error) = remove_file_if_present(&legacy) {
                tracing::warn!(
                    "ntfy snooze migrated but old state {} could not be removed: {error}",
                    legacy.display()
                );
            } else {
                tracing::info!(
                    "migrated ntfy snooze state from {} to {}",
                    legacy.display(),
                    primary.display()
                );
            }
        }
        Err(error) => {
            // Keep using the valid legacy deadline in memory. The legacy file
            // stays in place so migration can be retried on the next launch.
            tracing::warn!(
                "could not migrate ntfy snooze state from {} to {}: {error}",
                legacy.display(),
                primary.display()
            );
        }
    }

    Some(until)
}

fn next_poll_retry_interval(current: u64) -> u64 {
    current
        .saturating_mul(2)
        .clamp(POLL_SECONDS, POLL_BACKOFF_MAX_SECONDS)
}

fn remove_teams_disk_cache_entry(dir: &std::path::Path, cache_key: &str) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };

    let image_prefix = format!("{cache_key}-");
    let meta_name = format!("{cache_key}.meta");

    for entry in entries.flatten() {
        let name = entry.file_name();
        let name = name.to_string_lossy();
        let matches = name == meta_name
            || name.starts_with(&format!("{cache_key}.meta.tmp-"))
            || (name.starts_with(&image_prefix)
                && (name.ends_with(".png") || name.contains(".tmp-")));
        if matches {
            let _ = std::fs::remove_file(entry.path());
        }
    }
}

fn load_teams_disk_cache(
    dir: &std::path::Path,
    cache_key: &str,
) -> std::io::Result<Option<Vec<image::DynamicImage>>> {
    let meta_path = dir.join(format!("{cache_key}.meta"));
    let count_text = match std::fs::read_to_string(&meta_path) {
        Ok(value) => value,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(error),
    };

    let count = match count_text.trim().parse::<usize>() {
        Ok(value) if (1..=MAX_TEAMS_IMAGES).contains(&value) => value,
        _ => {
            remove_teams_disk_cache_entry(dir, cache_key);
            return Ok(None);
        }
    };

    let mut images = Vec::with_capacity(count);
    for index in 0..count {
        let path = dir.join(format!("{cache_key}-{index}.png"));
        let bytes = match std::fs::read(&path) {
            Ok(bytes) if bytes.len() <= MAX_TEAMS_IMAGE_BYTES => bytes,
            _ => {
                remove_teams_disk_cache_entry(dir, cache_key);
                return Ok(None);
            }
        };

        match image::load_from_memory(&bytes) {
            Ok(image)
                if image.width() > 1
                    && image.height() > 1
                    && image.width() <= MAX_TEAMS_IMAGE_WIDTH
                    && image.height() <= MAX_TEAMS_IMAGE_HEIGHT =>
            {
                images.push(image);
            }
            _ => {
                remove_teams_disk_cache_entry(dir, cache_key);
                return Ok(None);
            }
        }
    }

    // Refresh the metadata mtime on a successful hit so pruning behaves as LRU
    // without another filesystem dependency.
    write_private_cache_file(&meta_path, format!("{count}\n").as_bytes())?;
    Ok(Some(images))
}

fn prune_persistent_cache(cache_root: &std::path::Path, max_bytes: u64) -> std::io::Result<()> {
    use std::collections::HashMap;
    use std::time::SystemTime;

    #[derive(Debug)]
    enum Kind {
        Image(String),
        Conversation(std::path::PathBuf),
    }

    #[derive(Default)]
    struct ImageEntry {
        size: u64,
        modified: Option<SystemTime>,
    }

    struct Entry {
        kind: Kind,
        size: u64,
        modified: SystemTime,
    }

    let mut image_groups: HashMap<String, ImageEntry> = HashMap::new();
    if let Ok(items) = std::fs::read_dir(cache_root) {
        for item in items.flatten() {
            let metadata = match item.metadata() {
                Ok(metadata) if metadata.is_file() => metadata,
                _ => continue,
            };
            let name = item.file_name();
            let name = name.to_string_lossy();
            if name.len() < TEAMS_DISK_CACHE_KEY_HEX_LEN {
                continue;
            }
            let key = &name[..TEAMS_DISK_CACHE_KEY_HEX_LEN];
            if !key.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                continue;
            }
            let delimiter = name.as_bytes().get(TEAMS_DISK_CACHE_KEY_HEX_LEN).copied();
            if !matches!(delimiter, Some(b'.') | Some(b'-')) {
                continue;
            }

            let entry = image_groups.entry(key.to_string()).or_default();
            entry.size = entry.size.saturating_add(metadata.len());
            if let Ok(modified) = metadata.modified() {
                if entry.modified.is_none_or(|current| modified > current) {
                    entry.modified = Some(modified);
                }
            }
        }
    }

    let mut entries = Vec::new();
    for (key, entry) in image_groups {
        entries.push(Entry {
            kind: Kind::Image(key),
            size: entry.size,
            modified: entry.modified.unwrap_or(SystemTime::UNIX_EPOCH),
        });
    }

    let conversations = cache_root.join(TEAMS_CONVERSATION_CACHE_DIR);
    if let Ok(items) = std::fs::read_dir(&conversations) {
        for item in items.flatten() {
            let metadata = match item.metadata() {
                Ok(metadata) if metadata.is_file() => metadata,
                _ => continue,
            };
            let name = item.file_name();
            let name = name.to_string_lossy();
            let Some(key) = name.strip_suffix(".json") else {
                continue;
            };
            if key.len() != TEAMS_DISK_CACHE_KEY_HEX_LEN
                || !key.bytes().all(|byte| byte.is_ascii_hexdigit())
            {
                continue;
            }
            entries.push(Entry {
                kind: Kind::Conversation(item.path()),
                size: metadata.len(),
                modified: metadata.modified().unwrap_or(SystemTime::UNIX_EPOCH),
            });
        }
    }

    if let Ok(items) = std::fs::read_dir(cache_root.join("images")) {
        for item in items.flatten() {
            if let Ok(metadata) = item.metadata() {
                if metadata.is_file() {
                    entries.push(Entry {
                        kind: Kind::Conversation(item.path()),
                        size: metadata.len(),
                        modified: metadata.modified().unwrap_or(SystemTime::UNIX_EPOCH),
                    });
                }
            }
        }
    }
    let mut total = entries
        .iter()
        .fold(0u64, |sum, entry| sum.saturating_add(entry.size));
    if total <= max_bytes {
        return Ok(());
    }

    entries.sort_by_key(|entry| entry.modified);
    for entry in entries {
        if total <= max_bytes {
            break;
        }
        match entry.kind {
            Kind::Image(key) => remove_teams_disk_cache_entry(cache_root, &key),
            Kind::Conversation(path) => {
                let _ = std::fs::remove_file(path);
            }
        }
        total = total.saturating_sub(entry.size);
    }

    Ok(())
}

fn store_teams_disk_cache(
    dir: &std::path::Path,
    cache_key: &str,
    images: &[image::DynamicImage],
    max_bytes: u64,
) -> std::io::Result<()> {
    use std::io::Cursor;

    if images.is_empty() {
        return Ok(());
    }

    create_private_cache_dir(dir)?;

    let mut encoded = Vec::with_capacity(images.len());
    for image in images.iter().take(MAX_TEAMS_IMAGES) {
        let mut cursor = Cursor::new(Vec::new());
        image
            .write_to(&mut cursor, image::ImageFormat::Png)
            .map_err(|error| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string())
            })?;
        let bytes = cursor.into_inner();
        if bytes.len() > MAX_TEAMS_IMAGE_BYTES {
            return Ok(());
        }
        encoded.push(bytes);
    }

    if encoded.is_empty() {
        return Ok(());
    }

    remove_teams_disk_cache_entry(dir, cache_key);

    let pid = std::process::id();
    for (index, bytes) in encoded.iter().enumerate() {
        let final_path = dir.join(format!("{cache_key}-{index}.png"));
        let temp_path = dir.join(format!("{cache_key}-{index}.png.tmp-{pid}"));
        write_private_cache_file(&temp_path, bytes)?;
        if final_path.exists() {
            std::fs::remove_file(&final_path)?;
        }
        std::fs::rename(&temp_path, &final_path)?;
    }

    let meta_path = dir.join(format!("{cache_key}.meta"));
    let temp_meta = dir.join(format!("{cache_key}.meta.tmp-{pid}"));
    write_private_cache_file(&temp_meta, format!("{}\n", encoded.len()).as_bytes())?;
    if meta_path.exists() {
        std::fs::remove_file(&meta_path)?;
    }
    std::fs::rename(&temp_meta, &meta_path)?;

    prune_persistent_cache(dir, max_bytes)
}

fn reference_image_files(message: &ChatMessage) -> Vec<(String, String)> {
    message
        .attachments
        .iter()
        .filter(|attachment| {
            attachment
                .content_type
                .as_deref()
                .is_some_and(|kind| kind.eq_ignore_ascii_case("reference"))
        })
        .filter_map(|attachment| {
            let name = attachment.name.as_deref()?;
            if !is_supported_teams_image_name(name) {
                return None;
            }

            Some((name.to_string(), attachment.content_url.clone()?))
        })
        .take(MAX_TEAMS_IMAGES)
        .collect()
}

fn hosted_content_ids(message: &ChatMessage) -> Vec<String> {
    fn collect(text: &str, ids: &mut Vec<String>) {
        const MARKER: &str = "hostedContents/";
        let mut rest = text;

        while let Some(pos) = rest.find(MARKER) {
            let after = &rest[pos + MARKER.len()..];
            let end = after
                .find(|c: char| {
                    c == '/' || c == '"' || c == '\'' || c == '<' || c == '>' || c.is_whitespace()
                })
                .unwrap_or(after.len());

            if end > 0 {
                let id = &after[..end];
                if !ids.iter().any(|known| known == id) {
                    ids.push(id.to_string());
                }
            }

            rest = &after[end..];
            if rest.is_empty() {
                break;
            }
        }
    }

    let mut ids = Vec::new();

    if let Some(body) = message
        .body
        .as_ref()
        .and_then(|body| body.content.as_deref())
    {
        collect(body, &mut ids);
    }

    for attachment in &message.attachments {
        if let Some(content) = attachment.content.as_deref() {
            collect(content, &mut ids);
        }
        if let Some(url) = attachment.content_url.as_deref() {
            collect(url, &mut ids);
        }
    }

    ids.truncate(MAX_TEAMS_IMAGES);
    ids
}

fn calendar_month_first(offset: i32) -> chrono::NaiveDate {
    let today = chrono::Local::now().date_naive();
    let month_index =
        chrono::Datelike::year(&today) * 12 + chrono::Datelike::month0(&today) as i32 + offset;
    let year = month_index.div_euclid(12);
    let month = month_index.rem_euclid(12) as u32 + 1;
    chrono::NaiveDate::from_ymd_opt(year, month, 1).expect("valid calendar month")
}

fn calendar_state_local_datetime(
    value: &m365_core::models::DateTimeTimeZone,
) -> Option<chrono::DateTime<chrono::Local>> {
    if let Ok(parsed) = chrono::DateTime::parse_from_rfc3339(&value.date_time) {
        return Some(parsed.with_timezone(&chrono::Local));
    }

    let naive = chrono::NaiveDateTime::parse_from_str(&value.date_time, "%Y-%m-%dT%H:%M:%S%.f")
        .or_else(|_| chrono::NaiveDateTime::parse_from_str(&value.date_time, "%Y-%m-%dT%H:%M:%S"))
        .ok()?;

    Some(
        chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(naive, chrono::Utc)
            .with_timezone(&chrono::Local),
    )
}

fn ntfy_presence_is_away(presence: Option<&Presence>) -> bool {
    presence
        .and_then(|presence| presence.availability.as_deref())
        .is_some_and(|availability| availability.eq_ignore_ascii_case("Away"))
}

fn ntfy_mode_should_forward(
    mode: NtfyMode,
    presence: Option<&Presence>,
    working_now: Option<bool>,
) -> bool {
    match mode {
        NtfyMode::Never => false,
        NtfyMode::Always => true,
        NtfyMode::Away => ntfy_presence_is_away(presence),
        NtfyMode::AlwaysWd => working_now == Some(true),
        NtfyMode::AwayWd => working_now == Some(true) && ntfy_presence_is_away(presence),
    }
}

fn ntfy_tag_for_category(category: &str) -> Option<&'static str> {
    match category {
        "Mail" => Some("email"),
        "Teams" => Some("speech_balloon"),
        "Calendar" => Some("calendar"),
        _ => None,
    }
}

fn calendar_reminder_threshold_minutes(seconds_until: i64) -> Option<u16> {
    match seconds_until {
        1..=300 => Some(5),
        301..=900 => Some(15),
        _ => None,
    }
}

fn calendar_event_reminder(
    event: &CalEvent,
    now: &chrono::DateTime<chrono::Local>,
) -> Option<(u16, chrono::DateTime<chrono::Local>)> {
    if event.is_cancelled.unwrap_or(false) || event.is_all_day.unwrap_or(false) {
        return None;
    }
    if event
        .response_status
        .as_ref()
        .and_then(|status| status.response.as_deref())
        .is_some_and(|response| response.eq_ignore_ascii_case("declined"))
    {
        return None;
    }

    let start = event
        .start
        .as_ref()
        .and_then(calendar_state_local_datetime)?;
    let seconds_until = start.signed_duration_since(*now).num_seconds();
    let minutes = calendar_reminder_threshold_minutes(seconds_until)?;
    Some((minutes, start))
}

fn calendar_state_event_dates(event: &CalEvent) -> Option<(chrono::NaiveDate, chrono::NaiveDate)> {
    let start = event
        .start
        .as_ref()
        .and_then(calendar_state_local_datetime)?;
    let end = event.end.as_ref().and_then(calendar_state_local_datetime)?;
    let start_date = start.date_naive();
    let mut end_date = end.date_naive();

    if end_date > start_date
        && end.time() == chrono::NaiveTime::from_hms_opt(0, 0, 0).expect("valid midnight")
    {
        end_date -= chrono::Duration::days(1);
    }
    if end_date < start_date {
        end_date = start_date;
    }

    Some((start_date, end_date))
}

fn calendar_event_intersects_month(event: &CalEvent, month_offset: i32) -> bool {
    let Some((event_start, event_end)) = calendar_state_event_dates(event) else {
        return false;
    };
    let month_start = calendar_month_first(month_offset);
    let next_month = calendar_month_first(month_offset + 1);
    event_start < next_month && event_end >= month_start
}

impl App {
    pub fn new(session: Session, tx: mpsc::Sender<AppMessage>) -> Self {
        let mut ui_state = load_ui_state(session.config.cache_dir.as_deref());
        let stay_archived = load_stay_archived(session.config.cache_dir.as_deref());
        let month_view_fallback =
            ui_state.calendar_view == CalendarView::Month && !calendar_month_view_available();
        if month_view_fallback {
            ui_state.calendar_view = CalendarView::Agenda;
        }
        let image_picker = if std::io::stdin().is_terminal() && std::io::stdout().is_terminal() {
            match ratatui_image::picker::Picker::from_query_stdio() {
                Ok(picker)
                    if picker.protocol_type() == ratatui_image::picker::ProtocolType::Kitty =>
                {
                    tracing::debug!(
                        "Kitty graphics protocol detected; mail image previews enabled"
                    );
                    Some(picker)
                }
                Ok(picker) => {
                    tracing::debug!(
                        "terminal image protocol {:?} detected; keeping text fallback",
                        picker.protocol_type()
                    );
                    None
                }
                Err(error) => {
                    tracing::debug!(
                        "terminal image protocol detection failed ({error}); keeping text fallback"
                    );
                    None
                }
            }
        } else {
            None
        };

        let last_chat_id = ui_state.teams_chat_id.clone();
        let preview_chat_id = if ui_state.screen == Screen::Teams {
            last_chat_id.clone()
        } else {
            None
        };
        let teams = TeamsState {
            // Persisted state restores context, not input focus. Keep the cursor
            // in the chat list and show the previous conversation as a local
            // cache preview until the user explicitly opens it.
            focus: TeamsFocus::List,
            preview_chat_id,
            last_chat_id,
            archive_expanded: ui_state.teams_archive_expanded,
            stay_archived,
            ..TeamsState::default()
        };

        let ntfy_snooze_until = load_ntfy_snooze_until(&session.config.token_cache_path);
        let teams_background_limiter = TeamsBackgroundLimiter::new(
            session.config.teams_poll_budget_rps,
            session.graph.throttle_generation(),
        );
        let presence_activity =
            crate::activity::ActivityMonitor::new(session.config.presence_activity_source);
        let presence_state = PresenceStateManager::new(session.config.presence_lock_restore);
        let teams_automation_peers = session.config.teams_automation_peers.clone();

        let mut app = Self {
            session,
            tx,
            screen: ui_state.screen,
            outlook: OutlookState::default(),
            image_picker,
            outlook_focus: OutlookFocus::Messages,
            teams,
            calendar: CalendarState {
                days: ui_state.calendar_days,
                view: ui_state.calendar_view,
                ..CalendarState::default()
            },
            calendar_reminder_events: Vec::new(),
            calendar_reminder_fired: std::collections::HashSet::new(),
            calendar_reminder_last_refresh: None,
            overlay: month_view_fallback.then(|| {
                Overlay::Notice(format!(
                    "Month view requires at least {}x{} characters.\n\nAgenda view was opened instead.",
                    CALENDAR_MONTH_MIN_WIDTH, CALENDAR_MONTH_MIN_HEIGHT
                ))
            }),
            notice_until: month_view_fallback.then(|| {
                std::time::Instant::now() + std::time::Duration::from_secs(3)
            }),
            status: "loading…".into(),
            me: None,
            my_presence: None,
            ntfy_working_now: None,
            ntfy_snooze_until,
            poll_last_success_at: std::time::Instant::now(),
            poll_health_pending: false,
            poll_retry_interval_secs: POLL_SECONDS,
            poll_backoff_until: None,
            push: PushState::Off,
            notified: std::collections::HashSet::new(),
            chat_seen: None,
            teams_hot_chat_state: std::collections::HashMap::new(),
            teams_background_limiter,
            chat_unread_refresh_running: false,
            teams_automation_peers,
            own_sent_chat_message_ids: std::collections::HashSet::new(),
            mail_seen: None,
            presence_state,
            presence_activity,
            presence_preference_pending: None,
            rss_kb: read_rss_kb(),
            last_status: String::new(),
            status_ticks: 0,
            text_width_hint: std::cell::Cell::new(60),
            reading_max_scroll: std::cell::Cell::new(0),
            help_scroll: 0,
            help_max_scroll: std::cell::Cell::new(0),
            diagnostics: DiagnosticsState::default(),
            diagnostics_scroll: 0,
            diagnostics_max_scroll: std::cell::Cell::new(0),
            contact_diagnostics: ContactDiagnosticsState::default(),
            contact_diagnostics_scroll: 0,
            contact_diagnostics_max_scroll: std::cell::Cell::new(0),
            contact_diagnostics_chat_id: None,
            contact_diagnostics_user_id: None,
            contact_diagnostics_mri: None,
            contact_diagnostics_lookup_address: None,
            contact_profile: None,
            contact_profile_scroll: 0,
            contact_profile_max_scroll: std::cell::Cell::new(0),
            read_timer_generation: 0,
            teams_unread: false,
            chat_cache_warmup_started: false,
            copy_mode: false,
            copy_scroll: 0,
            panes_vertical: false,
            pending_external_edit: false,
            should_quit: false,
            settings: crate::settings::Settings::load(),
            graphics: None,
            image_cache: Default::default(),
            image_loading: Default::default(),
            image_failed: Default::default(),
            mail_contacts: Default::default(),
            images_max_scroll: Default::default(),
            copy_max_scroll: Default::default(),
        };

        if app.screen == Screen::Teams {
            if let Some(chat_id) = app.teams.preview_chat_id.clone() {
                let _ = app.restore_teams_conversation_cache(&chat_id);
            }
        }

        app
    }

    pub fn bootstrap(&mut self) {
        self.load_whoami();
        self.load_presence();
        self.refresh_ntfy_work_plan();
        self.start_primary_presence();
        self.load_folders();
        self.load_chats();
        self.refresh_calendar_reminders();
        if self.screen == Screen::Teams {
            if let Some(chat_id) = self.teams.open_chat_id.clone() {
                self.teams.chat_open_pending_read = Some(chat_id.clone());
                let mode = if self.teams.messages.is_empty() {
                    ListUpdate::Replace
                } else {
                    ListUpdate::Merge
                };
                self.load_chat_messages(chat_id, mode);
            }
        }
        if self.screen == Screen::Calendar {
            self.load_calendar();
        }
    }

    fn persist_ui_state(&self) {
        let Some(cache_dir) = self.session.config.cache_dir.as_deref() else {
            return;
        };

        if let Err(error) = store_ui_state(
            cache_dir,
            self.screen,
            self.calendar.days,
            self.calendar.view,
            self.teams.last_chat_id.as_deref(),
            self.teams.archive_expanded,
        ) {
            tracing::warn!(
                "could not persist UI state in {}: {error}",
                cache_dir.display()
            );
        }
    }

    fn persist_stay_archived(&self) -> Option<String> {
        let cache_dir = self.session.config.cache_dir.as_deref()?;

        match store_stay_archived(cache_dir, &self.teams.stay_archived) {
            Ok(()) => None,
            Err(error) => {
                tracing::warn!(
                    "could not persist Stay archived state in {}: {error}",
                    cache_dir.display()
                );
                Some(error.to_string())
            }
        }
    }

    fn spawn<F>(&self, fut: F)
    where
        F: Future<Output = anyhow::Result<AppMessage>> + Send + 'static,
    {
        let tx = self.tx.clone();
        tokio::spawn(async move {
            let msg = match fut.await {
                Ok(m) => m,
                Err(e) => {
                    // Also record it: the status line is transient, the log isn't.
                    tracing::warn!("background task failed: {e:#}");
                    AppMessage::Error(format!("{e:#}"))
                }
            };
            let _ = tx.send(msg).await;
        });
    }

    fn load_whoami(&self) {
        let s = self.session.clone();
        self.spawn(async move { Ok(AppMessage::Whoami(s.whoami().await?)) });
    }

    fn load_presence(&self) {
        let s = self.session.clone();
        self.spawn(async move {
            Ok(AppMessage::Presence {
                presence: people::my_presence(&s.graph).await?,
                requested: None,
            })
        });
    }

    fn refresh_presence_for_ntfy(&self) {
        let s = self.session.clone();
        self.spawn(async move {
            Ok(AppMessage::PresenceRefresh(
                people::my_presence(&s.graph).await?,
            ))
        });
    }

    fn apply_presence_result(&mut self, presence: Presence, requested: Option<String>) {
        let effective = presence.availability.clone().unwrap_or_default();
        self.status = match requested {
            Some(req) if effective.eq_ignore_ascii_case(&req) => {
                format!("presence set to {req}")
            }
            // e.g. requested Available, Teams reports AvailableIdle.
            Some(req) if effective.to_lowercase().starts_with(&req.to_lowercase()) => {
                format!("presence set to {req} (Teams reports {effective})")
            }
            // Chose "Appear offline": Offline is the expected outcome.
            Some(req) if effective.eq_ignore_ascii_case("Offline") => {
                format!("presence set to {req} (you appear Offline)")
            }
            // Graph settled on something else — usually another client or a
            // higher-priority meeting/call/preferred state.
            Some(req) => format!(
                "presence set to {req}; Teams currently shows '{effective}' (a higher-priority state can override this app session)"
            ),
            None => format!("presence: {effective}"),
        };
        self.my_presence = Some(presence);
    }

    fn refresh_ntfy_work_plan(&self) {
        if !self.session.config.ntfy.needs_work_plan() {
            return;
        }

        let s = self.session.clone();
        let tx = self.tx.clone();
        tokio::spawn(async move {
            let working = match work_plan::working_now(&s.graph, chrono::Utc::now()).await {
                Ok(working) => working,
                Err(error) => {
                    tracing::warn!("ntfy work-plan refresh failed: {error:#}");
                    false
                }
            };
            tracing::debug!("ntfy M365 work-plan active={working}");
            let _ = tx.send(AppMessage::NtfyWorkPlan(working)).await;
        });
    }

    fn refresh_diagnostics(&mut self) {
        self.diagnostics.loading = true;
        let s = self.session.clone();
        let tx = self.tx.clone();

        tokio::spawn(async move {
            let token = match s.auth.access_token().await {
                Ok(_) => s
                    .auth
                    .token_info()
                    .await
                    .ok_or_else(|| "token metadata unavailable".to_string()),
                Err(error) => Err(m365_core::util::graph_error_summary(&format!("{error:#}"))),
            };

            let work_plan = work_plan::diagnostics(&s.graph, chrono::Utc::now())
                .await
                .map_err(|error| m365_core::util::graph_error_summary(&format!("{error:#}")));

            let _ = tx
                .send(AppMessage::DiagnosticsLoaded(DiagnosticsRemote {
                    generated_at: chrono::Utc::now(),
                    token,
                    work_plan,
                }))
                .await;
        });
    }

    fn open_diagnostics(&mut self) {
        self.diagnostics_scroll = 0;
        self.overlay = Some(Overlay::Diagnostics);
        self.refresh_diagnostics();
    }

    fn open_contact_diagnostics(&mut self) {
        if self.screen != Screen::Teams || self.teams.mode != TeamsMode::Chats {
            self.status = "F7 contact diagnostics needs the Teams chat list".into();
            return;
        }

        let Some(chat) = self.teams_selected_chat().cloned() else {
            self.status = "select a Teams contact first".into();
            return;
        };
        if !chat
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
        {
            self.status = "F7 contact diagnostics is available for one-to-one chats".into();
            return;
        }

        let chat_id = chat.id.clone();
        let me_id = self.me.as_ref().map(|me| me.id.as_str());
        let peer_member = chat.members.iter().find(|member| {
            me_id
                .map(|id| member.user_id.as_deref() != Some(id))
                .unwrap_or(true)
        });
        let preview_peer = chat
            .last_message_preview
            .as_ref()
            .and_then(|preview| preview.from.as_ref())
            .and_then(|from| from.user.as_ref())
            .filter(|user| {
                me_id
                    .map(|id| user.id.as_deref() != Some(id))
                    .unwrap_or(true)
            });

        let member_resource_id = peer_member
            .and_then(|member| member.id.as_deref())
            .map(str::trim)
            .filter(|value| !value.is_empty());
        let member_id = peer_member
            .and_then(|member| member.user_id.as_deref())
            .map(str::trim)
            .filter(|value| !value.is_empty());
        let preview_id = preview_peer
            .and_then(|user| user.id.as_deref())
            .map(str::trim)
            .filter(|value| !value.is_empty());
        let cached_id = self
            .teams
            .contact_user_ids
            .get(&chat.id)
            .map(String::as_str)
            .map(str::trim)
            .filter(|value| !value.is_empty());

        // Personal-account chats can have no usable Entra GUID and the current
        // preview can be our own Workflows "Post as User" message. Recover the
        // real peer sender id from persistent conversation history when needed.
        let cached_message_peer_id = if cached_id.is_none() {
            self.session
                .config
                .cache_dir
                .as_deref()
                .and_then(|cache_root| {
                    load_teams_conversation_cache(cache_root, &chat.id)
                        .ok()
                        .flatten()
                })
                .and_then(|(messages, _)| {
                    messages
                        .iter()
                        .filter_map(|message| message.from.as_ref()?.user.as_ref())
                        .find(|user| {
                            me_id
                                .map(|id| user.id.as_deref() != Some(id))
                                .unwrap_or(true)
                        })
                        .and_then(|user| user.id.clone())
                })
                .map(|value| value.trim().to_string())
                .filter(|value| !value.is_empty())
        } else {
            None
        };

        let automation_peer_id = cached_id
            .map(str::to_string)
            .or(cached_message_peer_id)
            .or_else(|| member_id.map(str::to_string))
            .or_else(|| preview_id.map(str::to_string));
        let automation_peer_configured = automation_peer_id
            .as_deref()
            .is_some_and(|peer_id| self.teams_automation_peers.contains(peer_id));
        let automation_config_source =
            if self.session.config.teams_automation_peers_from_process_env {
                "process environment".to_string()
            } else if let Some(path) = self.session.config.dotenv_path.as_deref() {
                path.display().to_string()
            } else {
                "no .env loaded".to_string()
            };

        let member_guid = member_id.is_some_and(chats::looks_like_user_guid);
        let preview_guid = preview_id.is_some_and(chats::looks_like_user_guid);
        let cached_guid = cached_id.is_some_and(chats::looks_like_user_guid);

        let member_id_shape = crate::diagnostics::identifier_shape(member_resource_id).to_string();
        let member_user_id_shape = crate::diagnostics::identifier_shape(member_id).to_string();
        let preview_id_shape = crate::diagnostics::identifier_shape(preview_id).to_string();
        let cached_id_shape = crate::diagnostics::identifier_shape(cached_id).to_string();

        let mri_candidates = [
            ("member resource ID", member_resource_id),
            ("member userId", member_id),
            ("message sender ID", preview_id),
            ("cached contact ID", cached_id),
        ];
        let (mri_candidate, mri_candidate_source) = mri_candidates
            .into_iter()
            .find(|(_, value)| crate::diagnostics::is_presence_mri_candidate(*value))
            .map(|(source, value)| (value.map(str::to_string), source))
            .unwrap_or((None, "none"));

        let (probe_user_id, probe_source) = if cached_guid {
            (cached_id.map(str::to_string), "cached contact ID")
        } else if member_guid {
            (member_id.map(str::to_string), "chat member")
        } else if preview_guid {
            (preview_id.map(str::to_string), "message sender")
        } else {
            (None, "none")
        };

        let lookup_address = peer_member
            .and_then(|member| member.email.as_deref())
            .map(str::trim)
            .filter(|value| !value.is_empty())
            .map(str::to_string);

        let raw_member_type = peer_member.and_then(|member| member.odata_type.as_deref());
        let member_type = raw_member_type
            .map(|kind| {
                kind.rsplit('.')
                    .next()
                    .unwrap_or(kind)
                    .trim_start_matches('#')
                    .to_string()
            })
            .unwrap_or_else(|| "unknown".to_string());

        let personal = member_type.eq_ignore_ascii_case("microsoftAccountUserConversationMember");
        let skype = member_type.eq_ignore_ascii_case("skypeUserConversationMember")
            || member_type.eq_ignore_ascii_case("skypeForBusinessUserConversationMember");
        let aad = member_type.eq_ignore_ascii_case("aadUserConversationMember");
        let external = self.teams.external_chats.contains(&chat.id);

        let peer_tenant = peer_member
            .and_then(|member| member.tenant_id.as_deref())
            .or_else(|| preview_peer.and_then(|user| user.tenant_id.as_deref()))
            .map(str::trim)
            .filter(|value| !value.is_empty());
        let tenant_relation = match (self.teams.own_tenant_id.as_deref(), peer_tenant) {
            (Some(own), Some(peer)) if own.eq_ignore_ascii_case(peer) => "same tenant",
            (Some(_), Some(_)) => "external",
            _ if external => "external",
            _ => "unknown",
        }
        .to_string();

        let account_type = if personal {
            "Microsoft personal account"
        } else if skype {
            "Skype identity"
        } else if aad && external {
            "External / federated Entra user"
        } else if aad {
            "Organization user"
        } else if external {
            "External / federated identity"
        } else {
            "Unknown"
        }
        .to_string();

        let presence_supported = !personal && !skype && probe_user_id.is_some();
        let cross_tenant_candidate = external && presence_supported;

        let topic_available = chat
            .topic
            .as_deref()
            .map(str::trim)
            .is_some_and(|value| !value.is_empty());
        let member_display_name_available = peer_member
            .and_then(|member| member.display_name.as_deref())
            .map(str::trim)
            .is_some_and(|value| !value.is_empty());
        let member_email_available = peer_member
            .and_then(|member| member.email.as_deref())
            .map(str::trim)
            .is_some_and(|value| !value.is_empty());
        let preview_present = chat.last_message_preview.is_some();
        let preview_sender_available = chat
            .last_message_preview
            .as_ref()
            .and_then(|preview| preview.from.as_ref())
            .and_then(|from| from.user.as_ref())
            .is_some();
        let preview_sender_is_peer = preview_peer.is_some();
        let preview_display_name_available = preview_peer
            .and_then(|user| user.display_name.as_deref())
            .map(str::trim)
            .is_some_and(|value| !value.is_empty());
        let cached_name_available = self
            .teams
            .contact_names
            .get(&chat.id)
            .map(String::as_str)
            .map(str::trim)
            .is_some_and(|value| !value.is_empty());

        let chat_label_source = if topic_available {
            "topic"
        } else if member_display_name_available {
            "member displayName"
        } else if preview_display_name_available {
            "preview sender displayName"
        } else if member_email_available {
            "member email"
        } else {
            "chatType fallback"
        }
        .to_string();

        let list_label_source = if cached_name_available {
            "cached / learned contact name".to_string()
        } else {
            chat_label_source.clone()
        };

        self.contact_diagnostics = ContactDiagnosticsState {
            loading: true,
            member_type,
            account_type,
            member_guid,
            preview_guid,
            cached_guid,
            member_id_shape,
            member_user_id_shape,
            preview_id_shape,
            cached_id_shape,
            mri_candidate_source: mri_candidate_source.to_string(),
            lookup_address_available: lookup_address.is_some(),
            probe_source: probe_source.to_string(),
            tenant_relation,
            cross_tenant_candidate,
            presence_supported,
            current_member_count: chat.members.len(),
            topic_available,
            member_display_name_available,
            member_email_available,
            preview_present,
            preview_sender_available,
            preview_sender_is_peer,
            preview_display_name_available,
            cached_name_available,
            chat_label_source,
            list_label_source,
            automation_peer_id,
            automation_peer_configured,
            automation_config_source,
            remote: None,
        };
        self.contact_diagnostics_chat_id = Some(chat_id);
        self.contact_diagnostics_user_id = probe_user_id;
        self.contact_diagnostics_mri = mri_candidate;
        self.contact_diagnostics_lookup_address = lookup_address;
        self.contact_diagnostics_scroll = 0;
        self.overlay = Some(Overlay::ContactDiagnostics);
        self.refresh_contact_diagnostics();
    }

    fn refresh_contact_diagnostics(&mut self) {
        self.contact_diagnostics.loading = true;

        let s = self.session.clone();
        let tx = self.tx.clone();
        let chat_id = self.contact_diagnostics_chat_id.clone();
        let me_id = self.me.as_ref().map(|me| me.id.clone());
        let user_id = self.contact_diagnostics_user_id.clone();
        let lookup_address = self.contact_diagnostics_lookup_address.clone();
        let mri = self.contact_diagnostics_mri.clone();
        let personal_account = self
            .contact_diagnostics
            .member_type
            .eq_ignore_ascii_case("microsoftAccountUserConversationMember");
        let enabled = self.session.config.presence_read;
        let supported = self.contact_diagnostics.presence_supported;

        tokio::spawn(async move {
            let (expanded_name, members_name, messages_name) =
                if let Some(chat_id) = chat_id.as_deref() {
                    let expanded_name = match s
                        .graph
                        .get_json::<Chat>(&format!(
                            "me/chats/{chat_id}?$expand=members,lastMessagePreview"
                        ))
                        .await
                    {
                        Ok(chat) => Ok(crate::diagnostics::contact_chat_name_probe(
                            &chat,
                            me_id.as_deref(),
                        )),
                        Err(error) => Err(crate::diagnostics::safe_graph_error(&error)),
                    };

                    let members_name = match chats::list_members(&s.graph, chat_id).await {
                        Ok(members) => Ok(crate::diagnostics::contact_members_name_probe(
                            &members,
                            me_id.as_deref(),
                        )),
                        Err(error) => Err(crate::diagnostics::safe_graph_error(&error)),
                    };

                    let messages_name = match chats::list_messages(&s.graph, chat_id, 25).await {
                        Ok((messages, _)) => Ok(crate::diagnostics::contact_messages_name_probe(
                            &messages,
                            me_id.as_deref(),
                        )),
                        Err(error) => Err(crate::diagnostics::safe_graph_error(&error)),
                    };

                    (expanded_name, members_name, messages_name)
                } else {
                    let error = "selected chat id unavailable".to_string();
                    (Err(error.clone()), Err(error.clone()), Err(error))
                };

            let presence_read_all = match s.auth.access_token().await {
                Ok(_) => s
                    .auth
                    .token_info()
                    .await
                    .map(|token| {
                        token
                            .scopes
                            .iter()
                            .any(|scope| scope.eq_ignore_ascii_case("Presence.Read.All"))
                    })
                    .ok_or_else(|| "token metadata unavailable".to_string()),
                Err(error) => Err(crate::diagnostics::safe_graph_error(&error)),
            };

            let (batch, direct) = if !enabled {
                (
                    PresenceProbe::Skipped("disabled by configuration".into()),
                    PresenceProbe::Skipped("disabled by configuration".into()),
                )
            } else if !supported {
                (
                    PresenceProbe::Skipped("unsupported identity / no GUID".into()),
                    PresenceProbe::Skipped("unsupported identity / no GUID".into()),
                )
            } else if let Some(user_id) = user_id {
                let ids = vec![user_id.clone()];
                match people::presences(&s.graph, &ids).await {
                    Ok(items) => {
                        if let Some(presence) = items.into_iter().find(|presence| {
                            presence
                                .id
                                .as_deref()
                                .is_some_and(|id| id.eq_ignore_ascii_case(&user_id))
                        }) {
                            (
                                crate::diagnostics::presence_probe(&presence),
                                PresenceProbe::Skipped("not needed".into()),
                            )
                        } else {
                            let direct = match people::presence(&s.graph, &user_id).await {
                                Ok(presence) => crate::diagnostics::presence_probe(&presence),
                                Err(error) => PresenceProbe::Error(
                                    crate::diagnostics::safe_graph_error(&error),
                                ),
                            };
                            (PresenceProbe::Omitted, direct)
                        }
                    }
                    Err(error) => {
                        let batch =
                            PresenceProbe::Error(crate::diagnostics::safe_graph_error(&error));
                        let direct = match people::presence(&s.graph, &user_id).await {
                            Ok(presence) => crate::diagnostics::presence_probe(&presence),
                            Err(error) => {
                                PresenceProbe::Error(crate::diagnostics::safe_graph_error(&error))
                            }
                        };
                        (batch, direct)
                    }
                }
            } else {
                (
                    PresenceProbe::Skipped("no Entra GUID".into()),
                    PresenceProbe::Skipped("no Entra GUID".into()),
                )
            };

            let teams = if enabled {
                teams_presence::diagnose(
                    &s.auth,
                    lookup_address.as_deref(),
                    mri.as_deref(),
                    personal_account,
                )
                .await
            } else {
                m365_core::teams_presence::Diagnostics {
                    lookup_address_available: lookup_address.is_some(),
                    resource_token: m365_core::teams_presence::DiagnosticStep::Skipped(
                        "contact presence disabled".into(),
                    ),
                    authz: m365_core::teams_presence::DiagnosticStep::Skipped(
                        "contact presence disabled".into(),
                    ),
                    middle_tier_lookup: m365_core::teams_presence::DiagnosticStep::Skipped(
                        "contact presence disabled".into(),
                    ),
                    resolved_mri_shape: "none".into(),
                    ups_presence: m365_core::teams_presence::DiagnosticStep::Skipped(
                        "contact presence disabled".into(),
                    ),
                }
            };

            let _ = tx
                .send(AppMessage::ContactDiagnosticsLoaded(
                    ContactDiagnosticsRemote {
                        generated_at: chrono::Utc::now(),
                        presence_read_all,
                        batch,
                        direct,
                        teams,
                        expanded_name,
                        members_name,
                        messages_name,
                    },
                ))
                .await;
        });
    }

    fn start_primary_presence(&mut self) {
        if !self.session.config.presence_primary || !self.session.config.can_write_presence() {
            return;
        }

        let timeout = self.presence_idle_timeout();
        let snapshot = self.presence_activity.sample(timeout);
        let now = std::time::Instant::now();
        let last_activity = snapshot
            .idle_for
            .and_then(|idle| now.checked_sub(idle))
            .unwrap_or(now);
        let locked = snapshot.state == crate::activity::ActivityState::Locked;
        let base = if snapshot.state == crate::activity::ActivityState::Idle && !timeout.is_zero() {
            AWAY
        } else {
            AVAILABLE
        };

        self.presence_state
            .start_automatic(base, last_activity, locked);
        self.pump_presence_session_write();
    }

    fn presence_idle_timeout(&self) -> std::time::Duration {
        std::time::Duration::from_secs(
            self.session
                .config
                .presence_available_timeout_min
                .saturating_mul(60),
        )
    }

    fn pump_presence_session_write(&mut self) {
        let force_sync = self.presence_preference_pending.is_some();
        let now = std::time::Instant::now();
        let desired = self.presence_state.diagnostics(now).desired;
        let renew_after = presence_renew_after(desired);
        let Some(write) = self.presence_state.next_write(now, renew_after, force_sync) else {
            return;
        };

        let preference = self.presence_preference_pending.take();
        let requested = match preference {
            Some(PresencePreferenceAction::Set { label, .. }) => Some(label.to_string()),
            Some(PresencePreferenceAction::Clear) => None,
            None if write.reason == WriteReason::Sync => match write.target {
                SessionTarget::Present((availability, _)) => Some(availability.to_string()),
                SessionTarget::Absent => None,
            },
            None => None,
        };

        self.spawn_presence_session_write(write, preference, requested);
    }

    fn spawn_presence_session_write(
        &self,
        write: PresenceWrite,
        preference: Option<PresencePreferenceAction>,
        requested: Option<String>,
    ) {
        let s = self.session.clone();
        let tx = self.tx.clone();
        let client_id = self.session.config.client_id.clone();

        tokio::spawn(async move {
            if let Some(action) = preference {
                let result = match action {
                    PresencePreferenceAction::Set {
                        availability,
                        activity,
                        ..
                    } => people::set_preferred_presence(&s.graph, availability, activity).await,
                    PresencePreferenceAction::Clear => {
                        people::clear_preferred_presence(&s.graph).await
                    }
                };

                if let Err(error) = result {
                    let _ = tx
                        .send(AppMessage::PresenceSessionWriteFinished {
                            write,
                            requested,
                            presence: None,
                            success: false,
                            error: Some(format!("{error:#}")),
                            retry_preference: Some(action),
                        })
                        .await;
                    return;
                }
            }

            let result = match write.target {
                SessionTarget::Present((availability, activity)) => {
                    people::set_session_presence(
                        &s.graph,
                        &client_id,
                        availability,
                        activity,
                        PRESENCE_SESSION_LEASE,
                    )
                    .await
                }
                SessionTarget::Absent => people::clear_session_presence(&s.graph, &client_id).await,
            };

            if let Err(error) = result {
                let _ = tx
                    .send(AppMessage::PresenceSessionWriteFinished {
                        write,
                        requested,
                        presence: None,
                        success: false,
                        error: Some(format!("{error:#}")),
                        retry_preference: None,
                    })
                    .await;
                return;
            }

            // Lease renewal is intentionally silent. State transitions still
            // refresh Graph's effective presence for the tab bar/status message.
            let (presence, error) = if write.reason == WriteReason::Renew && preference.is_none() {
                (None, None)
            } else {
                match people::my_presence(&s.graph).await {
                    Ok(presence) => (Some(presence), None),
                    Err(error) => (None, Some(format!("presence readback failed: {error:#}"))),
                }
            };

            let _ = tx
                .send(AppMessage::PresenceSessionWriteFinished {
                    write,
                    requested,
                    presence,
                    success: true,
                    error,
                    retry_preference: None,
                })
                .await;
        });
    }

    fn note_presence_activity(&mut self) {
        let now = std::time::Instant::now();
        self.presence_state.on_activity(now);
        self.pump_presence_session_write();
    }

    fn refresh_presence_activity(&mut self) {
        if !self.session.config.presence_primary || !self.presence_state.is_automatic() {
            return;
        }

        let timeout = self.presence_idle_timeout();
        let snapshot = self.presence_activity.sample(timeout);
        let now = std::time::Instant::now();

        match snapshot.state {
            crate::activity::ActivityState::Active => {
                let activity_at = snapshot
                    .idle_for
                    .and_then(|idle| now.checked_sub(idle))
                    .unwrap_or(now);
                self.presence_state.on_activity(activity_at);
            }
            crate::activity::ActivityState::Locked => self.presence_state.on_locked(),
            crate::activity::ActivityState::Idle => {
                // Timeout zero disables idle->Away, but lock detection remains active.
                self.presence_state.on_idle(!timeout.is_zero());
            }
            crate::activity::ActivityState::Unknown => {
                // Backend unavailable: retain the app-local timeout fallback.
                self.presence_state.on_unknown(now, timeout);
            }
        }

        self.pump_presence_session_write();
    }

    pub fn presence_activity_diagnostics(&self) -> crate::activity::ActivityDiagnostics {
        self.presence_activity.diagnostics()
    }

    pub fn presence_state_diagnostics(&self) -> PresenceStateDiagnostics {
        self.presence_state.diagnostics(std::time::Instant::now())
    }

    pub fn presence_lock_restore_enabled(&self) -> bool {
        self.session.config.presence_lock_restore
    }

    pub fn presence_session_active_on_exit(&self) -> bool {
        self.presence_state.may_have_remote_session()
    }

    fn open_presence_status_message_editor(&mut self) {
        let current = self
            .my_presence
            .as_ref()
            .and_then(Presence::status_message_text)
            .unwrap_or_default();
        self.overlay = Some(Overlay::PresenceMessage(TextInput::from(current)));
    }

    fn set_presence_status_message(&mut self, message: String) {
        let message = message.trim().to_string();
        if message.is_empty() {
            self.clear_presence_status_message();
            return;
        }

        let Some(user_id) = self.me.as_ref().map(|user| user.id.clone()) else {
            self.status = "cannot set status message: signed-in user id is not loaded".into();
            return;
        };

        let s = self.session.clone();
        self.status = "setting status message…".into();
        self.spawn(async move {
            people::set_status_message(&s.graph, &user_id, &message).await?;
            Ok(AppMessage::PresenceStatusMessageUpdated {
                presence: people::my_presence(&s.graph).await?,
                cleared: false,
            })
        });
    }

    fn clear_presence_status_message(&mut self) {
        let Some(user_id) = self.me.as_ref().map(|user| user.id.clone()) else {
            self.status = "cannot clear status message: signed-in user id is not loaded".into();
            return;
        };

        let s = self.session.clone();
        self.status = "clearing status message…".into();
        self.spawn(async move {
            people::clear_status_message(&s.graph, &user_id).await?;
            Ok(AppMessage::PresenceStatusMessageUpdated {
                presence: people::my_presence(&s.graph).await?,
                cleared: true,
            })
        });
    }

    fn set_presence(&mut self, opt: &'static PresenceOption) {
        if self.session.config.presence_primary
            && opt.preferred == AVAILABLE
            && opt.session == Some(AVAILABLE)
        {
            tracing::debug!(
                "Available selected with primary presence; clearing preferred override and resuming automatic presence"
            );
            self.status = "returning to automatic Available/Away presence…".into();
            self.presence_state
                .clear_manual(true, std::time::Instant::now());
            self.presence_preference_pending = Some(PresencePreferenceAction::Clear);
            self.pump_presence_session_write();
            return;
        }

        self.presence_state.set_manual(opt.session);
        let (availability, activity) = opt.preferred;
        self.presence_preference_pending = Some(PresencePreferenceAction::Set {
            availability,
            activity,
            label: opt.label,
        });
        self.pump_presence_session_write();
    }

    fn renew_presence_session(&mut self) {
        self.pump_presence_session_write();
    }

    fn clear_presence(&mut self) {
        let primary =
            self.session.config.presence_primary && self.session.config.can_write_presence();
        self.presence_state
            .clear_manual(primary, std::time::Instant::now());
        self.presence_preference_pending = Some(PresencePreferenceAction::Clear);
        self.pump_presence_session_write();
    }

    fn load_folders(&self) {
        let s = self.session.clone();
        self.spawn(async move { Ok(AppMessage::Folders(mail::list_folders(&s.graph).await?)) });
    }

    fn load_messages(&self, folder_id: String, mode: ListUpdate) {
        let s = self.session.clone();
        self.spawn(async move {
            let (items, next) = mail::list_messages(&s.graph, &folder_id, PAGE_SIZE).await?;
            Ok(AppMessage::Messages { items, next, mode })
        });
    }

    fn load_more_messages(&mut self) {
        if self.outlook.loading_more {
            return;
        }
        let Some(next_link) = self.outlook.messages_next.clone() else {
            return;
        };
        self.outlook.loading_more = true;
        self.status = "loading more…".into();
        let s = self.session.clone();
        self.spawn(async move {
            let (items, next) = mail::list_messages_more(&s.graph, &next_link).await?;
            Ok(AppMessage::Messages {
                items,
                next,
                mode: ListUpdate::Append,
            })
        });
    }

    fn peek_inbox(&self) {
        if !self.notification_events_enabled() {
            return;
        }
        let s = self.session.clone();
        self.spawn(async move {
            let (items, _) = mail::list_messages(&s.graph, "inbox", 15).await?;
            Ok(AppMessage::InboxPeek(items))
        });
    }

    fn load_attachments(&self, message_id: String) {
        let s = self.session.clone();
        self.spawn(async move {
            let items = mail::list_attachments(&s.graph, &message_id).await?;
            Ok(AppMessage::Attachments { message_id, items })
        });
    }

    fn download_attachment(&mut self, index: usize) {
        let (Some(msg), Some(att)) = (
            self.outlook.reading.as_ref(),
            self.outlook.reading_attachments.get(index),
        ) else {
            return;
        };
        if !att.is_file() {
            self.status = "only file attachments can be saved".into();
            return;
        }
        let message_id = msg.id.clone();
        let attachment_id = att.id.clone();
        let name = att.display_name();
        self.status = format!("downloading {name}…");
        let s = self.session.clone();
        self.spawn(async move {
            let bytes = mail::download_attachment(&s.graph, &message_id, &attachment_id).await?;
            Ok(AppMessage::Downloaded { name, bytes })
        });
    }

    fn load_body(&self, id: String) {
        let s = self.session.clone();
        self.spawn(async move {
            Ok(AppMessage::MessageBody(
                mail::get_message(&s.graph, &id).await?,
            ))
        });
    }

    fn cancel_read_timer(&mut self) {
        self.read_timer_generation = self.read_timer_generation.wrapping_add(1);
    }

    fn schedule_read_timer(&mut self, id: String) {
        self.cancel_read_timer();
        let generation = self.read_timer_generation;
        let timeout = self.session.config.read_msg_timeout;
        self.spawn(async move {
            if timeout > 0 {
                tokio::time::sleep(std::time::Duration::from_secs(timeout)).await;
            }
            Ok(AppMessage::AutoMarkReadDue { id, generation })
        });
    }

    fn schedule_current_read_timer(&mut self) {
        if self.screen != Screen::Outlook || self.outlook_focus != OutlookFocus::Reading {
            return;
        }

        let Some(current_id) = self.current_mail().map(|message| message.id.clone()) else {
            return;
        };

        let id = match self.outlook.reading.as_ref() {
            Some(message) if message.id == current_id && !message.is_read.unwrap_or(false) => {
                message.id.clone()
            }
            _ => return,
        };
        self.schedule_read_timer(id);
    }

    fn set_mail_read(&mut self, id: &str, read: bool) {
        let in_current_folder = self.outlook.messages.iter().any(|m| m.id == id);
        let previous = self
            .outlook
            .messages
            .iter()
            .chain(self.outlook.reading_thread.iter())
            .find(|m| m.id == id)
            .or_else(|| self.outlook.reading.as_ref().filter(|m| m.id == id))
            .map(|m| m.is_read.unwrap_or(false));
        let Some(was_read) = previous else {
            return;
        };

        for message in self
            .outlook
            .messages
            .iter_mut()
            .chain(self.outlook.reading_thread.iter_mut())
        {
            if message.id == id {
                message.is_read = Some(read);
            }
        }
        if let Some(message) = self.outlook.reading.as_mut().filter(|m| m.id == id) {
            message.is_read = Some(read);
        }
        if let Some(message) = self.outlook.message_cache.get_mut(id) {
            message.is_read = Some(read);
        }
        for (messages, _) in self.outlook.thread_cache.values_mut() {
            if let Some(message) = messages.iter_mut().find(|message| message.id == id) {
                message.is_read = Some(read);
            }
        }

        if in_current_folder {
            if let Some(folder) = self.outlook.folders.get_mut(self.outlook.folder_sel) {
                if let Some(count) = folder.unread_item_count.as_mut() {
                    if read && !was_read {
                        *count = count.saturating_sub(1);
                    } else if !read && was_read {
                        *count += 1;
                    }
                }
            }
        }
        let id = id.to_string();
        let s = self.session.clone();
        self.status = if read {
            "marking as read…".into()
        } else {
            "marking as unread…".into()
        };
        self.spawn(async move {
            mail::mark_read(&s.graph, &id, read).await?;
            Ok(AppMessage::Status(if read {
                "marked as read".into()
            } else {
                "marked as unread".into()
            }))
        });
    }

    fn refresh_calendar_reminders_if_due(&mut self) {
        if !self.session.config.calendar_notify.enabled() {
            return;
        }
        if self
            .calendar_reminder_last_refresh
            .is_some_and(|last| last.elapsed() < std::time::Duration::from_secs(60))
        {
            return;
        }
        self.refresh_calendar_reminders();
    }

    fn refresh_calendar_reminders(&mut self) {
        if !self.session.config.calendar_notify.enabled() {
            return;
        }

        self.calendar_reminder_last_refresh = Some(std::time::Instant::now());

        // Keep this feed independent of the visible Calendar range. A user can
        // stay in Outlook or Teams all day and still receive event reminders.
        let now = chrono::Utc::now();
        let start =
            (now - chrono::Duration::minutes(1)).to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
        let end = (now + chrono::Duration::minutes(20))
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        let s = self.session.clone();
        let tx = self.tx.clone();
        tokio::spawn(async move {
            match calendar::calendar_view(&s.graph, &start, &end).await {
                Ok(events) => {
                    let _ = tx.send(AppMessage::CalendarReminders(events)).await;
                }
                Err(error) => {
                    // Reminder refresh is auxiliary and must never replace the
                    // user's current status line with a recurring network error.
                    tracing::warn!("calendar reminder refresh failed: {error:#}");
                }
            }
        });
    }

    fn check_calendar_reminders(&mut self) {
        let mode = self.session.config.calendar_notify;
        if !mode.enabled() {
            return;
        }

        let now = chrono::Local::now();
        let mut internal = Vec::new();
        let mut external = Vec::new();

        // Clone the small near-term feed so notification bookkeeping can mutate
        // App state without fighting an immutable borrow of the event vector.
        for event in self.calendar_reminder_events.clone() {
            let Some((minutes, start)) = calendar_event_reminder(&event, &now) else {
                continue;
            };

            let key = format!("{}:{}:{minutes}", event.id, start.timestamp());
            if !self.calendar_reminder_fired.insert(key) {
                continue;
            }

            let subject = event
                .subject
                .as_deref()
                .map(str::trim)
                .filter(|subject| !subject.is_empty())
                .unwrap_or("(no subject)")
                .split_whitespace()
                .collect::<Vec<_>>()
                .join(" ");
            let at = start.format("%H:%M").to_string();

            if mode.internal() {
                internal.push(format!("{minutes} min: {subject} ({at})"));
            }

            if mode.external() {
                let body = event
                    .location
                    .as_ref()
                    .and_then(|location| location.display_name.as_deref())
                    .map(str::trim)
                    .filter(|location| !location.is_empty())
                    .map(|location| format!("{subject} · {at} · {location}"))
                    .unwrap_or_else(|| format!("{subject} · {at}"));
                external.push((format!("📅 Calendar — {minutes} min"), body));
            }
        }

        for (title, body) in external {
            self.send_notification("Calendar", title, body, NotificationPriority::Normal);
        }

        if let Some(first) = internal.first() {
            self.status = if internal.len() == 1 {
                format!("📅 {first}")
            } else {
                format!("📅 {first} (+{} more)", internal.len() - 1)
            };
        }
    }

    fn load_calendar(&self) {
        let s = self.session.clone();
        let (start, end) = match self.calendar.view {
            CalendarView::Agenda => {
                let start = chrono::Utc::now();
                let end = start + chrono::Duration::days(self.calendar.days);
                (start, end)
            }
            CalendarView::Month => {
                let first = calendar_month_first(self.calendar.month_offset);
                let leading = chrono::Datelike::weekday(&first).num_days_from_monday() as i64;
                let grid_start = first - chrono::Duration::days(leading + 1);
                let grid_end = calendar_month_first(self.calendar.month_offset + 4)
                    + chrono::Duration::days(8);

                let start = chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(
                    grid_start
                        .and_hms_opt(0, 0, 0)
                        .expect("valid calendar midnight"),
                    chrono::Utc,
                );
                let end = chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(
                    grid_end
                        .and_hms_opt(0, 0, 0)
                        .expect("valid calendar midnight"),
                    chrono::Utc,
                );
                (start, end)
            }
        };
        let start = start.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
        let end = end.to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
        self.spawn(async move {
            Ok(AppMessage::Calendar(
                calendar::calendar_view(&s.graph, &start, &end).await?,
            ))
        });
    }

    fn load_chats(&self) {
        let s = self.session.clone();
        self.spawn(async move { Ok(AppMessage::Chats(chats::list_chats(&s.graph, 40).await?)) });
    }

    fn warm_teams_conversation_caches(&mut self, chats_list: &[Chat]) {
        if self.chat_cache_warmup_started || !self.session.config.teams_cache_warmup {
            return;
        }
        let Some(cache_root) = self.session.config.cache_dir.clone() else {
            return;
        };

        let selected_id = self.teams_selected_chat().map(|chat| chat.id.clone());
        let last_chat_id = self.teams.last_chat_id.clone();

        let mut ordered: Vec<(String, Option<String>)> = Vec::new();
        let mut seen = std::collections::HashSet::new();

        let mut push_chat = |chat: &Chat| {
            if seen.insert(chat.id.clone()) {
                ordered.push((
                    chat.id.clone(),
                    chat.last_message_preview
                        .as_ref()
                        .and_then(|preview| preview.id.clone()),
                ));
            }
        };

        if let Some(selected_id) = selected_id.as_deref() {
            if let Some(chat) = chats_list.iter().find(|chat| chat.id == selected_id) {
                push_chat(chat);
            }
        }
        if let Some(last_chat_id) = last_chat_id.as_deref() {
            if let Some(chat) = chats_list.iter().find(|chat| chat.id == last_chat_id) {
                push_chat(chat);
            }
        }
        for chat in chats_list {
            push_chat(chat);
        }

        self.chat_cache_warmup_started = true;
        let s = self.session.clone();
        let tx = self.tx.clone();
        let limiter = self.teams_background_limiter.clone();
        let cache_max_bytes = self.session.config.cache_max_mb.saturating_mul(1024 * 1024);

        tokio::spawn(async move {
            for (chat_id, latest_id) in ordered {
                let read_root = cache_root.clone();
                let read_chat_id = chat_id.clone();
                let cached = match tokio::task::spawn_blocking(move || {
                    load_teams_conversation_cache(&read_root, &read_chat_id)
                })
                .await
                {
                    Ok(Ok(value)) => value,
                    Ok(Err(error)) => {
                        tracing::debug!(
                            "Teams conversation cache warm-up read failed for {chat_id}: {error}"
                        );
                        None
                    }
                    Err(error) => {
                        tracing::debug!(
                            "Teams conversation cache warm-up reader failed for {chat_id}: {error}"
                        );
                        None
                    }
                };

                // Feed an existing disk cache back through the normal warm-up
                // message path even when it is already current. Besides restoring
                // the preview, local Teams metadata (including learned contact
                // names for personal/federated 1:1 chats) can then be rebuilt
                // immediately at startup without an unnecessary Graph request.
                if let Some((messages, _)) = cached.as_ref() {
                    let _ = tx
                        .send(AppMessage::ChatCacheWarmed {
                            chat_id: chat_id.clone(),
                            messages: messages.clone(),
                        })
                        .await;
                    tracing::debug!(
                        "Teams conversation cache restored into warm-up state for {chat_id}"
                    );
                }

                let cache_is_current = cached.as_ref().is_some_and(|(messages, _)| match latest_id
                    .as_deref()
                {
                    Some(latest_id) => {
                        messages.first().map(|message| message.id.as_str()) == Some(latest_id)
                    }
                    None => true,
                });
                if cache_is_current {
                    continue;
                }

                let page = {
                    let _permit = limiter.acquire(&s.graph, &chat_id).await;
                    chats::list_messages(&s.graph, &chat_id, PAGE_SIZE).await
                };
                let (messages, _) = match page {
                    Ok(page) => page,
                    Err(error) => {
                        tracing::warn!(
                            "Teams conversation cache warm-up failed for {chat_id}: {error:#}"
                        );
                        // Keep this deliberately gentle on Graph even when
                        // one conversation repeatedly fails.
                        tokio::time::sleep(std::time::Duration::from_millis(100)).await;
                        continue;
                    }
                };

                // Graph returns newest-first; persistent storage expects the
                // app's oldest-first ordering and reverses it for JSON.
                let mut oldest_first = messages.clone();
                oldest_first.reverse();

                let write_root = cache_root.clone();
                let write_chat_id = chat_id.clone();
                let write_messages = oldest_first;
                match tokio::task::spawn_blocking(move || {
                    store_teams_conversation_cache(
                        &write_root,
                        &write_chat_id,
                        &write_messages,
                        cache_max_bytes,
                    )
                })
                .await
                {
                    Ok(Ok(())) => {
                        tracing::debug!(
                            "Teams conversation cache automatically warmed for {chat_id}"
                        );
                    }
                    Ok(Err(error)) => {
                        tracing::warn!(
                            "Teams conversation cache warm-up write failed for {chat_id}: {error}"
                        );
                    }
                    Err(error) => {
                        tracing::warn!(
                            "Teams conversation cache warm-up writer failed for {chat_id}: {error}"
                        );
                    }
                }

                let _ = tx
                    .send(AppMessage::ChatCacheWarmed { chat_id, messages })
                    .await;

                // Sequential fetching already limits concurrency; this tiny gap
                // additionally avoids hammering Graph on a completely cold cache.
                tokio::time::sleep(std::time::Duration::from_millis(75)).await;
            }
        });
    }

    fn mark_teams_chat_read(&self, chat_id: String) {
        let s = self.session.clone();
        let known_user_id = self.me.as_ref().map(|me| me.id.clone());
        let tenant_id = self.session.config.tenant_id.clone();
        tokio::spawn(async move {
            let result = async {
                let user_id = match known_user_id {
                    Some(id) => id,
                    None => people::me(&s.graph).await?.id,
                };
                chats::mark_read(&s.graph, &chat_id, &user_id, &tenant_id).await
            }
            .await;

            match result {
                Ok(()) => tracing::debug!("marked Teams chat read on server: {chat_id}"),
                Err(e) => tracing::warn!("marking Teams chat read failed for {chat_id}: {e:#}"),
            }
        });
    }

    fn configured_tenant_id(&self) -> Option<&str> {
        let value = self.session.config.tenant_id.trim();
        if value.is_empty()
            || value.eq_ignore_ascii_case("common")
            || value.eq_ignore_ascii_case("organizations")
            || value.eq_ignore_ascii_case("consumers")
        {
            None
        } else {
            Some(value)
        }
    }

    fn remember_contact_from_chat(&mut self, chat: &Chat) -> bool {
        let me_id = self.me.as_ref().map(|me| me.id.as_str());

        // The hideForUser/unhideForUser actions need our tenant id for every
        // chat type, so learn it before restricting the contact-specific work
        // below to one-to-one chats.
        if self.teams.own_tenant_id.is_none() {
            let configured = self.configured_tenant_id().map(str::to_string);
            let from_member = me_id.and_then(|id| {
                chat.members
                    .iter()
                    .find(|member| member.user_id.as_deref() == Some(id))
                    .and_then(|member| member.tenant_id.clone())
            });
            let from_chat = chat
                .tenant_id
                .as_deref()
                .map(str::trim)
                .filter(|value| !value.is_empty())
                .map(str::to_string);
            self.teams.own_tenant_id = configured.or(from_member).or(from_chat);
        }

        if !chat
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
        {
            return false;
        }

        let peer_member = chat.members.iter().find(|member| {
            me_id
                .map(|id| member.user_id.as_deref() != Some(id))
                .unwrap_or(false)
        });

        let preview_peer = chat
            .last_message_preview
            .as_ref()
            .and_then(|preview| preview.from.as_ref())
            .and_then(|from| from.user.as_ref())
            .filter(|user| {
                me_id
                    .map(|id| user.id.as_deref() != Some(id))
                    .unwrap_or(false)
            });

        if let Some(name) = peer_member
            .and_then(|member| member.display_name.as_deref())
            .or_else(|| preview_peer.and_then(|user| user.display_name.as_deref()))
            .or_else(|| peer_member.and_then(|member| member.email.as_deref()))
            .map(str::trim)
            .filter(|value| !value.is_empty())
        {
            self.teams
                .contact_names
                .insert(chat.id.clone(), name.to_string());
        }

        let peer_id = peer_member
            .and_then(|member| member.user_id.as_deref())
            .or_else(|| preview_peer.and_then(|user| user.id.as_deref()))
            .map(str::trim)
            .filter(|value| !value.is_empty());

        let learned_peer_id = peer_id.is_some_and(|id| {
            let changed = self
                .teams
                .contact_user_ids
                .insert(chat.id.clone(), id.to_string())
                .as_deref()
                != Some(id);
            if changed {
                tracing::debug!("Teams 1:1 peer id learned for {}: {}", chat.id, id);
            }
            changed
        });

        let peer_tenant = peer_member
            .and_then(|member| member.tenant_id.as_deref())
            .or_else(|| preview_peer.and_then(|user| user.tenant_id.as_deref()))
            .map(str::trim)
            .filter(|value| !value.is_empty());

        let non_aad_member = peer_member
            .and_then(|member| member.odata_type.as_deref())
            .is_some_and(|kind| {
                !kind.eq_ignore_ascii_case("#microsoft.graph.aadUserConversationMember")
            });
        let opaque_peer_id = peer_id.is_some_and(|id| !chats::looks_like_user_guid(id));

        if non_aad_member || opaque_peer_id {
            self.teams.external_chats.insert(chat.id.clone());
            tracing::debug!(
                "Teams 1:1 chat {} marked external: member_type={:?}, peer_id={:?}",
                chat.id,
                peer_member.and_then(|member| member.odata_type.as_deref()),
                peer_id
            );
        } else if let (Some(own), Some(peer)) = (self.teams.own_tenant_id.as_deref(), peer_tenant) {
            if own.eq_ignore_ascii_case(peer) {
                self.teams.external_chats.remove(&chat.id);
            } else {
                self.teams.external_chats.insert(chat.id.clone());
                tracing::debug!(
                    "Teams 1:1 chat {} marked external: own tenant={}, peer tenant={}",
                    chat.id,
                    own,
                    peer
                );
            }
        }

        learned_peer_id
    }

    fn remember_contacts_from_chats(&mut self, chats: &[Chat]) -> bool {
        let mut learned_peer_id = false;
        for chat in chats {
            learned_peer_id |= self.remember_contact_from_chat(chat);
        }
        learned_peer_id
    }

    fn remember_contact_from_messages(&mut self, chat_id: &str, messages: &[ChatMessage]) -> bool {
        let me_id = self.me.as_ref().map(|me| me.id.as_str());

        // For Microsoft personal-account 1:1 chats the roster userId is the
        // authoritative peer identity. Message sender identities can be aadUser
        // GUIDs (notably Workflows/Post as User) and must never replace that
        // opaque personal-account id in contact_user_ids.
        let personal_roster_peer_id = self
            .teams
            .chats
            .iter()
            .find(|chat| chat.id == chat_id)
            .and_then(|chat| {
                chat.members.iter().find(|member| {
                    me_id
                        .map(|id| member.user_id.as_deref() != Some(id))
                        .unwrap_or(true)
                        && member.odata_type.as_deref().is_some_and(|kind| {
                            kind.ends_with("microsoftAccountUserConversationMember")
                        })
                })
            })
            .and_then(|member| member.user_id.as_deref())
            .map(str::trim)
            .filter(|value| !value.is_empty())
            .map(str::to_string);

        let mut learned_peer_id = false;
        if let Some(peer_id) = personal_roster_peer_id.as_deref() {
            learned_peer_id = self
                .teams
                .contact_user_ids
                .insert(chat_id.to_string(), peer_id.to_string())
                .as_deref()
                != Some(peer_id);
        }

        if self.teams.own_tenant_id.is_none() {
            self.teams.own_tenant_id = self.configured_tenant_id().map(str::to_string);
        }

        if self.teams.own_tenant_id.is_none() {
            if let Some(tenant_id) = messages
                .iter()
                .filter_map(|message| message.from.as_ref()?.user.as_ref())
                .find(|user| {
                    me_id
                        .map(|id| user.id.as_deref() == Some(id))
                        .unwrap_or(false)
                })
                .and_then(|user| user.tenant_id.clone())
            {
                self.teams.own_tenant_id = Some(tenant_id);
            }
        }

        let Some(peer) = messages
            .iter()
            .filter_map(|message| message.from.as_ref()?.user.as_ref())
            .find(|user| {
                me_id
                    .map(|id| user.id.as_deref() != Some(id))
                    .unwrap_or(false)
            })
        else {
            return learned_peer_id;
        };

        if let Some(name) =
            peer_display_name_from_messages(messages, me_id, personal_roster_peer_id.as_deref())
        {
            self.teams
                .contact_names
                .insert(chat_id.to_string(), name.to_string());
        }

        if personal_roster_peer_id.is_none() {
            learned_peer_id |= peer
                .id
                .as_deref()
                .map(str::trim)
                .filter(|value| !value.is_empty())
                .is_some_and(|id| {
                    self.teams
                        .contact_user_ids
                        .insert(chat_id.to_string(), id.to_string())
                        .as_deref()
                        != Some(id)
                });
        }

        let federated = peer
            .user_identity_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("federatedUser"));

        if federated {
            self.teams.external_chats.insert(chat_id.to_string());
        } else if let (Some(own), Some(peer_tenant)) = (
            self.teams.own_tenant_id.as_deref(),
            peer.tenant_id
                .as_deref()
                .map(str::trim)
                .filter(|value| !value.is_empty()),
        ) {
            if own.eq_ignore_ascii_case(peer_tenant) {
                self.teams.external_chats.remove(chat_id);
            } else {
                self.teams.external_chats.insert(chat_id.to_string());
            }
        }

        learned_peer_id
    }

    fn open_selected_contact_profile(&mut self) {
        if self.teams.mode != TeamsMode::Chats || self.teams.focus != TeamsFocus::List {
            return;
        }

        let Some(chat) = self.teams_selected_chat() else {
            self.status = "select a Teams contact first".into();
            return;
        };

        if !chat
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
        {
            self.status = "profile is available for one-to-one chats".into();
            return;
        }

        let me_id = self.me.as_ref().map(|me| me.id.as_str());
        let chat_id = chat.id.clone();
        let user_id = self
            .teams
            .contact_user_ids
            .get(&chat_id)
            .cloned()
            .or_else(|| chat.peer_user_id(me_id).map(str::to_string));
        let display_name = self
            .teams
            .contact_names
            .get(&chat_id)
            .cloned()
            .unwrap_or_else(|| chat.label(me_id));

        let peer_member = user_id
            .as_deref()
            .and_then(|id| {
                chat.members
                    .iter()
                    .find(|member| member.user_id.as_deref() == Some(id))
            })
            .or_else(|| {
                chat.members.iter().find(|member| {
                    me_id
                        .map(|id| member.user_id.as_deref() != Some(id))
                        .unwrap_or(true)
                })
            });

        let fallback_email = peer_member.and_then(|member| member.email.clone());
        let personal_account = peer_member
            .and_then(|member| member.odata_type.as_deref())
            .is_some_and(|kind| kind.ends_with("microsoftAccountUserConversationMember"));
        let external = self.teams.external_chats.contains(&chat_id);
        let account = if personal_account {
            "Personal Microsoft account".to_string()
        } else if external {
            "External / federated account".to_string()
        } else {
            "Organization user".to_string()
        };

        self.contact_profile = Some(ContactProfileState {
            chat_id: chat_id.clone(),
            user_id: user_id.clone(),
            display_name: display_name.clone(),
            fallback_email: fallback_email.clone(),
            account,
            person: None,
            avatar: None,
            loading: true,
        });
        self.contact_profile_scroll = 0;
        self.contact_profile_max_scroll.set(0);
        self.overlay = Some(Overlay::ContactProfile);

        let s = self.session.clone();
        self.spawn(async move {
            let directory_user = user_id
                .as_deref()
                .is_some_and(chats::looks_like_user_guid);
            let can_resolve_profile = directory_user || fallback_email.is_some();

            let person = if directory_user && s.config.directory_profile {
                match user_id.as_deref() {
                    Some(id) => match people::directory_profile_for_contact(&s.graph, id).await {
                        Ok(person) => Some(person),
                        Err(error) => {
                            tracing::debug!(
                                "Teams directory profile lookup failed for {chat_id}: {error:#}"
                            );
                            None
                        }
                    },
                    None => None,
                }
            } else if can_resolve_profile {
                match people::profile_for_contact(
                    &s.graph,
                    user_id.as_deref(),
                    fallback_email.as_deref(),
                    &display_name,
                )
                .await
                {
                    Ok(person) => person,
                    Err(error) => {
                        tracing::debug!(
                            "Teams contact profile lookup failed for {chat_id}: {error:#}"
                        );
                        None
                    }
                }
            } else {
                tracing::debug!(
                    "Teams contact profile lookup skipped for {chat_id}: opaque external id without email"
                );
                None
            };

            let avatar = if !s.config.can_read_profile_photos() {
                tracing::debug!(
                    "Teams contact avatar lookup skipped for {chat_id}: M365_PROFILE_PHOTO is disabled"
                );
                None
            } else if let Some(id) = user_id
                .as_deref()
                .filter(|id| chats::looks_like_user_guid(id))
            {
                match people::profile_photo(&s.graph, id).await {
                    Ok(bytes) if bytes.len() <= MAX_MAIL_IMAGE_BYTES => {
                        match image::load_from_memory(&bytes) {
                            Ok(image) => Some(image),
                            Err(error) => {
                                tracing::debug!(
                                    "Teams contact avatar decode failed for {chat_id}: {error}"
                                );
                                None
                            }
                        }
                    }
                    Ok(bytes) => {
                        tracing::debug!(
                            "Teams contact avatar ignored for {chat_id}: {} bytes exceeds limit",
                            bytes.len()
                        );
                        None
                    }
                    Err(error) => {
                        tracing::debug!(
                            "Teams contact avatar unavailable for {chat_id}: {error:#}"
                        );
                        None
                    }
                }
            } else {
                None
            };

            Ok(AppMessage::ContactProfileLoaded {
                chat_id,
                person,
                avatar,
            })
        });
    }

    fn load_contact_presences(&self, chats: &[Chat]) {
        if !self.session.config.presence_read {
            return;
        }

        // Include both fresh member ids and ids learned from chat messages.
        // Federated member objects sometimes omit fields that the message sender
        // identity still supplies.
        let mut user_ids: Vec<String> = chats
            .iter()
            .filter(|chat| {
                chat.chat_type
                    .as_deref()
                    .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
            })
            .flat_map(|chat| chat.members.iter())
            .filter_map(|member| member.user_id.clone())
            .chain(self.teams.contact_user_ids.values().cloned())
            .filter(|id| chats::looks_like_user_guid(id))
            .collect();
        user_ids.sort();
        user_ids.dedup();

        if user_ids.is_empty() {
            return;
        }

        let s = self.session.clone();
        let tx = self.tx.clone();
        tokio::spawn(async move {
            let requested_ids = user_ids;
            let mut items = match people::presences(&s.graph, &requested_ids).await {
                Ok(items) => items,
                Err(error) => {
                    tracing::debug!(
                        "batch contact presence refresh failed; trying direct lookups: {error:#}"
                    );
                    Vec::new()
                }
            };

            let mut returned: std::collections::HashSet<String> = items
                .iter()
                .filter_map(|presence| presence.id.clone())
                .collect();

            for user_id in requested_ids {
                if returned.contains(&user_id) {
                    continue;
                }

                match people::presence(&s.graph, &user_id).await {
                    Ok(presence) => {
                        tracing::debug!(
                            "direct contact presence lookup succeeded for {user_id}: availability={:?} activity={:?}",
                            presence.availability,
                            presence.activity
                        );
                        if let Some(id) = presence.id.clone() {
                            returned.insert(id);
                        }
                        items.push(presence);
                    }
                    Err(error) => {
                        // This is expected for external users when cross-tenant
                        // presence trust is not mutual.
                        tracing::debug!(
                            "direct contact presence lookup unavailable for {user_id}: {error:#}"
                        );
                    }
                }
            }

            if !items.is_empty() {
                let _ = tx.send(AppMessage::ContactPresences(items)).await;
            }
        });
    }

    fn teams_chat_is_automation_peer(&self, chat_id: &str) -> bool {
        self.teams
            .contact_user_ids
            .get(chat_id)
            .is_some_and(|peer_id| self.teams_automation_peers.contains(peer_id))
    }

    fn teams_preview_is_automation_incoming(
        &self,
        chat_id: &str,
        message_id: &str,
        preview: &m365_core::models::LastMessagePreview,
        my_id: Option<&str>,
    ) -> bool {
        if !self.teams_chat_is_automation_peer(chat_id)
            || self.own_sent_chat_message_ids.contains(message_id)
        {
            return false;
        }
        let from_id = preview
            .from
            .as_ref()
            .and_then(|from| from.user.as_ref())
            .and_then(|user| user.id.as_deref());
        from_id.is_some() && from_id == my_id
    }

    fn teams_message_is_automation_incoming(
        &self,
        chat_id: &str,
        message: &ChatMessage,
        my_id: Option<&str>,
    ) -> bool {
        if message.deleted_date_time.is_some()
            || message.is_system_event()
            || !self.teams_chat_is_automation_peer(chat_id)
            || self.own_sent_chat_message_ids.contains(&message.id)
        {
            return false;
        }
        let from_id = message
            .from
            .as_ref()
            .and_then(|from| from.user.as_ref())
            .and_then(|user| user.id.as_deref());
        from_id.is_some() && from_id == my_id
    }

    fn refresh_chat_unread_counts(&mut self, chats_list: &[Chat]) {
        let Some(me_id) = self.me.as_ref().map(|me| me.id.clone()) else {
            return;
        };

        let preserve_application_unread: std::collections::HashSet<&str> = chats_list
            .iter()
            .filter_map(|chat| {
                if !chat
                    .chat_type
                    .as_deref()
                    .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
                {
                    return None;
                }

                let latest_preview = chat.last_message_preview.as_ref()?;
                let latest_id = latest_preview.id.as_deref()?;
                if self
                    .teams
                    .locally_read_through
                    .get(&chat.id)
                    .is_some_and(|id| id == latest_id)
                {
                    return None;
                }

                if !teams_sender_is_incoming(latest_preview.from.as_ref(), Some(me_id.as_str())) {
                    return None;
                }

                latest_preview
                    .from
                    .as_ref()
                    .is_some_and(|from| from.application.is_some())
                    .then_some(chat.id.as_str())
            })
            .collect();

        let preserve_automation_unread: std::collections::HashSet<&str> = chats_list
            .iter()
            .filter_map(|chat| {
                if !chat
                    .chat_type
                    .as_deref()
                    .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
                {
                    return None;
                }
                let preview = chat.last_message_preview.as_ref()?;
                let message_id = preview.id.as_deref()?;
                if self
                    .teams
                    .locally_read_through
                    .get(&chat.id)
                    .is_some_and(|id| id == message_id)
                {
                    return None;
                }
                self.teams_preview_is_automation_incoming(
                    &chat.id,
                    message_id,
                    preview,
                    Some(me_id.as_str()),
                )
                .then_some(chat.id.as_str())
            })
            .collect();

        let candidates: Vec<(String, String)> = chats_list
            .iter()
            .filter_map(|chat| {
                // Counts are intentionally limited to one-to-one chats.
                if !chat
                    .chat_type
                    .as_deref()
                    .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"))
                {
                    return None;
                }

                let latest_preview = chat.last_message_preview.as_ref()?;
                let latest_id = latest_preview.id.as_deref()?;
                if self
                    .teams
                    .locally_read_through
                    .get(&chat.id)
                    .is_some_and(|id| id == latest_id)
                {
                    return None;
                }

                let read_at = chat
                    .viewpoint
                    .as_ref()?
                    .last_message_read_date_time
                    .as_deref()?;
                let latest_at = latest_preview.created_date_time.as_deref()?;

                let read_at_parsed = chrono::DateTime::parse_from_rfc3339(read_at).ok()?;
                let latest_at_parsed = chrono::DateTime::parse_from_rfc3339(latest_at).ok()?;
                (latest_at_parsed > read_at_parsed).then(|| (chat.id.clone(), read_at.to_string()))
            })
            .collect();

        // Immediately remove counts for chats that the server viewpoint now
        // considers read, while preserving known counts during a refresh.
        let candidate_ids: std::collections::HashSet<&str> =
            candidates.iter().map(|(id, _)| id.as_str()).collect();
        self.teams.chat_unread_counts.retain(|id, _| {
            candidate_ids.contains(id.as_str())
                || preserve_application_unread.contains(id.as_str())
                || preserve_automation_unread.contains(id.as_str())
        });

        if candidates.is_empty() || self.chat_unread_refresh_running {
            return;
        }
        self.chat_unread_refresh_running = true;

        let s = self.session.clone();
        let tx = self.tx.clone();
        let limiter = self.teams_background_limiter.clone();
        tokio::spawn(async move {
            for (chat_id, read_at) in candidates {
                let read_at = match chrono::DateTime::parse_from_rfc3339(&read_at) {
                    Ok(value) => value,
                    Err(_) => continue,
                };

                let mut count = 0usize;
                let mut next: Option<String> = None;
                let mut first_page = true;

                loop {
                    let result = {
                        let _permit = limiter.acquire(&s.graph, &chat_id).await;
                        if first_page {
                            first_page = false;
                            chats::list_messages_created_desc(&s.graph, &chat_id, 50).await
                        } else if let Some(link) = next.as_deref() {
                            chats::list_messages_more(&s.graph, link).await
                        } else {
                            break;
                        }
                    };

                    let (messages, next_link) = match result {
                        Ok(page) => page,
                        Err(e) => {
                            tracing::warn!(
                                "Teams unread count refresh failed for {chat_id}: {e:#}"
                            );
                            // Keep the previous known count on a transient error.
                            break;
                        }
                    };

                    let mut reached_viewpoint = false;
                    for message in messages {
                        let Some(created) = message
                            .created_date_time
                            .as_deref()
                            .and_then(|value| chrono::DateTime::parse_from_rfc3339(value).ok())
                        else {
                            continue;
                        };

                        if created <= read_at {
                            reached_viewpoint = true;
                            break;
                        }

                        if !teams_message_is_incoming(&message, Some(me_id.as_str())) {
                            continue;
                        }

                        count += 1;
                        if count >= 100 {
                            reached_viewpoint = true;
                            break;
                        }
                    }

                    if reached_viewpoint {
                        let _ = tx
                            .send(AppMessage::ChatUnreadCount {
                                chat_id: chat_id.clone(),
                                count,
                            })
                            .await;
                        break;
                    }

                    next = next_link;
                    if next.is_none() {
                        let _ = tx
                            .send(AppMessage::ChatUnreadCount {
                                chat_id: chat_id.clone(),
                                count,
                            })
                            .await;
                        break;
                    }
                }
            }
            let _ = tx.send(AppMessage::ChatUnreadRefreshDone).await;
        });
    }

    fn clear_teams_conversation_view(&mut self) {
        self.teams.messages.clear();
        self.teams.messages_rendered.clear();
        self.teams.messages_links.clear();
        self.teams.msg_sel = 0;
        self.teams.selected_images.clear();
        self.teams.selected_image_key = None;
        self.teams.messages_next = None;
        self.teams.loading_more = false;
        self.teams.unseen = 0;
        self.teams.replying_to = None;
        if self.teams.editing_message.take().is_some() {
            self.teams.composer.clear();
        }
    }

    fn restore_teams_conversation_cache(&mut self, chat_id: &str) -> bool {
        let Some(cache_root) = self.session.config.cache_dir.as_deref() else {
            return false;
        };

        match load_teams_conversation_cache(cache_root, chat_id) {
            Ok(Some((messages, next))) => {
                let count = messages.len();
                let learned_peer_id = self.remember_contact_from_messages(chat_id, &messages);
                if learned_peer_id {
                    self.load_contact_presences(&[]);
                }
                self.set_teams_messages(messages, next, ListUpdate::Replace);
                tracing::debug!("Teams conversation cache hit for {chat_id}: {count} message(s)");
                true
            }
            Ok(None) => false,
            Err(error) => {
                tracing::warn!("Teams conversation cache read failed for {chat_id}: {error}");
                false
            }
        }
    }

    fn persist_teams_conversation_cache(&self, chat_id: &str) {
        let Some(cache_root) = self.session.config.cache_dir.clone() else {
            return;
        };
        if self.teams.messages.is_empty() {
            return;
        }

        let chat_id = chat_id.to_string();
        let messages = self.teams.messages.clone();
        let cache_max_bytes = self.session.config.cache_max_mb.saturating_mul(1024 * 1024);

        tokio::spawn(async move {
            let cache_chat_id = chat_id.clone();
            match tokio::task::spawn_blocking(move || {
                store_teams_conversation_cache(
                    &cache_root,
                    &cache_chat_id,
                    &messages,
                    cache_max_bytes,
                )
            })
            .await
            {
                Ok(Ok(())) => {
                    tracing::debug!("Teams conversation cache stored for {chat_id}");
                }
                Ok(Err(error)) => {
                    tracing::warn!("Teams conversation cache write failed for {chat_id}: {error}");
                }
                Err(error) => {
                    tracing::warn!(
                        "Teams conversation cache writer task failed for {chat_id}: {error}"
                    );
                }
            }
        });
    }

    pub fn teams_archive_count(&self) -> usize {
        teams_archived_chat_count(&self.teams.chats, &self.teams.stay_archived)
    }

    pub fn teams_archive_row(&self) -> usize {
        teams_active_chat_count(&self.teams.chats, &self.teams.stay_archived)
    }

    pub fn teams_chat_row_count(&self) -> usize {
        teams_chat_row_count(
            &self.teams.chats,
            &self.teams.stay_archived,
            self.teams.archive_expanded,
        )
    }

    pub fn teams_chat_at_row(&self, row: usize) -> Option<&Chat> {
        teams_chat_index_for_row(
            &self.teams.chats,
            &self.teams.stay_archived,
            self.teams.archive_expanded,
            row,
        )
        .and_then(|index| self.teams.chats.get(index))
    }

    pub fn teams_chat_is_archived(&self, chat: &Chat) -> bool {
        teams_chat_is_archived(chat, &self.teams.stay_archived)
    }

    pub fn teams_chat_is_stay_archived(&self, chat_id: &str) -> bool {
        self.teams.stay_archived.contains(chat_id)
    }

    fn teams_selected_chat(&self) -> Option<&Chat> {
        self.teams_chat_at_row(self.teams.chat_sel)
    }

    fn teams_chat_row_for_id(&self, chat_id: &str) -> Option<usize> {
        teams_chat_row_for_id(
            &self.teams.chats,
            &self.teams.stay_archived,
            self.teams.archive_expanded,
            chat_id,
        )
    }

    fn selected_chat_tenant_id(&self, chat: &Chat) -> Option<String> {
        let me_id = self.me.as_ref().map(|me| me.id.as_str());

        self.configured_tenant_id()
            .map(str::to_string)
            .or_else(|| {
                me_id.and_then(|id| {
                    chat.members
                        .iter()
                        .find(|member| member.user_id.as_deref() == Some(id))
                        .and_then(|member| member.tenant_id.clone())
                })
            })
            .or_else(|| self.teams.own_tenant_id.clone())
            .or_else(|| {
                chat.tenant_id
                    .as_deref()
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(str::to_string)
            })
    }

    fn toggle_selected_teams_chat_hidden(&mut self) {
        let Some(chat) = self.teams_selected_chat() else {
            self.status = "Archive is a drawer; select a chat to archive or restore".into();
            return;
        };

        let chat_id = chat.id.clone();
        let server_hidden = teams_chat_is_hidden(chat);
        if server_hidden && self.teams.stay_archived.contains(&chat_id) {
            self.status = "Stay archived is enabled; press X to disable it before restoring".into();
            return;
        }
        let hidden = !server_hidden;
        let Some(user_id) = self.me.as_ref().map(|me| me.id.clone()) else {
            self.status = "still loading your profile".into();
            return;
        };
        let Some(tenant_id) = self.selected_chat_tenant_id(chat) else {
            self.status = "could not determine tenant id for Teams archive action".into();
            return;
        };

        self.status = if hidden {
            "archiving Teams chat…".into()
        } else {
            "restoring Teams chat…".into()
        };

        let s = self.session.clone();
        self.spawn(async move {
            chats::set_hidden(&s.graph, &chat_id, &user_id, &tenant_id, hidden).await?;
            Ok(AppMessage::ChatVisibilityChanged { chat_id, hidden })
        });
    }

    fn toggle_selected_teams_chat_stay_archived(&mut self) {
        let Some(chat) = self.teams_selected_chat() else {
            self.status = "Archive is a drawer; select an archived chat first".into();
            return;
        };

        let chat_id = chat.id.clone();
        let server_hidden = teams_chat_is_hidden(chat);
        let was_enabled = self.teams.stay_archived.contains(&chat_id);

        if !was_enabled && !server_hidden {
            self.status = "archive the chat with x before enabling Stay archived".into();
            return;
        }

        let previous_row = self.teams.chat_sel;
        let enabled = if was_enabled {
            self.teams.stay_archived.remove(&chat_id);
            false
        } else {
            self.teams.stay_archived.insert(chat_id.clone());
            self.teams.archive_expanded = true;
            true
        };

        let persistence_error = self.persist_stay_archived();
        self.teams.chat_sel = self
            .teams_chat_row_for_id(&chat_id)
            .unwrap_or(previous_row.min(self.teams_chat_row_count().saturating_sub(1)));

        let persistent = self.session.config.cache_dir.is_some();
        self.status = if enabled {
            if persistent {
                "Stay archived enabled".into()
            } else {
                "Stay archived enabled for this run".into()
            }
        } else if server_hidden {
            "Stay archived disabled; chat remains archived until restored with x".into()
        } else {
            "Stay archived disabled".into()
        };

        if let Some(error) = persistence_error {
            self.status = format!("{}; persistence failed: {error}", self.status);
        }

        if self.screen == Screen::Teams
            && self.teams.mode == TeamsMode::Chats
            && self.teams.focus == TeamsFocus::List
        {
            self.preview_selected_teams_chat();
        }
    }

    fn rehide_stay_archived_chats(&mut self) {
        let Some(user_id) = self.me.as_ref().map(|me| me.id.clone()) else {
            return;
        };

        let candidates = teams_stay_archived_rehide_candidates(
            &self.teams.chats,
            &self.teams.stay_archived,
            &self.teams.stay_archived_rehide_in_flight,
        );

        let jobs: Vec<(String, String)> = candidates
            .into_iter()
            .filter_map(|chat_id| {
                let chat = self.teams.chats.iter().find(|chat| chat.id == chat_id)?;
                self.selected_chat_tenant_id(chat)
                    .map(|tenant_id| (chat_id, tenant_id))
            })
            .collect();

        for (chat_id, tenant_id) in jobs {
            if !self
                .teams
                .stay_archived_rehide_in_flight
                .insert(chat_id.clone())
            {
                continue;
            }

            let s = self.session.clone();
            let tx = self.tx.clone();
            let limiter = self.teams_background_limiter.clone();
            let task_chat_id = chat_id.clone();
            let task_user_id = user_id.clone();

            tokio::spawn(async move {
                let permit = limiter.acquire(&s.graph, &task_chat_id).await;
                let result =
                    chats::set_hidden(&s.graph, &task_chat_id, &task_user_id, &tenant_id, true)
                        .await;
                drop(permit);

                let error = result.err().map(|error| format!("{error:#}"));
                let _ = tx
                    .send(AppMessage::StayArchivedRehideFinished {
                        chat_id: task_chat_id,
                        error,
                    })
                    .await;
            });
        }
    }

    fn preview_selected_teams_chat(&mut self) {
        if self.teams.mode != TeamsMode::Chats || self.teams.focus != TeamsFocus::List {
            return;
        }

        let Some(chat_id) = self.teams_selected_chat().map(|chat| chat.id.clone()) else {
            self.teams.preview_chat_id = None;
            self.teams.open_chat_id = None;
            self.teams.open_channel = None;
            self.teams.chat_open_pending_read = None;
            self.clear_teams_conversation_view();
            return;
        };

        if self.teams.preview_chat_id.as_deref() == Some(&chat_id) {
            return;
        }

        // Preview selection is local-only: selecting a row never starts a
        // request and never marks it read. A chat that independently qualifies
        // for accelerated background polling may still refresh this preview.
        self.teams.preview_chat_id = Some(chat_id.clone());
        self.teams.open_chat_id = None;
        self.teams.open_channel = None;
        self.teams.chat_open_pending_read = None;
        self.clear_teams_conversation_view();

        let cached = self.restore_teams_conversation_cache(&chat_id);
        if !cached {
            tracing::debug!("no Teams conversation cache for preview {chat_id}");
        }
    }

    fn open_teams_chat(&mut self, chat_id: String) {
        let preview_ready = self.teams.preview_chat_id.as_deref() == Some(&chat_id)
            && !self.teams.messages.is_empty();

        self.teams.mode = TeamsMode::Chats;
        self.teams.open_chat_id = Some(chat_id.clone());
        self.teams.preview_chat_id = None;
        self.teams.last_chat_id = Some(chat_id.clone());
        self.teams.open_channel = None;

        let cached = if preview_ready {
            true
        } else {
            self.clear_teams_conversation_view();
            self.restore_teams_conversation_cache(&chat_id)
        };
        self.teams.focus = TeamsFocus::Messages;
        self.teams.chat_open_pending_read = Some(chat_id.clone());
        self.persist_ui_state();

        self.load_chat_messages(
            chat_id,
            if cached {
                ListUpdate::Merge
            } else {
                ListUpdate::Replace
            },
        );
    }

    fn load_chat_messages(&mut self, chat_id: String, mode: ListUpdate) {
        if !self.begin_teams_foreground_request(&chat_id) {
            tracing::debug!(
                "Teams foreground chat request already queued for {chat_id}; coalescing refresh"
            );
            return;
        }

        let s = self.session.clone();
        let tx = self.tx.clone();
        let limiter = self.teams_background_limiter.clone();
        tokio::spawn(async move {
            let chat_guard = limiter.acquire_foreground_chat(&chat_id).await;
            let result = chats::list_messages(&s.graph, &chat_id, PAGE_SIZE).await;
            drop(chat_guard);
            let msg = match result {
                Ok((messages, next)) => AppMessage::ChatMessages {
                    chat_id,
                    messages,
                    next,
                    mode,
                },
                Err(error) => AppMessage::ChatMessagesFailed {
                    chat_id,
                    mode,
                    error: format!("{error:#}"),
                },
            };
            let _ = tx.send(msg).await;
        });
    }

    fn teams_image_key_for_index(&self, index: usize) -> Option<TeamsImageKey> {
        let message_id = self.teams.messages.get(index)?.id.clone();
        let source = match self.teams.mode {
            TeamsMode::Chats => {
                // Cached chat previews are intentionally not "open", but their
                // already-cached inline images use the same deterministic key.
                // Prefer the real open chat; otherwise use the local preview.
                let chat_id = self
                    .teams
                    .open_chat_id
                    .clone()
                    .or_else(|| self.teams.preview_chat_id.clone())?;
                TeamsImageSource::Chat(chat_id)
            }
            TeamsMode::Channels => {
                let (team_id, channel_id) = self.teams.open_channel.clone()?;
                TeamsImageSource::Channel {
                    team_id,
                    channel_id,
                }
            }
        };

        Some(TeamsImageKey { source, message_id })
    }

    fn selected_teams_image_key(&self) -> Option<TeamsImageKey> {
        self.teams_image_key_for_index(self.teams.msg_sel)
    }

    fn apply_cached_teams_images(&mut self, key: &TeamsImageKey) -> bool {
        let Some(images) = self.teams.image_cache.get(key).cloned() else {
            return false;
        };
        let Some(picker) = self.image_picker.as_ref() else {
            return false;
        };

        self.teams.selected_images = images
            .into_iter()
            .map(|image| MailImage {
                state: std::cell::RefCell::new(picker.new_resize_protocol(image)),
            })
            .collect();

        true
    }

    fn load_preview_teams_images_from_disk(&self, key: TeamsImageKey) {
        let Some(cache_dir) = self.session.config.cache_dir.clone() else {
            return;
        };

        let tx = self.tx.clone();
        tokio::spawn(async move {
            let disk_cache_key = teams_disk_cache_key(&key);
            let read_dir = cache_dir;
            let read_key = disk_cache_key;

            match tokio::task::spawn_blocking(move || load_teams_disk_cache(&read_dir, &read_key))
                .await
            {
                Ok(Ok(Some(images))) => {
                    tracing::debug!(
                        "Teams image preview disk cache hit for {:?}: {} image(s)",
                        key,
                        images.len()
                    );
                    let _ = tx.send(AppMessage::TeamsImages { key, images }).await;
                }
                Ok(Ok(None)) => {
                    tracing::debug!("Teams image preview disk cache miss for {:?}", key);
                }
                Ok(Err(error)) => {
                    tracing::debug!(
                        "Teams image preview disk cache read failed for {:?}: {error}",
                        key
                    );
                }
                Err(error) => {
                    tracing::debug!(
                        "Teams image preview disk cache reader task failed for {:?}: {error}",
                        key
                    );
                }
            }
        });
    }

    fn refresh_selected_teams_images(&mut self) {
        if self.image_picker.is_none() {
            self.teams.selected_images.clear();
            self.teams.selected_image_key = None;
            return;
        }

        let Some(key) = self.selected_teams_image_key() else {
            self.teams.selected_images.clear();
            self.teams.selected_image_key = None;
            return;
        };

        let preview_cache_only = self.teams.mode == TeamsMode::Chats
            && self.teams.open_chat_id.is_none()
            && self.teams.preview_chat_id.is_some();

        if self.teams.selected_image_key.as_ref() != Some(&key) {
            self.teams.selected_images.clear();
            self.teams.selected_image_key = Some(key.clone());

            let ram_hit = self.apply_cached_teams_images(&key);
            if preview_cache_only && !ram_hit {
                // Preview may read what is already on disk, but must not trigger
                // a Graph or Files.Read.All request on a cache miss.
                self.load_preview_teams_images_from_disk(key.clone());
            }
        }

        if !preview_cache_only {
            // Open chats/channels retain the normal RAM -> disk -> Graph path.
            self.prefetch_teams_images();
        }
    }

    fn prefetch_teams_images(&mut self) {
        const PREFETCH_RADIUS: usize = 4;

        if self.image_picker.is_none() || self.teams.messages.is_empty() {
            return;
        }

        let center = self
            .teams
            .msg_sel
            .min(self.teams.messages.len().saturating_sub(1));
        let mut indices = Vec::with_capacity(PREFETCH_RADIUS * 2 + 1);
        indices.push(center);
        for offset in 1..=PREFETCH_RADIUS {
            if let Some(index) = center.checked_sub(offset) {
                indices.push(index);
            }
            let index = center + offset;
            if index < self.teams.messages.len() {
                indices.push(index);
            }
        }

        for index in indices {
            let Some(key) = self.teams_image_key_for_index(index) else {
                continue;
            };

            if self.teams.image_cache.contains_key(&key)
                || self.teams.image_in_flight.contains(&key)
            {
                continue;
            }

            let (mut hosted_ids, mut reference_files, reference_owner_user_id) = {
                let Some(message) = self.teams.messages.get(index) else {
                    continue;
                };
                let hosted_ids = hosted_content_ids(message);
                let reference_files = reference_image_files(message);
                let owner = message
                    .from
                    .as_ref()
                    .and_then(|from| from.user.as_ref())
                    .and_then(|user| user.id.clone());
                (hosted_ids, reference_files, owner)
            };

            hosted_ids.truncate(MAX_TEAMS_IMAGES);

            if !self.session.config.teams_file_images {
                reference_files.clear();
            } else if !self.session.config.can_read_files() {
                if index == center && !reference_files.is_empty() {
                    self.status =
                        "M365_TEAMS_FILE_IMAGES=1 requires Files.Read.All in M365_SCOPES".into();
                }
                reference_files.clear();
            }

            let remaining = MAX_TEAMS_IMAGES.saturating_sub(hosted_ids.len());
            reference_files.truncate(remaining);

            if hosted_ids.is_empty() && reference_files.is_empty() {
                self.teams.image_cache.insert(key, Vec::new());
                continue;
            }

            self.teams.image_in_flight.insert(key.clone());
            tracing::debug!(
                "Teams image prefetch queued for {:?}: hosted={}, reference={}",
                key,
                hosted_ids.len(),
                reference_files.len()
            );

            let s = self.session.clone();
            let task_key = key.clone();
            self.spawn(async move {
                let disk_cache_dir = s.config.cache_dir.clone();
                let disk_cache_max_bytes = s
                    .config
                    .cache_max_mb
                    .saturating_mul(1024 * 1024);
                let disk_cache_key = teams_disk_cache_key(&task_key);

                if let Some(cache_dir) = disk_cache_dir.as_ref() {
                    let read_dir = cache_dir.clone();
                    let read_key = disk_cache_key.clone();
                    match tokio::task::spawn_blocking(move || {
                        load_teams_disk_cache(&read_dir, &read_key)
                    })
                    .await
                    {
                        Ok(Ok(Some(images))) => {
                            tracing::debug!(
                                "Teams image disk cache hit for {:?}: {} image(s)",
                                task_key,
                                images.len()
                            );
                            return Ok(AppMessage::TeamsImages {
                                key: task_key,
                                images,
                            });
                        }
                        Ok(Ok(None)) => {}
                        Ok(Err(error)) => {
                            tracing::debug!(
                                "Teams image disk cache read failed for {:?}: {error}",
                                task_key
                            );
                        }
                        Err(error) => {
                            tracing::debug!(
                                "Teams image disk cache reader task failed for {:?}: {error}",
                                task_key
                            );
                        }
                    }
                }

                let mut images =
                    Vec::with_capacity(hosted_ids.len() + reference_files.len());

                for hosted_content_id in hosted_ids {
                    let result = match &task_key.source {
                        TeamsImageSource::Chat(chat_id) => {
                            chats::hosted_content_bytes(
                                &s.graph,
                                chat_id,
                                &task_key.message_id,
                                &hosted_content_id,
                            )
                            .await
                        }
                        TeamsImageSource::Channel {
                            team_id,
                            channel_id,
                        } => {
                            channels::hosted_content_bytes(
                                &s.graph,
                                team_id,
                                channel_id,
                                &task_key.message_id,
                                &hosted_content_id,
                            )
                            .await
                        }
                    };

                    let bytes = match result {
                        Ok(bytes) => bytes,
                        Err(error) => {
                            tracing::debug!(
                                "Teams hosted image {hosted_content_id} could not be downloaded: {error:#}"
                            );
                            continue;
                        }
                    };

                    if bytes.len() > MAX_TEAMS_IMAGE_BYTES {
                        tracing::debug!(
                            "Teams hosted image {hosted_content_id} exceeds {} bytes; skipping",
                            MAX_TEAMS_IMAGE_BYTES
                        );
                        continue;
                    }

                    let decoded =
                        tokio::task::spawn_blocking(move || image::load_from_memory(&bytes)).await;
                    match decoded {
                        Ok(Ok(image)) if image.width() > 1 && image.height() > 1 => {
                            images.push(
                                image.thumbnail(
                                    MAX_TEAMS_IMAGE_WIDTH,
                                    MAX_TEAMS_IMAGE_HEIGHT,
                                ),
                            );
                        }
                        Ok(Ok(_)) => {
                            // Ignore tracking-size pixels.
                        }
                        Ok(Err(error)) => {
                            tracing::debug!(
                                "Teams hosted image {hosted_content_id} could not be decoded: {error}"
                            );
                        }
                        Err(error) => {
                            tracing::debug!(
                                "Teams hosted image decoder task failed for {hosted_content_id}: {error}"
                            );
                        }
                    }
                }

                for (name, content_url) in reference_files {
                    let bytes = match chats::reference_file_bytes(
                        &s.graph,
                        reference_owner_user_id.as_deref(),
                        &name,
                        &content_url,
                    )
                    .await
                    {
                        Ok(bytes) => bytes,
                        Err(error) => {
                            tracing::debug!(
                                "Teams image file attachment {name} could not be downloaded: {error:#}"
                            );
                            continue;
                        }
                    };

                    if bytes.len() > MAX_TEAMS_IMAGE_BYTES {
                        tracing::debug!(
                            "Teams image file attachment {name} exceeds {} bytes; skipping",
                            MAX_TEAMS_IMAGE_BYTES
                        );
                        continue;
                    }

                    let decoded =
                        tokio::task::spawn_blocking(move || image::load_from_memory(&bytes)).await;
                    match decoded {
                        Ok(Ok(image)) if image.width() > 1 && image.height() > 1 => {
                            images.push(
                                image.thumbnail(
                                    MAX_TEAMS_IMAGE_WIDTH,
                                    MAX_TEAMS_IMAGE_HEIGHT,
                                ),
                            );
                        }
                        Ok(Ok(_)) => {
                            // Ignore tracking-size pixels.
                        }
                        Ok(Err(error)) => {
                            tracing::debug!(
                                "Teams image file attachment {name} could not be decoded: {error}"
                            );
                        }
                        Err(error) => {
                            tracing::debug!(
                                "Teams image file decoder task failed for {name}: {error}"
                            );
                        }
                    }
                }

                if let Some(cache_dir) = disk_cache_dir {
                    if !images.is_empty() {
                        let write_images = images.clone();
                        let write_key = disk_cache_key.clone();
                        let cache_key_for_log = task_key.clone();
                        match tokio::task::spawn_blocking(move || {
                            store_teams_disk_cache(
                                &cache_dir,
                                &write_key,
                                &write_images,
                                disk_cache_max_bytes,
                            )
                        })
                        .await
                        {
                            Ok(Ok(())) => {
                                tracing::debug!(
                                    "Teams image disk cache stored for {:?}: {} image(s)",
                                    cache_key_for_log,
                                    images.len()
                                );
                            }
                            Ok(Err(error)) => {
                                tracing::debug!(
                                    "Teams image disk cache write failed for {:?}: {error}",
                                    cache_key_for_log
                                );
                            }
                            Err(error) => {
                                tracing::debug!(
                                    "Teams image disk cache writer task failed for {:?}: {error}",
                                    cache_key_for_log
                                );
                            }
                        }
                    }
                }

                Ok(AppMessage::TeamsImages {
                    key: task_key,
                    images,
                })
            });
        }
    }

    fn load_more_teams_messages(&mut self) {
        if self.teams.loading_more {
            return;
        }
        let Some(next_link) = self.teams.messages_next.clone() else {
            return; // reached the start of the conversation
        };
        self.teams.loading_more = true;
        self.status = "loading older messages…".into();
        let s = self.session.clone();

        if let Some(chat_id) = self.teams.open_chat_id.clone() {
            if !self.begin_teams_foreground_request(&chat_id) {
                self.teams.loading_more = false;
                self.status = "chat refresh already in progress".into();
                return;
            }
            let tx = self.tx.clone();
            let limiter = self.teams_background_limiter.clone();
            tokio::spawn(async move {
                let chat_guard = limiter.acquire_foreground_chat(&chat_id).await;
                let result = chats::list_messages_more(&s.graph, &next_link).await;
                drop(chat_guard);
                let msg = match result {
                    Ok((messages, next)) => AppMessage::ChatMessages {
                        chat_id,
                        messages,
                        next,
                        mode: ListUpdate::Append,
                    },
                    Err(error) => AppMessage::ChatMessagesFailed {
                        chat_id,
                        mode: ListUpdate::Append,
                        error: format!("{error:#}"),
                    },
                };
                let _ = tx.send(msg).await;
            });
        } else if let Some((team_id, channel_id)) = self.teams.open_channel.clone() {
            self.spawn(async move {
                let (messages, next) = channels::list_messages_more(&s.graph, &next_link).await?;
                Ok(AppMessage::ChannelMessages {
                    team_id,
                    channel_id,
                    messages,
                    next,
                    mode: ListUpdate::Append,
                })
            });
        } else {
            self.teams.loading_more = false;
        }
    }

    fn load_teams(&self) {
        let s = self.session.clone();
        self.spawn(async move { Ok(AppMessage::Teams(channels::joined_teams(&s.graph).await?)) });
    }

    fn load_channels(&self, team_id: String) {
        let s = self.session.clone();
        self.spawn(async move {
            let channels = channels::list_channels(&s.graph, &team_id).await?;
            Ok(AppMessage::Channels { team_id, channels })
        });
    }

    fn load_channel_messages(&self, team_id: String, channel_id: String, mode: ListUpdate) {
        let s = self.session.clone();
        self.spawn(async move {
            let (messages, next) =
                channels::list_messages(&s.graph, &team_id, &channel_id, PAGE_SIZE).await?;
            Ok(AppMessage::ChannelMessages {
                team_id,
                channel_id,
                messages,
                next,
                mode,
            })
        });
    }

    fn send_chat_message(&mut self, chat_id: String, text: String) {
        let s = self.session.clone();
        self.spawn(async move {
            let message =
                chats::send_message(&s.graph, &chat_id, &text, s.config.teams_markdown).await?;
            Ok(AppMessage::ChatMessageSent {
                chat_id,
                message_id: message.id,
                status: "message sent".into(),
            })
        });
    }

    fn send_channel_message(
        &self,
        team_id: String,
        channel_id: String,
        prepared: crate::images::Prepared,
    ) {
        let s = self.session.clone();
        self.spawn(async move {
            let body = prepared_body(&prepared);
            match if let crate::images::Prepared::Text(text) = &prepared {
                channels::send_message(
                    &s.graph,
                    &team_id,
                    &channel_id,
                    text,
                    s.config.teams_markdown,
                )
                .await
            } else {
                channels::send_message_with_body(&s.graph, &team_id, &channel_id, body).await
            } {
                Ok(_) => Ok(AppMessage::TeamsSent("message posted".into())),
                Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
            }
        });
    }

    fn open_chat_with_email(&self, email: String) {
        let s = self.session.clone();
        let me_id = self.me.as_ref().map(|m| m.id.clone());
        self.spawn(async move {
            let Some(me_id) = me_id else {
                return Ok(AppMessage::Error("still loading your profile".into()));
            };
            let id = navigation::chat_id_for_email(&s, &me_id, &email).await?;
            Ok(AppMessage::OpenChat(id))
        });
    }

    pub fn apply(&mut self, msg: AppMessage) {
        match msg {
            AppMessage::Status(s) => self.status = s,
            AppMessage::Error(e) => {
                self.status = format!("error: {}", m365_core::util::graph_error_summary(&e));
            }
            AppMessage::Whoami(u) => {
                self.status = format!(
                    "signed in as {} <{}>",
                    u.display_name.clone().unwrap_or_default(),
                    u.best_email().unwrap_or("")
                );
                self.me = Some(u);
                let chats = self.teams.chats.clone();
                if !chats.is_empty() {
                    self.remember_contacts_from_chats(&chats);
                    self.load_contact_presences(&chats);
                }
            }
            AppMessage::Folders(f) => {
                // Folder loading is part of every network poll, so any successful
                // result proves Graph is responsive and resets poll health/backoff.
                self.mark_poll_success();
                let first_load = self.outlook.folders.is_empty();
                self.outlook.folders = f;
                if first_load {
                    // Prefer Inbox as the initial selection, then load it.
                    if let Some(i) = self
                        .outlook
                        .folders
                        .iter()
                        .position(|f| f.display_name.as_deref() == Some("Inbox"))
                    {
                        self.outlook.folder_sel = i;
                    }
                    if let Some(f) = self.outlook.folders.get(self.outlook.folder_sel) {
                        self.load_messages(f.id.clone(), ListUpdate::Replace);
                    }
                } else {
                    // Refresh (poll): keep the user's selection, just clamp it.
                    self.outlook.folder_sel = self
                        .outlook
                        .folder_sel
                        .min(self.outlook.folders.len().saturating_sub(1));
                }
            }
            AppMessage::Messages { items, next, mode } => {
                self.cache_mail_contacts(&items);
                self.outlook.loading_more = false;
                let selected_key = self.current_mail().map(|m| self.mail_selection_key(m));

                match mode {
                    ListUpdate::Replace => {
                        self.outlook.messages = items;
                        self.outlook.messages_next = next;
                    }
                    ListUpdate::Append => {
                        self.outlook.messages.extend(items);
                        self.outlook.messages_next = next;
                        self.status = if self.outlook.messages_next.is_some() {
                            format!("{} loaded (more available)", self.outlook.messages.len())
                        } else {
                            format!("{} loaded (all)", self.outlook.messages.len())
                        };
                    }
                    ListUpdate::Merge => {
                        // Refresh the newest window without discarding the older
                        // pages already loaded; `messages_next` is left alone as
                        // it already points past everything we hold.
                        let existing = std::mem::take(&mut self.outlook.messages);
                        self.outlook.messages =
                            merge_newest_first(items, existing, |m| m.id.clone());
                    }
                }

                self.rebuild_mail_rows();
                self.outlook.msg_sel = selected_key
                    .and_then(|key| {
                        self.outlook.message_rows.iter().position(|&i| {
                            self.outlook
                                .messages
                                .get(i)
                                .is_some_and(|m| self.mail_selection_key(m) == key)
                        })
                    })
                    .unwrap_or(self.outlook.msg_sel)
                    .min(self.outlook.message_rows.len().saturating_sub(1));
                if self.settings.preview_mail_on_hover
                    && self.outlook_focus == OutlookFocus::Messages
                {
                    self.open_current_mail();
                }
            }
            AppMessage::MessageBody(m) => {
                self.cache_mail_contacts(std::slice::from_ref(&m));
                let id = m.id.clone();
                self.outlook.message_cache.insert(id.clone(), m.clone());
                let key = ReadingKey::Message(id);
                if self.outlook.pending_reading.as_ref() == Some(&key) {
                    self.outlook.pending_reading = None;
                }
                if self.selection_reading_key().as_ref() == Some(&key) {
                    self.show_message(m);
                }
            }
            AppMessage::AutoMarkReadDue { id, generation } => {
                let still_open = generation == self.read_timer_generation
                    && self.screen == Screen::Outlook
                    && self.outlook_focus == OutlookFocus::Reading
                    && self.outlook.reading.as_ref().is_some_and(|message| {
                        message.id == id && !message.is_read.unwrap_or(false)
                    });
                if still_open {
                    self.cancel_read_timer();
                    self.set_mail_read(&id, true);
                }
            }
            AppMessage::CalendarReminders(events) => {
                self.calendar_reminder_events = events;
                self.check_calendar_reminders();
            }
            AppMessage::Calendar(events) => {
                let selected_id = self
                    .calendar
                    .events
                    .get(self.calendar.selected)
                    .map(|event| event.id.clone());
                self.calendar.events = events;
                self.calendar.selected = selected_id
                    .and_then(|id| self.calendar.events.iter().position(|event| event.id == id))
                    .unwrap_or(self.calendar.selected)
                    .min(self.calendar.events.len().saturating_sub(1));

                if self.calendar.view == CalendarView::Month
                    && !self.calendar.events.is_empty()
                    && !calendar_event_intersects_month(
                        &self.calendar.events[self.calendar.selected],
                        self.calendar.month_offset,
                    )
                {
                    self.calendar.selected = self
                        .calendar
                        .events
                        .iter()
                        .position(|event| {
                            calendar_event_intersects_month(event, self.calendar.month_offset)
                        })
                        .unwrap_or(0);
                }
            }
            AppMessage::Chats(c) => {
                let selected_chat_id = self.teams_selected_chat().map(|chat| chat.id.clone());
                let archive_selected =
                    !self.teams.chats.is_empty() && self.teams.chat_sel == self.teams_archive_row();

                self.sync_teams_hot_activity_from_chats(&c);
                self.notify_for_chats(&c);
                self.remember_contacts_from_chats(&c);
                self.load_contact_presences(&c);
                self.warm_teams_conversation_caches(&c);

                let open_index = self
                    .teams
                    .open_chat_id
                    .as_deref()
                    .and_then(|open_id| c.iter().position(|chat| chat.id == open_id));

                if self.screen == Screen::Teams && self.teams.focus == TeamsFocus::Messages {
                    if let Some(index) = open_index {
                        let chat = &c[index];
                        self.teams.chat_unread_counts.remove(&chat.id);
                        if let Some(latest_id) = chat
                            .last_message_preview
                            .as_ref()
                            .and_then(|preview| preview.id.clone())
                        {
                            self.teams
                                .locally_read_through
                                .insert(chat.id.clone(), latest_id);
                        }
                    }
                }

                self.refresh_chat_unread_counts(&c);
                self.teams.chats = c;

                let visible_id = self
                    .teams
                    .preview_chat_id
                    .clone()
                    .or_else(|| self.teams.open_chat_id.clone())
                    .or(selected_chat_id);

                self.teams.chat_sel = visible_id
                    .as_deref()
                    .and_then(|id| self.teams_chat_row_for_id(id))
                    .or_else(|| archive_selected.then(|| self.teams_archive_row()))
                    .unwrap_or(self.teams.chat_sel)
                    .min(self.teams_chat_row_count().saturating_sub(1));

                if self.screen == Screen::Teams
                    && self.teams.mode == TeamsMode::Chats
                    && self.teams.focus == TeamsFocus::List
                    && self.teams.preview_chat_id.is_none()
                {
                    self.preview_selected_teams_chat();
                }

                // Teams automatically unhides a hidden chat when a new message
                // arrives. Keep the local Archive view stable, process all
                // normal chat state first, then restore the server hide state
                // as low-priority maintenance work.
                self.rehide_stay_archived_chats();
            }
            AppMessage::ChatVisibilityChanged { chat_id, hidden } => {
                let previous_row = self.teams.chat_sel;
                if let Some(chat) = self.teams.chats.iter_mut().find(|chat| chat.id == chat_id) {
                    chat.viewpoint
                        .get_or_insert_with(Default::default)
                        .is_hidden = Some(hidden);
                }

                self.status = if hidden {
                    "Teams chat archived".into()
                } else {
                    "Teams chat restored".into()
                };

                self.teams.chat_sel = self
                    .teams_chat_row_for_id(&chat_id)
                    .unwrap_or(previous_row.min(self.teams_chat_row_count().saturating_sub(1)));

                if self.screen == Screen::Teams
                    && self.teams.mode == TeamsMode::Chats
                    && self.teams.focus == TeamsFocus::List
                {
                    self.preview_selected_teams_chat();
                }

                self.load_chats();
            }
            AppMessage::StayArchivedRehideFinished { chat_id, error } => {
                self.teams.stay_archived_rehide_in_flight.remove(&chat_id);

                if let Some(error) = error {
                    tracing::warn!("Stay archived re-hide failed for {chat_id}: {error}");
                    self.status = "Stay archived re-hide failed; will retry".into();
                } else {
                    if let Some(chat) = self.teams.chats.iter_mut().find(|chat| chat.id == chat_id)
                    {
                        chat.viewpoint
                            .get_or_insert_with(Default::default)
                            .is_hidden = Some(true);
                    }
                    tracing::debug!("Stay archived re-hide completed for {chat_id}");
                }
            }
            AppMessage::ContactProfileLoaded {
                chat_id,
                person,
                avatar,
            } => {
                if let Some(profile) = self.contact_profile.as_mut() {
                    if profile.chat_id == chat_id {
                        profile.person = person;
                        profile.loading = false;
                        if let (Some(image), Some(picker)) = (avatar, self.image_picker.as_ref()) {
                            profile.avatar = Some(MailImage {
                                state: std::cell::RefCell::new(picker.new_resize_protocol(image)),
                            });
                        }
                    }
                }
            }
            AppMessage::ContactPresences(items) => {
                for presence in items {
                    if let Some(id) = presence.id.clone() {
                        self.teams.contact_presences.insert(id, presence);
                    }
                }
            }
            AppMessage::ChatUnreadCount { chat_id, count } => {
                if count == 0 {
                    self.teams.chat_unread_counts.remove(&chat_id);
                } else {
                    self.teams.chat_unread_counts.insert(chat_id, count);
                }
            }
            AppMessage::ChatUnreadRefreshDone => {
                self.chat_unread_refresh_running = false;
            }
            AppMessage::ChatMessages {
                chat_id,
                messages,
                next,
                mode,
            } => {
                if let Some(state) = self.teams_hot_chat_state.get_mut(&chat_id) {
                    state.foreground_in_flight = false;
                    state.last_poll_started = Some(std::time::Instant::now());
                }
                self.observe_teams_chat_page(&chat_id, &messages);
                if self.teams.open_chat_id.as_deref() == Some(&chat_id) {
                    let opened_now = self.teams.chat_open_pending_read.as_deref() == Some(&chat_id);
                    let learned_peer_id = self.remember_contact_from_messages(&chat_id, &messages);
                    if learned_peer_id {
                        self.load_contact_presences(&[]);
                    }

                    self.set_teams_messages(messages, next, mode);
                    self.persist_teams_conversation_cache(&chat_id);

                    if opened_now {
                        self.teams.chat_open_pending_read = None;
                        self.teams.chat_unread_counts.remove(&chat_id);
                        self.mark_teams_chat_read(chat_id.clone());

                        if let Some(latest_id) = self
                            .teams
                            .chats
                            .iter()
                            .find(|chat| chat.id == chat_id)
                            .and_then(|chat| chat.last_message_preview.as_ref())
                            .and_then(|preview| preview.id.clone())
                        {
                            self.teams.locally_read_through.insert(chat_id, latest_id);
                        }
                    }
                }
            }
            AppMessage::ChatMessagesFailed {
                chat_id,
                mode,
                error,
            } => {
                if mode == ListUpdate::Append {
                    self.teams.loading_more = false;
                }
                if let Some(state) = self.teams_hot_chat_state.get_mut(&chat_id) {
                    state.foreground_in_flight = false;
                    state.last_poll_started = Some(std::time::Instant::now());
                }
                self.status = format!("error: {}", m365_core::util::graph_error_summary(&error));
            }
            AppMessage::ChatMessageSent {
                chat_id,
                message_id,
                status,
            } => {
                self.teams.sending = false;
                self.teams.pending_restore = None;
                self.teams.editing_message = None;
                if self.own_sent_chat_message_ids.len() >= 1000 {
                    self.own_sent_chat_message_ids.clear();
                }
                self.own_sent_chat_message_ids.insert(message_id);
                self.mark_teams_chat_active(&chat_id);
                self.status = status;
                self.refresh_current();
            }
            AppMessage::ChatPinsLoaded {
                chat_id,
                messages,
                status,
            } => {
                if self.teams.open_chat_id.as_deref() == Some(chat_id.as_str()) {
                    self.status = status.unwrap_or_else(|| match messages.len() {
                        0 => "no pinned message".into(),
                        1 => "pinned message loaded".into(),
                        count => format!("{count} pinned messages returned by Graph"),
                    });
                    self.overlay = Some(Overlay::PinnedMessages {
                        chat_id,
                        messages,
                        sel: 0,
                    });
                }
            }
            AppMessage::ChatMarkedUnread { chat_id } => {
                self.teams.chat_unread_counts.insert(chat_id.clone(), 1);
                self.teams.locally_read_through.remove(&chat_id);
                self.status = "chat marked unread".into();
                self.load_chats();
            }
            AppMessage::ChatTopicUpdated {
                chat_id,
                topic,
                status,
            } => {
                if let Some(chat) = self.teams.chats.iter_mut().find(|chat| chat.id == chat_id) {
                    chat.topic = Some(topic);
                }
                self.status = status;
                self.load_chats();
            }
            AppMessage::HotChatPollFinished {
                chat_id,
                messages,
                next,
                error,
            } => {
                self.apply_hot_chat_poll(chat_id, messages, next, error);
            }
            AppMessage::HotChatPollSkipped { chat_id } => {
                if let Some(state) = self.teams_hot_chat_state.get_mut(&chat_id) {
                    state.hot_pending = false;
                    state.last_poll_started = Some(std::time::Instant::now());
                }
            }
            AppMessage::ChatCacheWarmed { chat_id, messages } => {
                let learned_peer_id = self.remember_contact_from_messages(&chat_id, &messages);
                if learned_peer_id {
                    self.load_contact_presences(&[]);
                }

                let still_previewing = self.screen == Screen::Teams
                    && self.teams.mode == TeamsMode::Chats
                    && self.teams.focus == TeamsFocus::List
                    && self.teams.open_chat_id.is_none()
                    && self.teams.preview_chat_id.as_deref() == Some(&chat_id);

                if still_previewing {
                    self.set_teams_messages(messages, None, ListUpdate::Replace);
                }
            }
            AppMessage::Teams(t) => self.teams.teams = t,
            AppMessage::Channels { team_id, channels } => {
                if self.teams.teams.get(self.teams.team_sel).map(|t| &t.id) == Some(&team_id) {
                    self.teams.channels = channels;
                    self.teams.channel_sel = 0;
                }
            }
            AppMessage::ChannelMessages {
                team_id,
                channel_id,
                messages,
                next,
                mode,
            } => {
                if self.teams.open_channel.as_ref() == Some(&(team_id, channel_id)) {
                    self.set_teams_messages(messages, next, mode);
                }
            }
            AppMessage::TeamsImages { key, images } => {
                self.teams.image_in_flight.remove(&key);
                let image_count = images.len();
                self.teams.image_cache.insert(key.clone(), images);

                tracing::debug!(
                    "Teams image RAM cache stored for {:?}: {} image(s)",
                    key,
                    image_count
                );

                if self.selected_teams_image_key().as_ref() == Some(&key) {
                    self.teams.selected_image_key = Some(key.clone());
                    let _ = self.apply_cached_teams_images(&key);
                }
            }
            AppMessage::Done(s) => {
                self.status = s;
                self.refresh_current();
            }
            AppMessage::OpenChat(Some(id)) => {
                self.cancel_read_timer();
                self.screen = Screen::Teams;
                self.teams_unread = false;
                self.open_teams_chat(id);
                self.load_chats();
            }
            AppMessage::OpenChat(None) => {
                self.status = "that sender isn't a Teams user in your directory".into();
            }
            AppMessage::Presence {
                presence,
                requested,
            } => self.apply_presence_result(presence, requested),
            AppMessage::PresenceRefresh(presence) => {
                self.my_presence = Some(presence);
            }
            AppMessage::PresenceStatusMessageUpdated { presence, cleared } => {
                self.my_presence = Some(presence);
                self.status = if cleared {
                    "status message cleared".into()
                } else {
                    "status message updated".into()
                };
            }
            AppMessage::PresenceSessionWriteFinished {
                write,
                requested,
                presence,
                success,
                error,
                retry_preference,
            } => {
                self.presence_state
                    .finish_write(write, success, std::time::Instant::now());

                if self.presence_preference_pending.is_none() {
                    self.presence_preference_pending = retry_preference;
                }

                if success {
                    if let Some(presence) = presence {
                        self.apply_presence_result(presence, requested);
                    } else if let Some(error) = error {
                        self.status = format!("presence session updated; {error}");
                    }
                } else {
                    let error = error.unwrap_or_else(|| "unknown Graph error".into());
                    tracing::warn!("presence session update failed: {error}");
                    self.status = format!("presence update failed: {error}");
                }

                // A newer desired state may have arrived while this write was
                // running. Start it only after the old write has completed.
                self.pump_presence_session_write();
            }
            AppMessage::NtfyWorkPlan(working) => {
                self.ntfy_working_now = Some(working);
            }
            AppMessage::OutOfOfficeLoaded { settings, edit } => {
                self.status.clear();
                self.overlay = Some(if edit {
                    Overlay::OutOfOfficeEdit(Box::new(out_of_office::Editor::from_settings(
                        &settings,
                    )))
                } else {
                    Overlay::OutOfOffice(Box::new(out_of_office::View::from_settings(settings)))
                });
            }
            AppMessage::OutOfOfficeSaved(status) => {
                self.overlay = None;
                self.status = status;
            }
            AppMessage::DiagnosticsLoaded(remote) => {
                self.diagnostics.remote = Some(remote);
                self.diagnostics.loading = false;
            }
            AppMessage::ContactDiagnosticsLoaded(remote) => {
                self.contact_diagnostics.remote = Some(remote);
                self.contact_diagnostics.loading = false;
            }
            AppMessage::InboxPeek(items) => self.notify_for_mail(&items),
            AppMessage::Attachments { message_id, items } => {
                self.outlook.attachment_loading.remove(&message_id);
                self.outlook
                    .attachment_cache
                    .insert(message_id.clone(), items.clone());
                // Ignore a late response for a message we've navigated away from.
                if self.outlook.reading.as_ref().map(|m| m.id.as_str()) == Some(&message_id) {
                    // Inline images (signatures, logos) aren't useful downloads.
                    self.outlook.reading_attachments = items
                        .into_iter()
                        .filter(|a| !a.is_inline.unwrap_or(false))
                        .collect();
                }
                self.fetch_mail_images();
            }
            AppMessage::Downloaded { name, bytes } => {
                let size = bytes.len();
                match crate::files::save(&name, &bytes) {
                    Ok(path) => {
                        self.status = format!("saved {} ({size} bytes) to {}", name, path.display())
                    }
                    Err(e) => self.status = format!("could not save {name}: {e:#}"),
                }
            }
            AppMessage::Push(state) => {
                // Surface a broken tunnel instead of silently falling back.
                if let PushState::Failed(reason) = &state {
                    self.status = format!("push unavailable, using 20s polling — {reason}");
                }
                self.push = state;
            }
            AppMessage::Tick => {
                if self
                    .notice_until
                    .is_some_and(|until| std::time::Instant::now() >= until)
                {
                    if matches!(self.overlay.as_ref(), Some(Overlay::Notice(_))) {
                        self.overlay = None;
                    }
                    self.notice_until = None;
                }
                self.refresh_presence_activity();
                self.pump_presence_session_write();
                self.expire_ntfy_snooze_if_due();
                self.rss_kb = read_rss_kb();
                // Clear a message once it has sat unchanged for a while, so the
                // status bar never shows something from half an hour ago.
                if self.status.is_empty() {
                    self.status_ticks = 0;
                } else if self.status == self.last_status {
                    self.status_ticks += 1;
                    if self.status_ticks >= STATUS_TICKS_TO_LIVE {
                        self.status.clear();
                        self.status_ticks = 0;
                    }
                } else {
                    self.last_status = self.status.clone();
                    self.status_ticks = 0;
                }

                self.refresh_calendar_reminders_if_due();
                self.check_calendar_reminders();
            }
            AppMessage::HotPollTick => self.poll_hot_chats_if_due(),
            AppMessage::Poll => self.scheduled_poll(),
            AppMessage::MailThread {
                conversation_id,
                items,
                truncated,
            } => {
                self.cache_mail_contacts(&items);
                self.outlook.message_cache.extend(
                    items
                        .iter()
                        .cloned()
                        .map(|message| (message.id.clone(), message)),
                );
                self.outlook
                    .thread_cache
                    .insert(conversation_id.clone(), (items.clone(), truncated));
                let key = ReadingKey::Thread(conversation_id);
                if self.outlook.pending_reading.as_ref() == Some(&key) {
                    self.outlook.pending_reading = None;
                }
                if self.selection_reading_key().as_ref() == Some(&key) {
                    self.show_thread(items, truncated);
                }
            }
            AppMessage::PeopleSearch { query, result } => {
                if let Some(Overlay::NewChat {
                    query: current,
                    results,
                    sel,
                    loading,
                }) = &mut self.overlay
                {
                    if *current == query {
                        *loading = false;
                        match result {
                            Ok(users) => {
                                *results = users;
                                *sel = 0;
                            }
                            Err(error) => {
                                self.status = format!("directory search: {error}");
                            }
                        }
                    }
                }
            }
            AppMessage::MailSent(s) => {
                self.status = s;
                self.outlook.thread_cache.clear();
                self.refresh_current();
            }
            AppMessage::TeamsSent(s) => {
                self.teams.sending = false;
                self.teams.pending_restore = None;
                self.teams.editing_message = None;
                self.status = s;
                self.refresh_current();
            }
            AppMessage::TeamsSendFailed(e) => {
                self.teams.sending = false;
                if let Some((text, images)) = self.teams.pending_restore.take() {
                    if self.teams.composer.is_empty() && self.teams.images.is_empty() {
                        self.teams.composer = TextInput::from(text.as_str());
                        self.teams.images = images;
                        self.rebuild_composer_previews();
                    }
                }
                self.status = format!("error: {}", m365_core::util::graph_error_summary(&e));
            }
            AppMessage::HostedImage { key, bytes } => {
                self.image_loading.remove(&key);
                let max_rows = if crate::termimg::is_photo_key(&key) {
                    crate::termimg::AVATAR_IMAGE_ROWS
                } else if key.starts_with("mail:") {
                    crate::termimg::MAIL_IMAGE_ROWS
                } else {
                    crate::termimg::CONVO_IMAGE_ROWS
                };
                let decoded = self
                    .graphics
                    .as_ref()
                    .and_then(|g| g.decode(&bytes, max_rows).ok());
                match decoded {
                    Some(img) => {
                        if let Some(root) = self.session.config.cache_dir.clone() {
                            let key_for_disk = key.clone();
                            let bytes_for_disk = bytes.clone();
                            let max_mb = self.session.config.cache_max_mb;
                            tokio::task::spawn_blocking(move || {
                                let _ = crate::teams_cache::store_image_bytes(
                                    &root,
                                    &key_for_disk,
                                    &bytes_for_disk,
                                );
                                let _ = prune_persistent_cache(&root, max_mb * 1024 * 1024);
                            });
                        }
                        self.image_cache.insert(key, img);
                    }
                    None => {
                        self.image_failed.insert(key);
                    }
                }
            }
            AppMessage::HostedImageFailed { key } => {
                self.image_loading.remove(&key);
                self.image_failed.insert(key);
            }
            AppMessage::OpenMailDue { generation } => {
                if generation == self.outlook.mail_fetch_generation {
                    self.open_current_mail_now();
                }
            }
        }
    }

    fn graph_message_time(value: Option<&str>) -> Option<chrono::DateTime<chrono::Utc>> {
        chrono::DateTime::parse_from_rfc3339(value?)
            .ok()
            .map(|value| value.with_timezone(&chrono::Utc))
    }

    fn activity_instant_from_graph_time(
        value: Option<chrono::DateTime<chrono::Utc>>,
        now: std::time::Instant,
    ) -> std::time::Instant {
        let age_seconds = value
            .map(|value| {
                chrono::Utc::now()
                    .signed_duration_since(value)
                    .num_seconds()
                    .max(0) as u64
            })
            .unwrap_or(TEAMS_COOL_AGE_SECONDS + 1)
            .min(TEAMS_COOL_AGE_SECONDS + 1);

        now.checked_sub(std::time::Duration::from_secs(age_seconds))
            .unwrap_or(now)
    }

    fn state_message_is_newer(
        state: &TeamsHotChatState,
        message_id: &str,
        message_at: Option<chrono::DateTime<chrono::Utc>>,
    ) -> bool {
        if state.latest_message_id.as_deref() == Some(message_id) {
            return false;
        }

        match (message_at.as_ref(), state.latest_message_at.as_ref()) {
            (Some(new), Some(old)) if new > old => true,
            (Some(new), Some(old)) if new < old => false,
            (Some(_), Some(_)) => match state.latest_message_id.as_deref() {
                Some(old_id) => teams_message_id_is_newer(message_id, old_id),
                None => true,
            },
            (Some(_), None) => true,
            (None, Some(_)) => false,
            (None, None) => true,
        }
    }

    fn sync_teams_hot_activity_from_chats(&mut self, chats: &[Chat]) {
        let now = std::time::Instant::now();
        let present: std::collections::HashSet<&str> =
            chats.iter().map(|chat| chat.id.as_str()).collect();

        for chat in chats {
            let Some(preview) = chat.last_message_preview.as_ref() else {
                continue;
            };
            let Some(message_id) = preview.id.as_deref() else {
                continue;
            };
            let message_at = Self::graph_message_time(preview.created_date_time.as_deref());

            match self.teams_hot_chat_state.get_mut(&chat.id) {
                Some(state) if Self::state_message_is_newer(state, message_id, message_at) => {
                    if preview
                        .from
                        .as_ref()
                        .is_some_and(|from| from.application.is_some())
                    {
                        tracing::warn!(
                            "TEAMS_APP_DIAG source=chat-list-hot chat_id={} chat_type={} preview_id={} previous_id={} preview_at={}",
                            chat.id,
                            chat.chat_type.as_deref().unwrap_or("<none>"),
                            message_id,
                            state.latest_message_id.as_deref().unwrap_or("<none>"),
                            preview.created_date_time.as_deref().unwrap_or("<none>")
                        );
                    }
                    state.latest_message_id = Some(message_id.to_string());
                    state.latest_message_at = message_at;
                    state.last_activity = now;
                    state.last_poll_started = Some(now);
                }
                Some(_) => {}
                None => {
                    self.teams_hot_chat_state.insert(
                        chat.id.clone(),
                        TeamsHotChatState {
                            last_activity: Self::activity_instant_from_graph_time(message_at, now),
                            // Do not burst immediately at startup; a recent chat
                            // becomes due after its normal tier interval.
                            last_poll_started: Some(now),
                            latest_message_id: Some(message_id.to_string()),
                            latest_message_at: message_at,
                            hot_pending: false,
                            foreground_in_flight: false,
                        },
                    );
                }
            }
        }

        self.teams_hot_chat_state.retain(|id, state| {
            present.contains(id.as_str()) || state.hot_pending || state.foreground_in_flight
        });
    }

    fn observe_teams_chat_page(&mut self, chat_id: &str, messages: &[ChatMessage]) -> bool {
        let Some(newest) = messages.iter().max_by_key(|message| sort_key(message)) else {
            return false;
        };
        let now = std::time::Instant::now();
        let message_at = Self::graph_message_time(newest.created_date_time.as_deref());

        match self.teams_hot_chat_state.get_mut(chat_id) {
            Some(state) if Self::state_message_is_newer(state, &newest.id, message_at) => {
                state.latest_message_id = Some(newest.id.clone());
                state.latest_message_at = message_at;
                state.last_activity = now;
                // This page itself was a message GET, so the next accelerated
                // poll starts from now rather than trying to catch up.
                state.last_poll_started = Some(now);
                true
            }
            Some(_) => false,
            None => {
                self.teams_hot_chat_state.insert(
                    chat_id.to_string(),
                    TeamsHotChatState {
                        last_activity: Self::activity_instant_from_graph_time(message_at, now),
                        last_poll_started: Some(now),
                        latest_message_id: Some(newest.id.clone()),
                        latest_message_at: message_at,
                        hot_pending: false,
                        foreground_in_flight: false,
                    },
                );
                false
            }
        }
    }

    fn hot_page_differs_from_loaded_messages(&self, messages: &[ChatMessage]) -> bool {
        if messages.is_empty() {
            return false;
        }

        let loaded: std::collections::HashMap<&str, &ChatMessage> = self
            .teams
            .messages
            .iter()
            .map(|message| (message.id.as_str(), message))
            .collect();

        let current_page: Option<Vec<&ChatMessage>> = messages
            .iter()
            .map(|incoming| loaded.get(incoming.id.as_str()).copied())
            .collect();
        let Some(current_page) = current_page else {
            return true;
        };

        match (
            serde_json::to_vec(&current_page),
            serde_json::to_vec(messages),
        ) {
            (Ok(current), Ok(incoming)) => current != incoming,
            _ => true,
        }
    }

    fn begin_teams_foreground_request(&mut self, chat_id: &str) -> bool {
        self.ensure_teams_chat_poll_state(chat_id);
        let Some(state) = self.teams_hot_chat_state.get_mut(chat_id) else {
            return false;
        };
        if state.foreground_in_flight {
            return false;
        }
        state.foreground_in_flight = true;
        state.last_poll_started = Some(std::time::Instant::now());
        true
    }

    fn ensure_teams_chat_poll_state(&mut self, chat_id: &str) {
        if self.teams_hot_chat_state.contains_key(chat_id) {
            return;
        }

        let now = std::time::Instant::now();
        let preview = self
            .teams
            .chats
            .iter()
            .find(|chat| chat.id == chat_id)
            .and_then(|chat| chat.last_message_preview.as_ref());
        let latest_message_id = preview.and_then(|preview| preview.id.clone());
        let latest_message_at = preview
            .and_then(|preview| Self::graph_message_time(preview.created_date_time.as_deref()));
        let last_activity = Self::activity_instant_from_graph_time(latest_message_at, now);

        self.teams_hot_chat_state.insert(
            chat_id.to_string(),
            TeamsHotChatState {
                last_activity,
                last_poll_started: None,
                latest_message_id,
                latest_message_at,
                hot_pending: false,
                foreground_in_flight: false,
            },
        );
    }

    fn mark_teams_chat_active(&mut self, chat_id: &str) {
        let now = std::time::Instant::now();
        if let Some(state) = self.teams_hot_chat_state.get_mut(chat_id) {
            state.last_activity = now;
            state.last_poll_started = Some(now);
            return;
        }

        self.teams_hot_chat_state.insert(
            chat_id.to_string(),
            TeamsHotChatState {
                last_activity: now,
                last_poll_started: Some(now),
                latest_message_id: None,
                latest_message_at: None,
                hot_pending: false,
                foreground_in_flight: false,
            },
        );
    }

    fn teams_chat_poll_tier_at(&self, chat_id: &str, now: std::time::Instant) -> ChatPollTier {
        let Some(state) = self.teams_hot_chat_state.get(chat_id) else {
            return ChatPollTier::Normal;
        };
        let tier = chat_poll_tier_from_age(now.duration_since(state.last_activity));
        if tier == ChatPollTier::Normal {
            return tier;
        }

        // HOT + WARM + COOL share the same top-N set. Tie-break by chat id so
        // selection remains deterministic when two activities land together.
        let ahead = self
            .teams_hot_chat_state
            .iter()
            .filter(|(other_id, other)| {
                chat_poll_tier_from_age(now.duration_since(other.last_activity))
                    != ChatPollTier::Normal
                    && (other.last_activity > state.last_activity
                        || (other.last_activity == state.last_activity
                            && other_id.as_str() < chat_id))
            })
            .count();

        if ahead < self.session.config.teams_hot_chats {
            tier
        } else {
            ChatPollTier::Normal
        }
    }

    pub fn teams_chat_poll_tier(&self, chat_id: &str) -> ChatPollTier {
        self.teams_chat_poll_tier_at(chat_id, std::time::Instant::now())
    }

    pub fn teams_poll_diagnostics(&self) -> TeamsPollDiagnostics {
        let now = std::time::Instant::now();
        let mut out = TeamsPollDiagnostics::default();

        for chat in &self.teams.chats {
            match self.teams_chat_poll_tier_at(&chat.id, now) {
                ChatPollTier::Hot => out.hot += 1,
                ChatPollTier::Warm => out.warm += 1,
                ChatPollTier::Cool => out.cool += 1,
                ChatPollTier::Normal => out.normal += 1,
            }
        }

        for state in self.teams_hot_chat_state.values() {
            if state.hot_pending {
                out.hot_pending += 1;
            }
            if state.foreground_in_flight {
                out.foreground_in_flight += 1;
            }
        }

        let (runtime_budget_rps, reserved_chats) =
            self.teams_background_limiter.diagnostics_snapshot();
        out.runtime_budget_rps = runtime_budget_rps;
        out.reserved_chats = reserved_chats;
        out
    }

    fn poll_hot_chats_if_due(&mut self) {
        let now = std::time::Instant::now();

        // A slow/failed normal Graph poll wins. Teams accelerated polling is
        // opportunistic background work and must yield to Mail/Calendar health.
        if self.poll_health_pending || self.poll_backoff_until.is_some_and(|until| now < until) {
            return;
        }

        let mut due: Vec<(std::time::Instant, String)> = self
            .teams_hot_chat_state
            .iter()
            .filter_map(|(chat_id, state)| {
                if state.hot_pending || state.foreground_in_flight {
                    return None;
                }

                let tier = self.teams_chat_poll_tier_at(chat_id, now);
                let interval = tier.interval()?;
                let due_at = state
                    .last_poll_started
                    .map(|started| started + interval)
                    .unwrap_or(now);

                (now >= due_at).then(|| (due_at, chat_id.clone()))
            })
            .collect();

        // Oldest due first. Each chat gets at most one pending hot task; the
        // shared limiter controls actual HTTP concurrency and per-chat spacing.
        due.sort_by_key(|(due_at, _)| *due_at);

        for (_, chat_id) in due {
            if let Some(state) = self.teams_hot_chat_state.get_mut(&chat_id) {
                state.hot_pending = true;
                state.last_poll_started = Some(now);
            }
            self.spawn_hot_chat_poll(chat_id, now);
        }
    }

    fn spawn_hot_chat_poll(&self, chat_id: String, scheduled_at: std::time::Instant) {
        let s = self.session.clone();
        let tx = self.tx.clone();
        let limiter = self.teams_background_limiter.clone();

        tokio::spawn(async move {
            let Some(permit) = limiter.acquire_hot(&s.graph, &chat_id, scheduled_at).await else {
                let _ = tx.send(AppMessage::HotChatPollSkipped { chat_id }).await;
                return;
            };
            let result = chats::list_messages(&s.graph, &chat_id, PAGE_SIZE).await;
            drop(permit);

            let message = match result {
                Ok((messages, next)) => AppMessage::HotChatPollFinished {
                    chat_id,
                    messages,
                    next,
                    error: None,
                },
                Err(error) => AppMessage::HotChatPollFinished {
                    chat_id,
                    messages: Vec::new(),
                    next: None,
                    error: Some(format!("{error:#}")),
                },
            };
            let _ = tx.send(message).await;
        });
    }

    fn persist_teams_page_cache(&self, chat_id: &str, messages: &[ChatMessage]) {
        let Some(cache_root) = self.session.config.cache_dir.clone() else {
            return;
        };
        if messages.is_empty() {
            return;
        }

        let chat_id = chat_id.to_string();
        let fresh = messages.to_vec();
        let cache_max_bytes = self.session.config.cache_max_mb.saturating_mul(1024 * 1024);

        tokio::spawn(async move {
            let cache_chat_id = chat_id.clone();
            match tokio::task::spawn_blocking(move || {
                // The freshly fetched newest page is authoritative for its
                // covered time range. Keep genuinely older cached history, but
                // do not resurrect entries that disappeared from Graph inside
                // the fresh window. Deleted tombstones are not persisted.
                let cached = load_teams_conversation_cache(&cache_root, &cache_chat_id)?
                    .map(|(messages, _)| messages)
                    .unwrap_or_default();
                let mut oldest_first = reconcile_authoritative_message_window(fresh, cached);
                if oldest_first.len() > MAX_TEAMS_CACHED_MESSAGES {
                    let drop = oldest_first.len() - MAX_TEAMS_CACHED_MESSAGES;
                    oldest_first.drain(..drop);
                }
                store_teams_conversation_cache(
                    &cache_root,
                    &cache_chat_id,
                    &oldest_first,
                    cache_max_bytes,
                )
            })
            .await
            {
                Ok(Ok(())) => {
                    tracing::debug!("Teams hot-poll cache stored for {chat_id}");
                }
                Ok(Err(error)) => {
                    tracing::warn!("Teams hot-poll cache write failed for {chat_id}: {error}");
                }
                Err(error) => {
                    tracing::warn!(
                        "Teams hot-poll cache writer task failed for {chat_id}: {error}"
                    );
                }
            }
        });
    }

    fn hot_poll_new_messages<'a>(
        messages: &'a [ChatMessage],
        previous_message_id: Option<&str>,
        newest_changed: bool,
    ) -> Vec<&'a ChatMessage> {
        if !newest_changed || messages.is_empty() {
            return Vec::new();
        }

        // Graph does not guarantee the order of this endpoint strongly enough
        // for notification sequencing. Reuse the conversation sort key so
        // multiple arrivals are announced oldest -> newest.
        let mut ordered: Vec<&ChatMessage> = messages.iter().collect();
        ordered.sort_by_key(|message| sort_key(message));

        if let Some(previous_message_id) = previous_message_id {
            if let Some(previous_index) = ordered
                .iter()
                .position(|message| message.id == previous_message_id)
            {
                return ordered.into_iter().skip(previous_index + 1).collect();
            }
        }

        // The baseline can fall out of PAGE_SIZE after a burst, or be absent
        // when the hot state was created without a preview. Never replay a
        // whole page of history in that case: keep the old safe behaviour and
        // announce only the newest message.
        ordered.into_iter().rev().take(1).collect()
    }

    fn notify_for_hot_chat_message(&mut self, chat_id: &str, message: &ChatMessage) {
        if let Some(seen) = self.chat_seen.as_mut() {
            seen.insert(chat_id.to_string(), message.id.clone());
        }

        let my_id = self.me.as_ref().map(|me| me.id.clone());
        let my_name = self.me.as_ref().and_then(|me| me.display_name.clone());
        let diag_application_sender = message
            .from
            .as_ref()
            .is_some_and(|from| from.application.is_some());
        if diag_application_sender {
            tracing::warn!(
                "TEAMS_APP_DIAG source=hot-message-enter chat_id={} message_id={} incoming={} deleted={} system={}",
                chat_id,
                message.id,
                teams_message_is_incoming(message, my_id.as_deref()),
                message.deleted_date_time.is_some(),
                message.is_system_event()
            );
        }
        let automation_sender =
            self.teams_message_is_automation_incoming(chat_id, message, my_id.as_deref());
        if !automation_sender && !teams_message_is_incoming(message, my_id.as_deref()) {
            return;
        }
        let application_sender = message
            .from
            .as_ref()
            .is_some_and(|from| from.application.is_some());

        let Some(chat) = self.teams.chats.iter().find(|chat| chat.id == chat_id) else {
            return;
        };
        let one_on_one = chat
            .chat_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"));
        let chat_type = chat.chat_type.clone();
        let fallback_label = teams_notification_label(
            self.teams.contact_names.get(chat_id).map(String::as_str),
            chat,
            my_id.as_deref(),
        );
        let open = self.teams.open_chat_id.as_deref() == Some(chat_id);

        if application_sender {
            tracing::warn!(
                "TEAMS_APP_DIAG source=hot-message-chat chat_id={} message_id={} chat_type={} one_on_one={} open={} notifications_enabled={}",
                chat_id,
                message.id,
                chat_type.as_deref().unwrap_or("<none>"),
                one_on_one,
                open,
                self.notification_events_enabled()
            );
        }

        if !open {
            if one_on_one {
                let count = self
                    .teams
                    .chat_unread_counts
                    .entry(chat_id.to_string())
                    .or_insert(0);
                *count = count.saturating_add(1).min(100);
            }
            if self.screen != Screen::Teams {
                self.teams_unread = true;
            }
        }

        if !self.notification_events_enabled() || !self.notified.insert(message.id.clone()) {
            return;
        }

        let body = message
            .body
            .as_ref()
            .and_then(|body| body.content.clone())
            .unwrap_or_default();
        let mentioned = crate::notify::mentions_me(
            &message.mentions,
            &body,
            my_id.as_deref(),
            my_name.as_deref(),
        );
        let should_notify = crate::notify::should_notify(chat_type.as_deref(), mentioned);
        if application_sender {
            tracing::warn!(
                "TEAMS_APP_DIAG source=hot-message-policy chat_id={} message_id={} mentioned={} should_notify={} unread_count={}",
                chat_id,
                message.id,
                mentioned,
                should_notify,
                self.teams
                    .chat_unread_counts
                    .get(chat_id)
                    .copied()
                    .unwrap_or(0)
            );
        }
        if !should_notify {
            return;
        }

        let who = if automation_sender {
            fallback_label.clone()
        } else {
            message
                .from
                .as_ref()
                .and_then(|from| from.user.as_ref().or(from.application.as_ref()))
                .and_then(|sender| sender.display_name.clone())
                .unwrap_or(fallback_label)
        };
        let title = if mentioned {
            format!("{who} mentioned you")
        } else {
            who
        };
        let plain = content::plain(&content::render_body(None, &body).text);
        let priority = if (application_sender || automation_sender)
            && self.session.config.teams_application_high_priority
        {
            NotificationPriority::High
        } else {
            NotificationPriority::Normal
        };
        self.send_notification("Teams", title, plain, priority);

        if self.notified.len() > 1000 {
            self.notified.clear();
        }
    }

    fn apply_hot_chat_poll(
        &mut self,
        chat_id: String,
        messages: Vec<ChatMessage>,
        next: Option<String>,
        error: Option<String>,
    ) {
        if let Some(state) = self.teams_hot_chat_state.get_mut(&chat_id) {
            state.hot_pending = false;
            // Never catch up missed intervals after a queued/slow request.
            state.last_poll_started = Some(std::time::Instant::now());
        }

        if let Some(error) = error {
            tracing::debug!("Teams hot poll failed for {chat_id}: {error}");
            return;
        }

        let previous_message_id = self
            .teams_hot_chat_state
            .get(&chat_id)
            .and_then(|state| state.latest_message_id.clone());
        let newest = messages
            .iter()
            .max_by_key(|message| sort_key(message))
            .cloned();
        let changed = self.observe_teams_chat_page(&chat_id, &messages);

        if newest
            .as_ref()
            .and_then(|message| message.from.as_ref())
            .is_some_and(|from| from.application.is_some())
        {
            tracing::warn!(
                "TEAMS_APP_DIAG source=hot-page chat_id={} previous_id={} newest_id={} changed={} message_count={}",
                chat_id,
                previous_message_id.as_deref().unwrap_or("<none>"),
                newest.as_ref().map(|message| message.id.as_str()).unwrap_or("<none>"),
                changed,
                messages.len()
            );
        }

        for message in
            Self::hot_poll_new_messages(&messages, previous_message_id.as_deref(), changed)
        {
            self.notify_for_hot_chat_message(&chat_id, message);
        }

        let previewing = self.screen == Screen::Teams
            && self.teams.mode == TeamsMode::Chats
            && self.teams.focus == TeamsFocus::List
            && self.teams.open_chat_id.is_none()
            && self.teams.preview_chat_id.as_deref() == Some(&chat_id);
        let open = self.teams.open_chat_id.as_deref() == Some(&chat_id);
        let page_content_changed =
            (open || previewing) && self.hot_page_differs_from_loaded_messages(&messages);
        let view_latest_id = self
            .teams
            .messages
            .last()
            .map(|message| message.id.as_str());
        let page_latest_id = newest.as_ref().map(|message| message.id.as_str());
        let view_needs_refresh = (open || previewing)
            && (self.teams.messages.is_empty() || view_latest_id != page_latest_id);

        if open || previewing {
            if changed || view_needs_refresh || page_content_changed {
                let learned_peer_id = self.remember_contact_from_messages(&chat_id, &messages);
                if learned_peer_id {
                    self.load_contact_presences(&[]);
                }
                self.set_teams_messages(messages, next, ListUpdate::Merge);
                self.persist_teams_conversation_cache(&chat_id);
            }

            if open && self.teams.chat_open_pending_read.as_deref() == Some(&chat_id) {
                self.teams.chat_open_pending_read = None;
                self.teams.chat_unread_counts.remove(&chat_id);
                self.mark_teams_chat_read(chat_id.clone());
                if let Some(latest_id) = page_latest_id {
                    self.teams
                        .locally_read_through
                        .insert(chat_id.clone(), latest_id.to_string());
                }
            }
        } else if changed {
            self.persist_teams_page_cache(&chat_id, &messages);
        }
    }

    pub fn poll_elapsed(&self) -> std::time::Duration {
        self.poll_last_success_at.elapsed()
    }

    pub fn kitty_images_available(&self) -> bool {
        self.image_picker.is_some()
    }

    fn mark_poll_success(&mut self) {
        self.poll_last_success_at = std::time::Instant::now();
        self.poll_health_pending = false;
        self.poll_retry_interval_secs = POLL_SECONDS;
        self.poll_backoff_until = None;
    }

    fn scheduled_poll(&mut self) {
        let now = std::time::Instant::now();

        if self.poll_health_pending {
            self.poll_health_pending = false;
            self.poll_retry_interval_secs = next_poll_retry_interval(self.poll_retry_interval_secs);
            let remaining_backoff = self.poll_retry_interval_secs.saturating_sub(POLL_SECONDS);
            self.poll_backoff_until = Some(now + std::time::Duration::from_secs(remaining_backoff));
            tracing::warn!(
                "poll health response is late; Graph poll interval backed off to {}s",
                self.poll_retry_interval_secs
            );
        }

        if self.poll_backoff_until.is_some_and(|until| now < until) {
            return;
        }

        self.poll_backoff_until = None;
        self.poll_health_pending = true;
        self.poll();
    }

    fn force_poll(&mut self) {
        self.poll_health_pending = true;
        self.poll_retry_interval_secs = POLL_SECONDS;
        self.poll_backoff_until = None;
        let accelerated_open_chat = self
            .teams
            .open_chat_id
            .clone()
            .filter(|id| self.teams_chat_poll_tier(id) != ChatPollTier::Normal);
        self.poll();
        if let Some(id) = accelerated_open_chat {
            // Manual F5 is foreground work. It bypasses the background token
            // budget but still shares the per-chat in-flight coalescing guard.
            self.load_chat_messages(id, ListUpdate::Merge);
        }
    }

    fn poll(&mut self) {
        self.renew_presence_session();
        if self.session.config.ntfy.needs_presence() {
            self.refresh_presence_for_ntfy();
        }
        self.refresh_ntfy_work_plan();
        self.peek_inbox();
        // Mail: refresh folder unread counts + the open folder's messages.
        self.load_folders();
        if let Some(f) = self.outlook.folders.get(self.outlook.folder_sel) {
            self.load_messages(f.id.clone(), ListUpdate::Merge);
        }
        // Teams: refresh the chat list. An accelerated open chat is already
        // covered by the hot scheduler, so do not add a duplicate message GET
        // every 20 seconds.
        self.load_chats();
        if let Some(id) = self.teams.open_chat_id.clone() {
            if self.teams_chat_poll_tier(&id) == ChatPollTier::Normal {
                self.load_chat_messages(id, ListUpdate::Merge);
            }
        }
        if let Some((t, c)) = self.teams.open_channel.clone() {
            self.load_channel_messages(t, c, ListUpdate::Merge);
        }
        if self.screen == Screen::Calendar
            || matches!(self.overlay.as_ref(), Some(Overlay::Calendar))
        {
            self.load_calendar();
        }
    }

    pub fn on_change(&mut self, change: ChangeEvent) {
        match change.kind() {
            ChangeKind::Mail => {
                self.status = "📬 new mail".into();
                self.peek_inbox();
                if let Some(f) = self.outlook.folders.get(self.outlook.folder_sel) {
                    self.load_messages(f.id.clone(), ListUpdate::Merge);
                }
            }
            ChangeKind::Chat => {
                self.status = "💬 chat update".into();
                self.load_chats();
                if let Some(id) = self.teams.open_chat_id.clone() {
                    self.load_chat_messages(id, ListUpdate::Merge);
                }
            }
            ChangeKind::Channel => {
                if let Some((t, c)) = self.teams.open_channel.clone() {
                    self.load_channel_messages(t, c, ListUpdate::Merge);
                }
            }
            ChangeKind::Other => {}
        }
    }

    pub fn teams_message_visible(&self, message: &ChatMessage) -> bool {
        if !message.is_system_event() {
            return true;
        }

        let class = message
            .system_event_display()
            .map(|display| display.class)
            .unwrap_or(SystemEventClass::Unknown);

        match self.session.config.teams_system_events {
            TeamsSystemEvents::None => false,
            TeamsSystemEvents::All => true,
            TeamsSystemEvents::Useful => class != SystemEventClass::Noise,
        }
    }

    fn teams_message_selectable(&self, message: &ChatMessage) -> bool {
        message.deleted_date_time.is_none() && self.teams_message_visible(message)
    }

    fn first_selectable_teams_message(&self) -> Option<usize> {
        self.teams
            .messages
            .iter()
            .position(|message| self.teams_message_selectable(message))
    }

    fn last_selectable_teams_message(&self) -> Option<usize> {
        self.teams
            .messages
            .iter()
            .rposition(|message| self.teams_message_selectable(message))
    }

    fn nearest_selectable_teams_message(&self, preferred: usize) -> Option<usize> {
        if self
            .teams
            .messages
            .get(preferred)
            .is_some_and(|message| self.teams_message_selectable(message))
        {
            return Some(preferred);
        }

        let before = (0..preferred.min(self.teams.messages.len()))
            .rev()
            .find(|index| self.teams_message_selectable(&self.teams.messages[*index]));
        before.or_else(|| {
            ((preferred + 1).min(self.teams.messages.len())..self.teams.messages.len())
                .find(|index| self.teams_message_selectable(&self.teams.messages[*index]))
        })
    }

    fn set_teams_messages(
        &mut self,
        messages: Vec<ChatMessage>,
        next: Option<String>,
        mode: ListUpdate,
    ) {
        self.teams.loading_more = false;
        let mut follow_newest = false;
        // Remember what was selected so a refresh doesn't move the cursor to a
        // different message when others arrive.
        let selected_id = self
            .teams
            .messages
            .get(self.teams.msg_sel)
            .map(|m| m.id.clone());

        // Graph answers newest-first; the conversation reads oldest-first, so
        // the newest message sits at the bottom next to the composer.
        let mut page = messages;
        page.reverse();

        match mode {
            ListUpdate::Replace => {
                self.teams.messages = page;
                self.teams.messages_next = next;
            }
            ListUpdate::Append => {
                // An older page belongs *before* everything already loaded.
                // Cache-restored ranges can overlap a fresh Graph page, so drop
                // duplicate ids before prepending it.
                let existing = std::mem::take(&mut self.teams.messages);
                let known: std::collections::HashSet<String> =
                    existing.iter().map(|message| message.id.clone()).collect();
                page.retain(|message| !known.contains(&message.id));
                self.teams.messages = page;
                self.teams.messages.extend(existing);
                self.teams.messages_next = next;
                self.status = if self.teams.messages_next.is_some() {
                    format!("{} messages loaded (more above)", self.teams.messages.len())
                } else {
                    format!(
                        "{} messages loaded (start of conversation)",
                        self.teams.messages.len()
                    )
                };
            }
            ListUpdate::Merge => {
                if self.teams.messages_next.is_none() {
                    self.teams.messages_next = next;
                }

                // Chat convention: if the user is sitting on the newest message,
                // follow new arrivals; if they've scrolled back to read history,
                // hold their place and just count what came in.
                follow_newest = self.last_selectable_teams_message() == Some(self.teams.msg_sel);
                let known: std::collections::HashSet<&str> =
                    self.teams.messages.iter().map(|m| m.id.as_str()).collect();
                let arrived = page
                    .iter()
                    .filter(|m| !known.contains(m.id.as_str()) && self.teams_message_selectable(m))
                    .count();
                if !follow_newest {
                    self.teams.unseen += arrived;
                }

                // The fetched page is authoritative for its newest covered
                // time window. Keep only cached history that is genuinely older
                // than that boundary; entries missing from Graph inside the
                // window are stale and must disappear. Fresh deleted tombstones
                // remove their cached originals and are not rendered/persisted.
                let existing = std::mem::take(&mut self.teams.messages);
                self.teams.messages = reconcile_authoritative_message_window(page, existing);
            }
        }

        // Order the list ourselves rather than trusting the order pages arrive
        // in: Graph's default ordering is not strictly by creation time, and
        // merging pages can interleave. Sorting by timestamp (id breaks ties —
        // Teams ids are epoch milliseconds) is cheap and always correct.
        self.teams.messages.sort_by_key(sort_key);

        // Re-render bodies for whatever the list now holds (HTML parse is cached
        // per message, so this only costs on changed content).
        let rendered: Vec<_> = self
            .teams
            .messages
            .iter()
            .map(|m| {
                let ct = m.body.as_ref().and_then(|b| b.content_type.as_deref());
                content::render_body(ct, &m.text())
            })
            .collect();
        // Shared files are openable too: fold their URLs into the link list.
        self.teams.messages_links = self
            .teams
            .messages
            .iter()
            .zip(rendered.iter())
            .map(|(m, r)| {
                let mut links = r.links.clone();
                for att in &m.attachments {
                    if let Some(url) = att.content_url.as_ref().filter(|u| !u.is_empty()) {
                        if !links.contains(url) {
                            links.push(url.clone());
                        }
                    }
                }
                if let Some(url) = m
                    .system_event_display()
                    .and_then(|display| display.url)
                    .filter(|url| !url.is_empty())
                {
                    if !links.contains(&url) {
                        links.push(url);
                    }
                }
                links
            })
            .collect();
        self.teams.messages_images = rendered.iter().map(|r| r.images.clone()).collect();
        self.teams.messages_rendered = rendered.into_iter().map(|r| r.text).collect();
        self.fetch_message_images();
        self.fetch_user_photos();

        let raw_last = self.teams.messages.len().saturating_sub(1);
        let preferred = match mode {
            // Opening a conversation lands on the newest visible message.
            ListUpdate::Replace => self.last_selectable_teams_message().unwrap_or(0),
            _ if follow_newest => {
                self.teams.unseen = 0;
                self.last_selectable_teams_message().unwrap_or(0)
            }
            // Otherwise stay on the same raw message, wherever it moved to.
            _ => selected_id
                .and_then(|id| self.teams.messages.iter().position(|m| m.id == id))
                .unwrap_or(self.teams.msg_sel)
                .min(raw_last),
        };
        self.teams.msg_sel = self
            .nearest_selectable_teams_message(preferred)
            .unwrap_or(0);
        self.refresh_selected_teams_images();
    }

    fn refresh_current(&mut self) {
        match self.screen {
            Screen::Outlook => {
                if let Some(f) = self.outlook.folders.get(self.outlook.folder_sel) {
                    self.load_messages(f.id.clone(), ListUpdate::Merge);
                }
            }
            Screen::Teams => match self.teams.mode {
                TeamsMode::Chats => {
                    self.load_chats();
                    if let Some(id) = self.teams.open_chat_id.clone() {
                        self.load_chat_messages(id, ListUpdate::Merge);
                    }
                }
                TeamsMode::Channels => {
                    if let Some((t, c)) = self.teams.open_channel.clone() {
                        self.load_channel_messages(t, c, ListUpdate::Merge);
                    }
                }
            },
            Screen::Calendar => self.load_calendar(),
        }
    }

    pub fn focused_links(&self) -> &[String] {
        match self.screen {
            Screen::Outlook => &self.outlook.reading_links,
            Screen::Teams => self
                .teams
                .messages_links
                .get(self.teams.msg_sel)
                .map(|v| v.as_slice())
                .unwrap_or(&[]),
            Screen::Calendar => &[],
        }
    }

    fn open_web_link(&self, url: &str) -> anyhow::Result<()> {
        match self.session.config.web_link_opener.as_deref() {
            Some(executable) => crate::opener::open_url_with(url, executable),
            None => crate::opener::open_url(url),
        }
    }

    fn open_link(&mut self, index: usize) {
        let Some(url) = self.focused_links().get(index).cloned() else {
            return;
        };
        match self.open_web_link(&url) {
            Ok(()) => self.status = format!("opening {}", truncate_url(&url)),
            Err(e) => self.status = format!("could not open link: {e:#}"),
        }
    }

    pub fn ntfy_snooze_remaining(&self) -> Option<std::time::Duration> {
        ntfy_snooze_remaining_at(self.ntfy_snooze_until, std::time::SystemTime::now())
    }

    fn persist_ntfy_snooze(&self) {
        let result = match self.ntfy_snooze_until {
            Some(until) => store_ntfy_snooze_until(&self.session.config.token_cache_path, until),
            None => clear_ntfy_snooze_state(&self.session.config.token_cache_path),
        };

        if let Err(error) = result {
            tracing::warn!("could not persist ntfy snooze state: {error}");
        }
    }

    fn set_ntfy_snooze(&mut self, hours: u64) {
        let duration = std::time::Duration::from_secs(hours.saturating_mul(60 * 60));
        let Some(until) = std::time::SystemTime::now().checked_add(duration) else {
            self.status = "could not set ntfy snooze".into();
            return;
        };

        self.ntfy_snooze_until = Some(until);
        self.persist_ntfy_snooze();
        self.status = format!("ntfy snoozed for {hours}h");
    }

    fn clear_ntfy_snooze(&mut self) {
        self.ntfy_snooze_until = None;
        self.persist_ntfy_snooze();
        self.status = "ntfy snooze cleared".into();
    }

    fn expire_ntfy_snooze_if_due(&mut self) {
        if self.ntfy_snooze_until.is_some() && self.ntfy_snooze_remaining().is_none() {
            self.ntfy_snooze_until = None;
            self.persist_ntfy_snooze();
        }
    }

    fn notification_events_enabled(&self) -> bool {
        self.session.config.notifications || self.session.config.ntfy.enabled()
    }

    fn ntfy_should_forward(&self) -> bool {
        if self.ntfy_snooze_remaining().is_some() {
            return false;
        }

        ntfy_mode_should_forward(
            self.session.config.ntfy,
            self.my_presence.as_ref(),
            self.ntfy_working_now,
        )
    }

    fn send_notification(
        &self,
        category: &'static str,
        title: String,
        body: String,
        priority: NotificationPriority,
    ) {
        if self.session.config.notifications {
            crate::notify::send(&title, &body, priority);
        }

        if !self.ntfy_should_forward() {
            return;
        }

        let (Some(server), Some(topic)) = (
            self.session.config.ntfy_server.clone(),
            self.session.config.ntfy_topic.clone(),
        ) else {
            tracing::warn!("ntfy forwarding enabled without server/topic");
            return;
        };
        let token = self.session.config.ntfy_token.clone();
        let tag = ntfy_tag_for_category(category);

        tokio::spawn(async move {
            if let Err(error) = m365_core::ntfy::send(
                &server,
                &topic,
                token.as_deref(),
                tag,
                priority.ntfy_priority(),
                &title,
                &body,
            )
            .await
            {
                tracing::warn!("ntfy forwarding failed: {error:#}");
            }
        });
    }

    fn notify_for_chats(&mut self, chats: &[Chat]) {
        let notification_events_enabled = self.notification_events_enabled();
        let my_id = self.me.as_ref().map(|m| m.id.clone());
        let my_name = self.me.as_ref().and_then(|m| m.display_name.clone());

        // First sync only records the baseline; notifying here would announce
        // every conversation's history at startup.
        let Some(seen) = self.chat_seen.as_mut() else {
            self.chat_seen = Some(
                chats
                    .iter()
                    .filter_map(|c| {
                        let id = c.last_message_preview.as_ref()?.id.clone()?;
                        Some((c.id.clone(), id))
                    })
                    .collect(),
            );
            return;
        };

        let mut to_notify: Vec<(String, String, NotificationPriority)> = Vec::new();
        for chat in chats {
            let Some(preview) = chat.last_message_preview.as_ref() else {
                continue;
            };
            let Some(msg_id) = preview.id.clone() else {
                continue;
            };
            let previous = seen.insert(chat.id.clone(), msg_id.clone());
            if previous.as_deref() == Some(msg_id.as_str()) {
                continue; // nothing new here
            }

            let application_sender = preview
                .from
                .as_ref()
                .is_some_and(|from| from.application.is_some());

            if application_sender {
                tracing::warn!(
                    "TEAMS_APP_DIAG source=chat-list-change chat_id={} chat_type={} preview_id={} previous_id={} one_on_one={} open={} notifications_enabled={} high_priority={}",
                    chat.id,
                    chat.chat_type.as_deref().unwrap_or("<none>"),
                    msg_id,
                    previous.as_deref().unwrap_or("<none>"),
                    chat.chat_type
                        .as_deref()
                        .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne")),
                    self.teams.open_chat_id.as_deref() == Some(chat.id.as_str()),
                    notification_events_enabled,
                    self.session.config.teams_application_high_priority
                );
            }

            // Never notify for our own messages, or twice for the same one.
            let from_id = preview
                .from
                .as_ref()
                .and_then(|f| f.user.as_ref())
                .and_then(|u| u.id.as_deref());
            let automation_sender = self
                .teams
                .contact_user_ids
                .get(&chat.id)
                .is_some_and(|peer_id| self.teams_automation_peers.contains(peer_id))
                && !self.own_sent_chat_message_ids.contains(&msg_id)
                && from_id.is_some()
                && from_id == my_id.as_deref();
            if from_id.is_some() && from_id == my_id.as_deref() && !automation_sender {
                continue;
            }

            let one_on_one = chat
                .chat_type
                .as_deref()
                .is_some_and(|kind| kind.eq_ignore_ascii_case("oneOnOne"));
            let open = self.teams.open_chat_id.as_deref() == Some(chat.id.as_str());
            if (application_sender
                && teams_sender_is_incoming(preview.from.as_ref(), my_id.as_deref())
                || automation_sender)
                && one_on_one
                && !open
            {
                self.teams
                    .chat_unread_counts
                    .entry(chat.id.clone())
                    .and_modify(|count| *count = (*count).max(1))
                    .or_insert(1);
            }

            if application_sender {
                tracing::warn!(
                    "TEAMS_APP_DIAG source=chat-list-unread chat_id={} incoming={} unread_count={}",
                    chat.id,
                    teams_sender_is_incoming(preview.from.as_ref(), my_id.as_deref()),
                    self.teams
                        .chat_unread_counts
                        .get(&chat.id)
                        .copied()
                        .unwrap_or(0)
                );
            }

            if self.screen != Screen::Teams {
                self.teams_unread = true;
            }
            if !notification_events_enabled {
                continue;
            }
            if !self.notified.insert(msg_id.clone()) {
                continue;
            }

            let body = preview
                .body
                .as_ref()
                .and_then(|b| b.content.clone())
                .unwrap_or_default();
            let mentioned =
                crate::notify::mentions_me(&[], &body, my_id.as_deref(), my_name.as_deref());
            let should_notify = crate::notify::should_notify(chat.chat_type.as_deref(), mentioned);
            if application_sender {
                tracing::warn!(
                    "TEAMS_APP_DIAG source=chat-list-policy chat_id={} mentioned={} should_notify={} already_notified={}",
                    chat.id,
                    mentioned,
                    should_notify,
                    self.notified.contains(&msg_id)
                );
            }
            if !should_notify {
                continue;
            }

            let fallback_label = teams_notification_label(
                self.teams.contact_names.get(&chat.id).map(String::as_str),
                chat,
                my_id.as_deref(),
            );
            let who = if automation_sender {
                fallback_label.clone()
            } else {
                preview
                    .from
                    .as_ref()
                    .and_then(|from| from.user.as_ref().or(from.application.as_ref()))
                    .and_then(|sender| sender.display_name.clone())
                    .unwrap_or(fallback_label)
            };
            let title = if mentioned {
                format!("{who} mentioned you")
            } else {
                who
            };
            let priority = if (application_sender || automation_sender)
                && self.session.config.teams_application_high_priority
            {
                NotificationPriority::High
            } else {
                NotificationPriority::Normal
            };
            to_notify.push((
                title,
                content::plain(&content::render_body(None, &body).text),
                priority,
            ));
        }

        // Keep the id set from growing without bound over a long session.
        if self.notified.len() > 1000 {
            self.notified.clear();
        }
        for (title, body, priority) in to_notify {
            self.send_notification("Teams", title, body, priority);
        }
    }

    fn notify_for_mail(&mut self, items: &[MailMessage]) {
        if !self.notification_events_enabled() {
            return;
        }
        // First sync records a baseline; otherwise the whole inbox would
        // announce itself on startup.
        let Some(seen) = self.mail_seen.as_mut() else {
            self.mail_seen = Some(items.iter().map(|m| m.id.clone()).collect());
            return;
        };

        let mut to_notify: Vec<(String, String)> = Vec::new();
        for m in items {
            if !seen.insert(m.id.clone()) {
                continue; // already known
            }
            // Read elsewhere (phone, Outlook) before we got here.
            if m.is_read.unwrap_or(false) {
                continue;
            }
            if !self.notified.insert(m.id.clone()) {
                continue;
            }
            to_notify.push((
                m.sender_name(),
                m.subject.clone().unwrap_or_else(|| "(no subject)".into()),
            ));
        }

        if seen.len() > 1000 {
            seen.clear();
        }
        if self.notified.len() > 1000 {
            self.notified.clear();
        }
        for (who, subject) in to_notify {
            self.send_notification(
                "Mail",
                format!("✉ {who}"),
                subject,
                NotificationPriority::Normal,
            );
        }
    }

    fn automation_peer_config_value_with(&self, peer_id: &str) -> String {
        let mut peers: Vec<String> = self.teams_automation_peers.iter().cloned().collect();
        if !peers.iter().any(|item| item == peer_id) {
            peers.push(peer_id.to_string());
        }
        peers.sort();
        peers.dedup();
        peers.join(",")
    }

    fn copy_log_contact_automation_peer(&mut self) -> Option<(String, String)> {
        let Some(peer_id) = self.contact_diagnostics.automation_peer_id.clone() else {
            self.status = "F7 has no usable peer ID for this contact".into();
            return None;
        };

        let value = self.automation_peer_config_value_with(&peer_id);
        tracing::info!("F7 automation peer ID: {peer_id}");
        tracing::info!("F7 automation peer .env: M365_TEAMS_AUTOMATION_PEERS={value}");
        tracing::info!("F7 automation peer export: export M365_TEAMS_AUTOMATION_PEERS={value}");

        match crate::clipboard::copy(&peer_id) {
            Ok(via) => self.status = format!("copied automation peer ID via {via}"),
            Err(error) => {
                self.status = format!("automation peer ID logged; clipboard failed: {error:#}")
            }
        }

        Some((peer_id, value))
    }

    fn begin_add_contact_automation_peer(&mut self) {
        let Some((peer_id, _)) = self.copy_log_contact_automation_peer() else {
            return;
        };

        if self.teams_automation_peers.contains(&peer_id) {
            tracing::info!("F7 automation peer already configured: peer_id={peer_id}");
            self.status = format!("automation peer already configured: {peer_id}");
            self.contact_diagnostics.automation_peer_configured = true;
            return;
        }

        if self.session.config.teams_automation_peers_from_process_env {
            self.status =
                "M365_TEAMS_AUTOMATION_PEERS comes from process environment; use the logged export line"
                    .into();
            return;
        }

        let Some(path) = self.session.config.dotenv_path.as_deref() else {
            self.status =
                "no loaded .env file; peer ID and ready-to-use config lines were logged".into();
            return;
        };

        self.overlay = Some(Overlay::ConfirmAutomationPeer {
            peer_id,
            dotenv_path: path.display().to_string(),
        });
    }

    fn persist_contact_automation_peer(&mut self, peer_id: String, dotenv_path: String) {
        let value = self.automation_peer_config_value_with(&peer_id);
        match write_dotenv_assignment(
            std::path::Path::new(&dotenv_path),
            "M365_TEAMS_AUTOMATION_PEERS",
            &value,
        ) {
            Ok(()) => {
                self.teams_automation_peers.insert(peer_id.clone());
                if let Some(chat_id) = self.contact_diagnostics_chat_id.clone() {
                    self.teams.contact_user_ids.insert(chat_id, peer_id.clone());
                }
                self.contact_diagnostics.automation_peer_configured = true;
                self.contact_diagnostics.automation_config_source = dotenv_path.clone();
                tracing::info!("F7 automation peer added: peer_id={peer_id} config={dotenv_path}");
                tracing::info!("F7 automation peer .env: M365_TEAMS_AUTOMATION_PEERS={value}");
                tracing::info!(
                    "F7 automation peer export: export M365_TEAMS_AUTOMATION_PEERS={value}"
                );
                self.status =
                    format!("automation peer added and activated: {peer_id} ({dotenv_path})");
            }
            Err(error) => {
                tracing::warn!(
                    "F7 automation peer add failed: peer_id={peer_id} config={dotenv_path}: {error:#}"
                );
                self.status = format!("could not update {dotenv_path}: {error:#}");
            }
        }
    }

    fn yank_current(&mut self) {
        let text = match self.screen {
            Screen::Outlook => self.outlook.reading_body.as_ref().map(content::plain),
            Screen::Teams => self
                .teams
                .messages_rendered
                .get(self.teams.msg_sel)
                .map(content::plain),
            Screen::Calendar => self
                .calendar
                .events
                .get(self.calendar.selected)
                .and_then(|event| event.subject.clone()),
        };
        match text {
            Some(t) if !t.trim().is_empty() => self.copy_to_clipboard(&t, "message"),
            _ => self.status = "nothing to copy — open a message first".into(),
        }
    }

    fn yank_all(&mut self) {
        let text = match self.screen {
            Screen::Outlook => crate::ui::email_lines(self)
                .map(|lines| lines_to_plain(&lines))
                .unwrap_or_default(),
            Screen::Teams => lines_to_plain(&crate::ui::conversation_lines(self, false).0),
            Screen::Calendar => self
                .calendar
                .events
                .get(self.calendar.selected)
                .map(|event| event.subject.clone().unwrap_or_default())
                .unwrap_or_default(),
        };
        if text.trim().is_empty() {
            self.status = "nothing to copy".into();
        } else {
            self.copy_to_clipboard(&text, "all");
        }
    }

    fn copy_to_clipboard(&mut self, text: &str, what: &str) {
        match crate::clipboard::copy(text) {
            Ok(via) => {
                let bytes = text.len();
                self.status = format!("copied {what} to clipboard ({bytes} bytes, via {via})");
            }
            Err(e) => self.status = format!("copy failed: {e:#}"),
        }
    }

    fn export_diagnostics(&mut self, force_log: bool) {
        let text = crate::diagnostics::text(self);

        if !force_log {
            if let Some(via) = crate::clipboard::copy_native(&text) {
                self.status = format!("diagnostics copied to clipboard via {via}");
                return;
            }
        }

        match crate::diagnostics::save_log(&text) {
            Ok(path) => {
                self.status = if force_log {
                    format!("diagnostics saved to {}", path.display())
                } else {
                    format!(
                        "clipboard helper unavailable — diagnostics saved to {}",
                        path.display()
                    )
                };
            }
            Err(error) => {
                self.status = format!("could not export diagnostics: {error:#}");
            }
        }
    }

    fn export_contact_diagnostics(&mut self, force_log: bool) {
        let text = crate::diagnostics::contact_text(self);

        if !force_log {
            if let Some(via) = crate::clipboard::copy_native(&text) {
                self.status = format!("contact diagnostics copied to clipboard via {via}");
                return;
            }
        }

        match crate::diagnostics::save_contact_log(&text) {
            Ok(path) => {
                self.status = if force_log {
                    format!("contact diagnostics saved to {}", path.display())
                } else {
                    format!(
                        "clipboard helper unavailable — contact diagnostics saved to {}",
                        path.display()
                    )
                };
            }
            Err(error) => {
                self.status = format!("could not export contact diagnostics: {error:#}");
            }
        }
    }

    pub fn wants_mouse_capture(&self) -> bool {
        !self.copy_mode
            && self.overlay.is_none()
            && self.screen == Screen::Teams
            && self.teams.mode == TeamsMode::Chats
            && self.teams.focus == TeamsFocus::List
    }

    pub fn on_mouse(&mut self, mouse: MouseEvent) {
        if !self.wants_mouse_capture() {
            return;
        }

        let delta = match mouse.kind {
            MouseEventKind::ScrollUp => -1,
            MouseEventKind::ScrollDown => 1,
            _ => return,
        };

        self.note_presence_activity();
        self.teams_move(delta);
    }

    pub fn on_key(&mut self, key: KeyEvent) {
        if key.code == KeyCode::F(4) {
            if self.session.config.ntfy.enabled() {
                self.overlay = Some(Overlay::NtfySnooze { sel: 0 });
            }
            return;
        }
        if key.code == KeyCode::F(5) {
            self.force_poll();
            return;
        }
        // F6 is deliberately global, including while another overlay or
        // the Teams composer has focus.
        if key.code == KeyCode::F(6) {
            self.open_diagnostics();
            return;
        }
        if key.code == KeyCode::F(7) {
            self.open_contact_diagnostics();
            return;
        }
        if matches!(self.overlay.as_ref(), Some(Overlay::Notice(_))) {
            self.overlay = None;
            self.notice_until = None;
            return;
        }

        self.note_presence_activity();
        // Overlays capture input first.
        if self.overlay.is_some() {
            self.on_key_overlay(key);
            return;
        }

        // Copy mode swallows input so stray keys can't fire actions while the
        // user is selecting text.
        if self.copy_mode {
            match key.code {
                KeyCode::Char('z') | KeyCode::Esc | KeyCode::Char('q') => {
                    self.copy_mode = false;
                }
                KeyCode::Char('y') => self.yank_all(),
                KeyCode::Down | KeyCode::Char('j') => {
                    self.copy_scroll = self.copy_scroll.saturating_add(1)
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    self.copy_scroll = self.copy_scroll.saturating_sub(1)
                }
                KeyCode::PageDown => self.copy_scroll = self.copy_scroll.saturating_add(20),
                KeyCode::PageUp => self.copy_scroll = self.copy_scroll.saturating_sub(20),
                KeyCode::Home | KeyCode::Char('g')
                    if !key.modifiers.contains(KeyModifiers::SHIFT) =>
                {
                    self.copy_scroll = 0
                }
                KeyCode::End | KeyCode::Char('G') => self.copy_scroll = self.copy_max_scroll.get(),
                KeyCode::Char('g') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                    self.copy_scroll = self.copy_max_scroll.get()
                }
                KeyCode::Char('c') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    self.should_quit = true;
                }
                _ => {}
            }
            return;
        }

        // While typing a Teams message, only Ctrl-modified/function keys act
        // globally — otherwise letters like q/p/z would be stolen from the text.
        let typing = self.screen == Screen::Teams && self.teams.focus == TeamsFocus::Composer;

        // Global bindings.
        match (key.code, key.modifiers) {
            (KeyCode::Char('c'), KeyModifiers::CONTROL) => {
                self.should_quit = true;
                return;
            }
            (KeyCode::Char('q'), KeyModifiers::NONE) if !typing => {
                self.should_quit = true;
                return;
            }
            (KeyCode::Char('z'), KeyModifiers::NONE) if !typing => {
                self.copy_mode = true;
                self.copy_scroll = 0;
                return;
            }
            (KeyCode::Char('y'), KeyModifiers::NONE) if !typing => {
                self.yank_current();
                return;
            }
            (KeyCode::Char('Y'), _) if !typing => {
                self.yank_all();
                return;
            }
            (KeyCode::F(1), _) => {
                self.switch_screen(Screen::Outlook);
                return;
            }
            (KeyCode::F(2), _) => {
                self.switch_screen(Screen::Teams);
                return;
            }
            (KeyCode::F(3), _) => {
                self.switch_screen(Screen::Calendar);
                return;
            }
            (KeyCode::Char('|') | KeyCode::Char('\\'), _) if !typing => {
                self.toggle_pane_layout();
                return;
            }
            (KeyCode::Char('p'), KeyModifiers::CONTROL) => {
                self.overlay = Some(Overlay::Palette {
                    query: String::new(),
                    sel: 0,
                });
                return;
            }
            (KeyCode::Char('?'), _) if !typing => {
                self.help_scroll = 0;
                self.overlay = Some(Overlay::Help);
                return;
            }
            (KeyCode::Char('p'), KeyModifiers::NONE) if !typing => {
                self.overlay = Some(Overlay::Presence);
                return;
            }
            (KeyCode::Char('s'), KeyModifiers::NONE) if !typing => {
                self.overlay = Some(Overlay::Settings { sel: 0 });
                return;
            }
            (KeyCode::Char('A'), _) if !typing => {
                if self.screen == Screen::Outlook {
                    if self.outlook.reading_attachments.is_empty() {
                        self.status = "no attachments on this message".into();
                    } else {
                        self.overlay = Some(Overlay::Attachments);
                    }
                }
                return;
            }
            (KeyCode::Char('o'), KeyModifiers::NONE) if !typing => {
                if self.screen == Screen::Calendar {
                    self.open_selected_calendar_meeting();
                } else if self.focused_links().is_empty() {
                    self.status = "no links in this message".into();
                } else {
                    self.overlay = Some(Overlay::Links);
                }
                return;
            }
            _ => {}
        }

        match self.screen {
            Screen::Outlook => self.on_key_outlook(key),
            Screen::Teams => self.on_key_teams(key),
            Screen::Calendar => self.on_key_calendar(key),
        }

        if self.screen == Screen::Teams {
            self.refresh_selected_teams_images();
        }
    }

    fn switch_screen(&mut self, screen: Screen) {
        if self.screen == Screen::Outlook
            && screen != Screen::Outlook
            && self.outlook_focus == OutlookFocus::Reading
        {
            self.cancel_read_timer();
        }

        self.screen = screen;
        match screen {
            Screen::Outlook => {
                if self.outlook_focus == OutlookFocus::Reading {
                    self.schedule_current_read_timer();
                }
            }
            Screen::Teams => {
                self.teams_unread = false;
                if self.teams.chats.is_empty() {
                    self.load_chats();
                }

                if self.teams.mode == TeamsMode::Chats && self.teams.open_chat_id.is_none() {
                    if let Some(chat_id) = self.teams.last_chat_id.clone() {
                        self.teams.open_chat_id = Some(chat_id.clone());
                        self.teams.focus = TeamsFocus::Messages;
                        let cached = self.restore_teams_conversation_cache(&chat_id);
                        self.teams.chat_open_pending_read = Some(chat_id.clone());
                        self.load_chat_messages(
                            chat_id,
                            if cached {
                                ListUpdate::Merge
                            } else {
                                ListUpdate::Replace
                            },
                        );
                    }
                }
            }
            Screen::Calendar => self.load_calendar(),
        }
        self.persist_ui_state();
    }

    fn on_key_outlook(&mut self, key: KeyEvent) {
        match key.code {
            KeyCode::Char('H') => self.resize_folder_panel(-1),
            KeyCode::Char('L') => self.resize_folder_panel(1),
            KeyCode::Char('h') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.resize_folder_panel(-1)
            }
            KeyCode::Char('l') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.resize_folder_panel(1)
            }
            KeyCode::Tab => {
                if self.outlook_focus == OutlookFocus::Reading {
                    self.cancel_read_timer();
                }
                self.outlook_focus = match self.outlook_focus {
                    OutlookFocus::Folders => OutlookFocus::Messages,
                    OutlookFocus::Messages => OutlookFocus::Reading,
                    OutlookFocus::Reading => OutlookFocus::Folders,
                };
                if self.outlook_focus == OutlookFocus::Reading {
                    self.schedule_current_read_timer();
                }
            }

            KeyCode::Char('t') => self.toggle_mail_threads(),
            KeyCode::Char('c') => {
                self.overlay = Some(Overlay::Compose(empty_compose()));
            }
            KeyCode::Char('/') => {
                self.overlay = if self.outlook_focus == OutlookFocus::Folders {
                    Some(Overlay::FolderSearch {
                        query: String::new(),
                        sel: 0,
                    })
                } else {
                    Some(Overlay::Search {
                        query: String::new(),
                    })
                };
            }
            KeyCode::Char('r') => self.open_reply(ReplyMode::Reply),
            KeyCode::Char('a') => self.open_reply(ReplyMode::ReplyAll),
            KeyCode::Char('f') => self.open_reply(ReplyMode::Forward),
            KeyCode::Char('u') => self.toggle_mail_read(),
            KeyCode::Char('m') => self.open_move_mail(),
            KeyCode::Char('i') => self.open_mail_images(),
            KeyCode::Char('e') => self.load_calendar_and_show(),
            KeyCode::Char('d') | KeyCode::Delete => self.trash_current_mail(),
            KeyCode::Up | KeyCode::Char('k') => self.outlook_move(-1),
            KeyCode::Down | KeyCode::Char('j') => self.outlook_move(1),
            KeyCode::PageUp => self.outlook_move(-10),
            KeyCode::PageDown => self.outlook_move(10),
            // h/l move between panes: out to the left, into the thing on the
            // right. Esc is a synonym for backing out.
            KeyCode::Home | KeyCode::Char('g') if !key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.outlook_jump(false);
            }
            KeyCode::End | KeyCode::Char('G') => self.outlook_jump(true),
            KeyCode::Char('g') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.outlook_jump(true);
            }
            KeyCode::Esc | KeyCode::Left | KeyCode::Char('h') => self.outlook_out(),
            KeyCode::Right | KeyCode::Char('l') => self.outlook_into(),
            KeyCode::Enter => self.outlook_enter(),
            _ => {}
        }
    }

    fn load_calendar_and_show(&mut self) {
        self.load_calendar();
        self.overlay = Some(Overlay::Calendar);
    }

    fn on_key_calendar(&mut self, key: KeyEvent) {
        match key.code {
            KeyCode::Up | KeyCode::Char('k') => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_move_event(-1);
                } else {
                    self.calendar.selected =
                        step(self.calendar.selected, -1, self.calendar.events.len());
                }
            }
            KeyCode::Down | KeyCode::Char('j') => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_move_event(1);
                } else {
                    self.calendar.selected =
                        step(self.calendar.selected, 1, self.calendar.events.len());
                }
            }
            KeyCode::PageUp => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_move_event(-5);
                } else {
                    self.calendar.selected =
                        step(self.calendar.selected, -10, self.calendar.events.len());
                }
            }
            KeyCode::PageDown => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_move_event(5);
                } else {
                    self.calendar.selected =
                        step(self.calendar.selected, 10, self.calendar.events.len());
                }
            }
            KeyCode::Home => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_select_edge(false);
                } else {
                    self.calendar.selected = 0;
                }
            }
            KeyCode::End => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar_month_select_edge(true);
                } else {
                    self.calendar.selected = self.calendar.events.len().saturating_sub(1);
                }
            }
            KeyCode::Char('a') => self.respond_calendar(calendar::Rsvp::Accept),
            KeyCode::Char('d') => self.respond_calendar(calendar::Rsvp::Decline),
            KeyCode::Char('t') => self.respond_calendar(calendar::Rsvp::Tentative),
            KeyCode::Enter | KeyCode::Char('g') => {
                if self.calendar.events.is_empty() {
                    self.status = "no calendar event selected".into();
                } else {
                    self.overlay = Some(Overlay::CalendarEvent);
                }
            }
            KeyCode::Char('n') => {
                if self.calendar.view == CalendarView::Month {
                    self.calendar.month_offset = 0;
                    self.calendar_month_select_edge(false);
                    self.status = "calendar: current month".into();
                    self.load_calendar();
                } else {
                    self.calendar_jump_today();
                }
            }
            KeyCode::Char('v') => {
                if self.calendar.view == CalendarView::Agenda && !calendar_month_view_available() {
                    self.overlay = Some(Overlay::Notice(format!(
                        "Month view requires at least {}x{} characters.\n\nAgenda view remains active.",
                        CALENDAR_MONTH_MIN_WIDTH, CALENDAR_MONTH_MIN_HEIGHT
                    )));
                    self.notice_until =
                        Some(std::time::Instant::now() + std::time::Duration::from_secs(3));
                    return;
                }

                self.calendar.view = match self.calendar.view {
                    CalendarView::Agenda => CalendarView::Month,
                    CalendarView::Month => CalendarView::Agenda,
                };
                self.calendar.month_offset = 0;
                self.calendar.selected = 0;
                self.persist_ui_state();
                self.status = match self.calendar.view {
                    CalendarView::Agenda => "calendar view: agenda".into(),
                    CalendarView::Month => "calendar view: month".into(),
                };
                self.load_calendar();
            }
            KeyCode::Left if self.calendar.view == CalendarView::Month => {
                self.calendar.month_offset -= 1;
                self.calendar_month_select_edge(false);
                self.status = "calendar: previous month".into();
                self.load_calendar();
            }
            KeyCode::Right if self.calendar.view == CalendarView::Month => {
                self.calendar.month_offset += 1;
                self.calendar_month_select_edge(false);
                self.status = "calendar: next month".into();
                self.load_calendar();
            }
            KeyCode::Char('r') => self.load_calendar(),
            KeyCode::Char('w') if self.calendar.view == CalendarView::Agenda => {
                let next = CALENDAR_RANGES
                    .iter()
                    .position(|days| *days == self.calendar.days)
                    .map(|index| (index + 1) % CALENDAR_RANGES.len())
                    .unwrap_or(0);
                self.calendar.days = CALENDAR_RANGES[next];
                self.calendar.selected = 0;
                self.persist_ui_state();
                self.status = format!("calendar range: {} days", self.calendar.days);
                self.load_calendar();
            }
            _ => {}
        }
    }

    fn open_selected_calendar_meeting(&mut self) {
        let Some(url) = self
            .calendar
            .events
            .get(self.calendar.selected)
            .and_then(|event| event.online_meeting.as_ref())
            .and_then(|meeting| meeting.join_url.as_deref())
            .map(str::trim)
            .filter(|url| !url.is_empty())
            .map(str::to_string)
        else {
            return;
        };

        let result = self.open_web_link(&url);

        match result {
            Ok(()) => self.status = "opening online meeting".into(),
            Err(error) => self.status = format!("could not open meeting: {error:#}"),
        }
    }

    fn calendar_month_indices(&self) -> Vec<usize> {
        self.calendar
            .events
            .iter()
            .enumerate()
            .filter_map(|(index, event)| {
                calendar_event_intersects_month(event, self.calendar.month_offset).then_some(index)
            })
            .collect()
    }

    fn calendar_month_move_event(&mut self, delta: i32) {
        let indices = self.calendar_month_indices();
        if indices.is_empty() {
            self.status = "no events in active month".into();
            return;
        }

        let next = match indices
            .iter()
            .position(|index| *index == self.calendar.selected)
        {
            Some(position) => step(position, delta, indices.len()),
            None if delta < 0 => indices.len().saturating_sub(1),
            None => 0,
        };
        self.calendar.selected = indices[next];
    }

    fn calendar_month_select_edge(&mut self, last: bool) {
        let indices = self.calendar_month_indices();
        if let Some(index) = if last {
            indices.last()
        } else {
            indices.first()
        } {
            self.calendar.selected = *index;
        }
    }

    fn calendar_jump_today(&mut self) {
        let today = chrono::Local::now().date_naive();
        let selected = self.calendar.events.iter().position(|event| {
            let Some(start) = event.start.as_ref() else {
                return false;
            };

            if let Ok(value) = chrono::DateTime::parse_from_rfc3339(&start.date_time) {
                return value.with_timezone(&chrono::Local).date_naive() == today;
            }

            let parsed =
                chrono::NaiveDateTime::parse_from_str(&start.date_time, "%Y-%m-%dT%H:%M:%S%.f")
                    .or_else(|_| {
                        chrono::NaiveDateTime::parse_from_str(&start.date_time, "%Y-%m-%dT%H:%M:%S")
                    });

            let Ok(value) = parsed else {
                return false;
            };

            if start.time_zone.as_deref() == Some("UTC") {
                chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(value, chrono::Utc)
                    .with_timezone(&chrono::Local)
                    .date_naive()
                    == today
            } else {
                value.date() == today
            }
        });

        if let Some(index) = selected {
            self.calendar.selected = index;
            self.status = "jumped to today".into();
        } else {
            self.calendar.selected = 0;
            self.status = "no remaining events today".into();
        }
    }

    fn respond_calendar(&mut self, rsvp: calendar::Rsvp) {
        let Some(event) = self.calendar.events.get(self.calendar.selected) else {
            self.status = "no calendar event selected".into();
            return;
        };

        let response = event
            .response_status
            .as_ref()
            .and_then(|status| status.response.as_deref())
            .unwrap_or("");
        if event.is_organizer.unwrap_or(false) || response.eq_ignore_ascii_case("organizer") {
            self.status = "you are the organizer of this event".into();
            return;
        }
        if event.is_cancelled.unwrap_or(false) {
            self.status = "this event is cancelled".into();
            return;
        }
        if event.response_requested == Some(false) {
            self.status = "the organizer did not request a response".into();
            return;
        }

        let id = event.id.clone();
        let (progress, done) = match rsvp {
            calendar::Rsvp::Accept => ("accepting event...", "event accepted"),
            calendar::Rsvp::Decline => ("declining event...", "event declined"),
            calendar::Rsvp::Tentative => ("marking event tentative...", "event marked tentative"),
        };
        self.status = progress.into();

        let s = self.session.clone();
        self.spawn(async move {
            calendar::respond(&s.graph, &id, rsvp, "").await?;
            Ok(AppMessage::Done(done.into()))
        });
    }

    fn outlook_move(&mut self, delta: i32) {
        match self.outlook_focus {
            OutlookFocus::Folders => {
                self.outlook.folder_sel =
                    step(self.outlook.folder_sel, delta, self.outlook.folders.len());
            }
            // In the reading pane, j/k scroll the message body.
            OutlookFocus::Reading => {
                let max = self.reading_max_scroll.get();
                self.outlook.reading_scroll = if delta > 0 {
                    self.outlook
                        .reading_scroll
                        .saturating_add(delta as u16)
                        .min(max)
                } else {
                    self.outlook
                        .reading_scroll
                        .saturating_sub(delta.unsigned_abs() as u16)
                };
            }
            OutlookFocus::Messages => {
                let len = self.outlook.message_rows.len();
                let previous = self.outlook.msg_sel;
                self.outlook.msg_sel = step(self.outlook.msg_sel, delta, len);
                // Scrolling down onto the last row pulls the next page.
                if delta > 0 && len > 0 && self.outlook.msg_sel == len - 1 {
                    self.load_more_messages();
                }
                if self.settings.preview_mail_on_hover && self.outlook.msg_sel != previous {
                    self.open_current_mail();
                }
            }
        }
    }

    fn outlook_out(&mut self) {
        self.outlook_focus = match self.outlook_focus {
            OutlookFocus::Reading => OutlookFocus::Messages,
            OutlookFocus::Messages | OutlookFocus::Folders => OutlookFocus::Folders,
        };
    }

    fn outlook_into(&mut self) {
        if self.outlook_focus != OutlookFocus::Reading {
            self.outlook_enter();
        }
    }

    fn outlook_enter(&mut self) {
        match self.outlook_focus {
            OutlookFocus::Folders => {
                if let Some(f) = self.outlook.folders.get(self.outlook.folder_sel) {
                    self.outlook.msg_sel = 0;
                    self.load_messages(f.id.clone(), ListUpdate::Replace);
                    self.outlook_focus = OutlookFocus::Messages;
                }
            }
            OutlookFocus::Messages | OutlookFocus::Reading => {
                if self.current_mail().is_some() {
                    self.open_current_mail_now();
                    self.outlook_focus = OutlookFocus::Reading;
                    self.schedule_current_read_timer();
                }
            }
        }
    }

    fn current_mail(&self) -> Option<&MailMessage> {
        self.outlook
            .message_rows
            .get(self.outlook.msg_sel)
            .and_then(|&i| self.outlook.messages.get(i))
    }

    fn open_reply(&mut self, mode: ReplyMode) {
        let Some(m) = self.current_mail() else {
            self.status = "select a message first".into();
            return;
        };
        let id = m.id.clone();
        let (kind, field) = match mode {
            ReplyMode::Reply => (ComposeKind::ReplyMail { id }, COMPOSE_BODY),
            ReplyMode::ReplyAll => (ComposeKind::ReplyAllMail { id }, COMPOSE_BODY),
            ReplyMode::Forward => (ComposeKind::ForwardMail { id }, COMPOSE_TO),
        };
        let mut compose = Compose::new(kind, field);
        self.populate_reply_recipients(&mut compose);
        self.overlay = Some(Overlay::Compose(compose));
    }

    fn on_key_teams(&mut self, key: KeyEvent) {
        // Composer captures typing when focused.
        if self.teams.focus == TeamsFocus::Composer {
            let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
            match key.code {
                KeyCode::Esc => {
                    self.teams.replying_to = None;
                    if self.teams.editing_message.take().is_some() {
                        self.teams.composer.clear();
                    }
                    self.teams.focus = TeamsFocus::Messages;
                }
                KeyCode::Tab => {
                    if !self.complete_teams_path() {
                        self.teams.focus = TeamsFocus::List;
                    }
                }
                KeyCode::Char('v') if ctrl => self.paste_clipboard_image(),
                KeyCode::Char('x') if ctrl => self.unstage_teams_image(),
                // Enter sends; Shift/Alt+Enter inserts a newline instead.
                KeyCode::Enter
                    if key
                        .modifiers
                        .intersects(KeyModifiers::SHIFT | KeyModifiers::ALT) =>
                {
                    self.teams.composer.insert('\n')
                }
                KeyCode::Enter => self.teams_send(),
                KeyCode::Up
                    if self.teams.editing_message.is_none()
                        && self.teams.composer.is_empty()
                        && self.teams.images.is_empty() =>
                {
                    self.teams_edit_last_own();
                }
                KeyCode::Up => self.teams.composer.move_row(-1, self.text_width_hint.get()),
                KeyCode::Down => self.teams.composer.move_row(1, self.text_width_hint.get()),
                KeyCode::Backspace => self.teams.composer.backspace(),
                KeyCode::Delete => self.teams.composer.delete(),
                KeyCode::Char('w') if ctrl => self.teams.composer.delete_word_before(),
                KeyCode::Char('u') if ctrl => self.teams.composer.delete_to_line_start(),
                KeyCode::Char('k') if ctrl => self.teams.composer.delete_to_line_end(),
                KeyCode::Left if ctrl => self.teams.composer.word_left(),
                KeyCode::Right if ctrl => self.teams.composer.word_right(),
                KeyCode::Left => self.teams.composer.left(),
                KeyCode::Right => self.teams.composer.right(),
                KeyCode::Home if ctrl => self.teams.composer.start_of_text(),
                KeyCode::End if ctrl => self.teams.composer.end_of_text(),
                KeyCode::Char('a') if ctrl => self.teams.composer.home(),
                KeyCode::Char('e') if ctrl => self.teams.composer.end(),
                KeyCode::Home => self.teams.composer.home(),
                KeyCode::End => self.teams.composer.end(),
                KeyCode::Char(c) if !ctrl => self.teams.composer.insert(c),
                _ => {}
            }
            return;
        }

        match key.code {
            KeyCode::Char('n')
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List =>
            {
                self.show_new_chat()
            }
            KeyCode::Char('v') if self.teams.focus == TeamsFocus::Messages => {
                self.open_full_image()
            }
            KeyCode::Char('x')
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List =>
            {
                self.toggle_selected_teams_chat_hidden();
            }
            KeyCode::Char('X')
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List =>
            {
                self.toggle_selected_teams_chat_stay_archived();
            }
            // Back out to the conversation list from the messages pane.
            KeyCode::Esc | KeyCode::Left | KeyCode::Char('h') => {
                self.teams.focus = TeamsFocus::List;
            }
            // ...and into the conversation from the list, opening it.
            KeyCode::Right | KeyCode::Char('l') if self.teams.focus == TeamsFocus::List => {
                self.teams_enter();
            }
            KeyCode::Tab => {
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List {
                    if self.teams_selected_chat().is_none() {
                        self.status = "open Archive with Enter, or select a chat first".into();
                        return;
                    }
                    if let Some(chat_id) = self.teams.preview_chat_id.clone() {
                        self.open_teams_chat(chat_id);
                        return;
                    }
                }

                self.teams.focus = match self.teams.focus {
                    TeamsFocus::List => TeamsFocus::Messages,
                    TeamsFocus::Messages => TeamsFocus::Composer,
                    TeamsFocus::Composer => TeamsFocus::List,
                };
            }
            KeyCode::Char('c') if self.teams.focus == TeamsFocus::List => {
                self.teams.replying_to = None;
                if self.teams.editing_message.take().is_some() {
                    self.teams.composer.clear();
                }
                if self.teams.preview_chat_id.take().is_some() {
                    self.clear_teams_conversation_view();
                }

                // Toggle chats/channels mode. Listing teams needs a scope that
                // chats don't, so say so plainly instead of letting Graph 403.
                self.teams.mode = match self.teams.mode {
                    TeamsMode::Chats => {
                        if !self.session.config.can_read_teams() {
                            self.status = "channels need the Team.ReadBasic.All permission — grant it, then set M365_TEAMS_CHANNELS=1 and sign in again".into();
                            return;
                        }
                        if self.teams.teams.is_empty() {
                            self.load_teams();
                        }
                        TeamsMode::Channels
                    }
                    TeamsMode::Channels => TeamsMode::Chats,
                };

                if self.teams.mode == TeamsMode::Chats {
                    self.preview_selected_teams_chat();
                }
            }
            KeyCode::Char('i') | KeyCode::Char('a') => {
                if self.teams.editing_message.take().is_some() {
                    self.teams.composer.clear();
                }
                self.teams.replying_to = None;
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List {
                    if self.teams_selected_chat().is_none() {
                        self.status = "open Archive with Enter, or select a chat first".into();
                        return;
                    }
                    if let Some(chat_id) = self.teams.preview_chat_id.clone() {
                        self.open_teams_chat(chat_id);
                    }
                }
                self.teams.focus = TeamsFocus::Composer;
            }
            // Reply to the selected message: same composer, quoted on send.
            KeyCode::Char('r')
                if self.teams.focus == TeamsFocus::Messages && !self.teams.messages.is_empty() =>
            {
                if self
                    .teams
                    .messages
                    .get(self.teams.msg_sel)
                    .is_some_and(ChatMessage::is_system_event)
                {
                    self.status = "system events cannot be replied to".into();
                } else {
                    if self.teams.editing_message.take().is_some() {
                        self.teams.composer.clear();
                    }
                    self.teams.replying_to = Some(self.teams.msg_sel);
                    self.teams.focus = TeamsFocus::Composer;
                }
            }
            KeyCode::Char('t')
                if self.teams.focus == TeamsFocus::Messages && !self.teams.messages.is_empty() =>
            {
                if self
                    .teams
                    .messages
                    .get(self.teams.msg_sel)
                    .is_some_and(ChatMessage::is_system_event)
                {
                    self.status = "system events cannot be reacted to".into();
                } else {
                    self.overlay = Some(Overlay::React { row: 0, col: 0 });
                }
            }
            KeyCode::Char('E')
                if self.teams.focus == TeamsFocus::Messages && !self.teams.messages.is_empty() =>
            {
                self.teams_edit_selected();
            }
            KeyCode::Char('e')
                if key.modifiers.contains(KeyModifiers::SHIFT)
                    && self.teams.focus == TeamsFocus::Messages
                    && !self.teams.messages.is_empty() =>
            {
                self.teams_edit_selected();
            }
            KeyCode::Char('e')
                if self.teams.focus == TeamsFocus::Messages && !self.teams.messages.is_empty() =>
            {
                self.begin_edit_selected_chat_message();
            }
            KeyCode::Home if self.teams.focus == TeamsFocus::List => self.teams_jump_list(false),
            KeyCode::End | KeyCode::Char('G') if self.teams.focus == TeamsFocus::List => {
                self.teams_jump_list(true)
            }
            KeyCode::Char('g')
                if self.teams.mode == TeamsMode::Chats && self.teams.focus == TeamsFocus::List =>
            {
                self.open_selected_contact_profile();
            }
            // Jump back to the newest visible message and resume following it.
            KeyCode::End | KeyCode::Char('g') if self.teams.focus == TeamsFocus::Messages => {
                if let Some(index) = self.last_selectable_teams_message() {
                    self.teams.msg_sel = index;
                }
                self.teams.unseen = 0;
            }
            KeyCode::Home if self.teams.focus == TeamsFocus::Messages => {
                if let Some(index) = self.first_selectable_teams_message() {
                    self.teams.msg_sel = index;
                }
            }
            KeyCode::Up | KeyCode::Char('k') => self.teams_move(-1),
            KeyCode::Down | KeyCode::Char('j') => self.teams_move(1),
            KeyCode::Enter => self.teams_enter(),
            _ => {}
        }
    }

    fn teams_move(&mut self, delta: i32) {
        // In the conversation pane, j/k move the selected message (skipping
        // deleted ones); the selection drives both scroll and reactions.
        if self.teams.focus == TeamsFocus::Messages {
            let n = self.teams.messages.len() as i32;
            if n == 0 {
                return;
            }
            let mut i = self.teams.msg_sel as i32;
            loop {
                i += delta.signum();
                if i < 0 {
                    // Past the oldest loaded message — pull the previous page.
                    self.load_more_teams_messages();
                    return;
                }
                if i >= n {
                    return; // already on the newest
                }
                if self.teams_message_selectable(&self.teams.messages[i as usize]) {
                    self.teams.msg_sel = i as usize;
                    if self.last_selectable_teams_message() == Some(self.teams.msg_sel) {
                        self.teams.unseen = 0;
                        self.teams.replying_to = None; // caught up with the newest
                    }
                    // Prefetch when there are no more visible messages above.
                    if delta < 0
                        && !(0..self.teams.msg_sel)
                            .any(|index| self.teams_message_selectable(&self.teams.messages[index]))
                    {
                        self.load_more_teams_messages();
                    }
                    return;
                }
            }
        }
        match (self.teams.mode, self.teams.focus) {
            (TeamsMode::Chats, TeamsFocus::List) => {
                self.teams.chat_sel = step(self.teams.chat_sel, delta, self.teams_chat_row_count());
                self.preview_selected_teams_chat();
            }
            (TeamsMode::Channels, TeamsFocus::List) => {
                // Navigate channels; if none loaded, navigate teams.
                if self.teams.channels.is_empty() {
                    self.teams.team_sel = step(self.teams.team_sel, delta, self.teams.teams.len());
                } else {
                    self.teams.channel_sel =
                        step(self.teams.channel_sel, delta, self.teams.channels.len());
                }
            }
            _ => {}
        }
    }

    fn teams_enter(&mut self) {
        match self.teams.mode {
            TeamsMode::Chats => {
                if self.teams.chat_sel == self.teams_archive_row() {
                    self.teams.archive_expanded = !self.teams.archive_expanded;
                    self.teams.preview_chat_id = None;
                    self.teams.open_chat_id = None;
                    self.clear_teams_conversation_view();
                    self.status = if self.teams.archive_expanded {
                        format!("Archive opened ({} chat(s))", self.teams_archive_count())
                    } else {
                        "Archive closed".into()
                    };
                    self.persist_ui_state();
                } else if let Some(c) = self.teams_selected_chat() {
                    self.open_teams_chat(c.id.clone());
                }
            }
            TeamsMode::Channels => {
                if self.teams.channels.is_empty() {
                    if let Some(t) = self.teams.teams.get(self.teams.team_sel) {
                        self.load_channels(t.id.clone());
                    }
                } else if let (Some(t), Some(ch)) = (
                    self.teams.teams.get(self.teams.team_sel),
                    self.teams.channels.get(self.teams.channel_sel),
                ) {
                    let team_id = t.id.clone();
                    let ch_id = ch.id.clone();
                    self.teams.open_channel = Some((team_id.clone(), ch_id.clone()));
                    self.teams.open_chat_id = None;
                    self.teams.preview_chat_id = None;
                    self.teams.messages.clear();
                    self.teams.messages_rendered.clear();
                    self.teams.messages_links.clear();
                    self.teams.msg_sel = 0;
                    self.teams.messages_next = None;
                    self.teams.loading_more = false;
                    self.teams.unseen = 0;
                    self.teams.replying_to = None;
                    if self.teams.editing_message.take().is_some() {
                        self.teams.composer.clear();
                    }
                    self.load_channel_messages(team_id, ch_id, ListUpdate::Replace);
                    self.teams.focus = TeamsFocus::Messages;
                }
            }
        }
    }

    fn teams_send(&mut self) {
        if self.teams.sending {
            return;
        }
        if !self.teams.images.is_empty()
            || self.teams.editing_message.is_some()
            || !crate::images::path_tokens(&self.teams.composer.text().chars().collect::<Vec<_>>())
                .is_empty()
        {
            self.teams_send_prepared();
            return;
        }
        let text = self.teams.composer.text().trim().to_string();
        if text.is_empty() {
            return;
        }

        if let Some((chat_id, message_id)) = self.teams.editing_message.clone() {
            self.teams.composer.clear();
            self.teams.editing_message = None;
            let s = self.session.clone();
            self.status = "saving message edit…".into();
            self.spawn(async move {
                chats::edit_message(
                    &s.graph,
                    &chat_id,
                    &message_id,
                    &text,
                    s.config.teams_markdown,
                )
                .await?;
                Ok(AppMessage::Done("message edited".into()))
            });
            return;
        }

        let replying_to = self.teams.replying_to.take();

        if replying_to.is_none() {
            match parse_teams_slash_command(&text) {
                Ok(Some(command)) => {
                    self.teams.composer.clear();
                    self.execute_teams_slash_command(command);
                    return;
                }
                Ok(None) => {}
                Err(error) => {
                    self.status = error;
                    return;
                }
            }
        }

        self.teams.composer.clear();

        match (replying_to, self.teams.mode) {
            (Some(idx), TeamsMode::Chats) => {
                let (Some(chat_id), Some(original)) = (
                    self.teams.open_chat_id.clone(),
                    self.teams.messages.get(idx),
                ) else {
                    return;
                };
                let author = original.author();
                let original = original.clone();
                let s = self.session.clone();
                self.status = format!("replying to {author}…");
                self.spawn(async move {
                    let message = chats::send_reply(
                        &s.graph,
                        &chat_id,
                        &original,
                        &text,
                        s.config.teams_markdown,
                    )
                    .await?;
                    Ok(AppMessage::ChatMessageSent {
                        chat_id,
                        message_id: message.id,
                        status: "reply sent".into(),
                    })
                });
            }
            (Some(idx), TeamsMode::Channels) => {
                let (Some((team_id, channel_id)), Some(original)) = (
                    self.teams.open_channel.clone(),
                    self.teams.messages.get(idx),
                ) else {
                    return;
                };
                let message_id = original.id.clone();
                let s = self.session.clone();
                self.status = "replying…".into();
                self.spawn(async move {
                    channels::send_reply(
                        &s.graph,
                        &team_id,
                        &channel_id,
                        &message_id,
                        &text,
                        s.config.teams_markdown,
                    )
                    .await?;
                    Ok(AppMessage::Done("reply sent".into()))
                });
            }
            (None, TeamsMode::Chats) => {
                if let Some(id) = self.teams.open_chat_id.clone() {
                    self.send_chat_message(id, text);
                }
            }
            (None, TeamsMode::Channels) => {
                if let Some((t, c)) = self.teams.open_channel.clone() {
                    self.send_channel_message(t, c, crate::images::Prepared::Text(text));
                }
            }
        }
    }

    fn selected_owned_chat_message(
        &mut self,
        action: &str,
        past_participle: &str,
    ) -> Option<(String, ChatMessage)> {
        if self.teams.mode != TeamsMode::Chats {
            self.status = format!("{action} is currently available for chats only");
            return None;
        }

        let Some(chat_id) = self.teams.open_chat_id.clone() else {
            self.status = "open a chat first".into();
            return None;
        };
        let Some(message) = self.teams.messages.get(self.teams.msg_sel).cloned() else {
            self.status = "select a message first".into();
            return None;
        };
        if message.is_system_event() {
            self.status = format!("system events cannot be {past_participle}");
            return None;
        }
        if message.deleted_date_time.is_some() {
            self.status = format!("deleted messages cannot be {past_participle}");
            return None;
        }
        let Some(me_id) = self.me.as_ref().map(|me| me.id.as_str()) else {
            self.status = "still loading your profile".into();
            return None;
        };
        if message.author_id() != Some(me_id) {
            self.status = format!("only your own messages can be {past_participle}");
            return None;
        }

        Some((chat_id, message))
    }

    fn begin_edit_selected_chat_message(&mut self) {
        let Some((chat_id, message)) = self.selected_owned_chat_message("edit", "edited") else {
            return;
        };
        if let Err(reason) = validate_plain_message_edit(&message) {
            self.status = reason.into();
            return;
        }

        let raw = message.text();
        let content_type = message
            .body
            .as_ref()
            .and_then(|body| body.content_type.as_deref());
        let rendered = content::render_body(content_type, &raw);
        if !rendered.links.is_empty() {
            self.status = "editing messages with links is not supported yet".into();
            return;
        }
        let text = content::plain(&rendered.text);
        if text.trim().is_empty() {
            self.status = "message has no editable text".into();
            return;
        }

        self.teams.replying_to = None;
        self.teams.editing_message = Some((chat_id, message.id));
        self.teams.composer = TextInput::from(text.as_str());
        self.teams.focus = TeamsFocus::Composer;
        self.status = "editing message".into();
    }

    fn out_of_office_ready(&mut self) -> bool {
        if !self.session.config.out_of_office {
            self.status =
                "Out of Office is disabled; set M365_OUT_OF_OFFICE=1 and sign in again".into();
            return false;
        }
        if !self.session.config.can_write_out_of_office() {
            self.status = "Out of Office needs MailboxSettings.ReadWrite in M365_SCOPES".into();
            return false;
        }
        true
    }

    fn load_out_of_office(&mut self, edit: bool) {
        if !self.out_of_office_ready() {
            return;
        }
        let s = self.session.clone();
        self.status = "loading Out of Office…".into();
        self.spawn(async move {
            let settings = mailbox::settings(&s.graph)
                .await
                .context("loading Out of Office settings")?;
            Ok(AppMessage::OutOfOfficeLoaded { settings, edit })
        });
    }

    fn save_out_of_office(&mut self, editor: out_of_office::Editor) {
        if !self.out_of_office_ready() {
            self.overlay = Some(Overlay::OutOfOfficeEdit(Box::new(editor)));
            return;
        }
        let setting = match editor.setting(self.session.config.teams_markdown) {
            Ok(setting) => setting,
            Err(error) => {
                self.status = error;
                self.overlay = Some(Overlay::OutOfOfficeEdit(Box::new(editor)));
                return;
            }
        };

        let s = self.session.clone();
        self.status = "saving Out of Office…".into();
        self.spawn(async move {
            mailbox::set_automatic_replies(&s.graph, &setting)
                .await
                .context("saving Out of Office settings")?;
            Ok(AppMessage::OutOfOfficeSaved(
                "Out of Office scheduled".into(),
            ))
        });
    }

    fn clear_out_of_office(&mut self) {
        if !self.out_of_office_ready() {
            return;
        }
        let s = self.session.clone();
        self.status = "disabling Out of Office…".into();
        self.spawn(async move {
            mailbox::clear_automatic_replies(&s.graph)
                .await
                .context("disabling Out of Office")?;
            Ok(AppMessage::OutOfOfficeSaved(
                "Out of Office disabled".into(),
            ))
        });
    }

    fn open_out_of_office_web(&mut self) {
        match self.open_web_link(out_of_office::OUTLOOK_OOO_URL) {
            Ok(()) => self.status = "opening Outlook Web automatic replies".into(),
            Err(error) => self.status = format!("could not open Outlook Web: {error:#}"),
        }
    }

    fn execute_teams_slash_command(&mut self, command: TeamsSlashCommand) {
        if matches!(&command, TeamsSlashCommand::Help) {
            self.help_scroll = 0;
            self.overlay = Some(Overlay::TeamsCommands);
            return;
        }
        if matches!(&command, TeamsSlashCommand::HelpAll) {
            self.help_scroll = 0;
            self.overlay = Some(Overlay::Help);
            return;
        }
        match &command {
            TeamsSlashCommand::OutOfOffice => {
                self.load_out_of_office(false);
                return;
            }
            TeamsSlashCommand::OutOfOfficeSet => {
                self.load_out_of_office(true);
                return;
            }
            TeamsSlashCommand::OutOfOfficeClear => {
                if self.out_of_office_ready() {
                    self.overlay = Some(Overlay::ConfirmOutOfOfficeClear);
                }
                return;
            }
            _ => {}
        }

        if self.teams.mode != TeamsMode::Chats {
            self.status = "Teams slash commands are currently available for chats only".into();
            return;
        }

        let Some(chat_id) = self.teams.open_chat_id.clone() else {
            self.status = "open a chat first".into();
            return;
        };

        match command {
            TeamsSlashCommand::Topic(topic) => {
                let Some(chat) = self.teams.chats.iter().find(|chat| chat.id == chat_id) else {
                    tracing::warn!(
                        chat_id = %chat_id,
                        "Teams /topic failed locally: chat metadata is not loaded"
                    );
                    self.status = "ERROR: /topic requires loaded chat metadata".into();
                    return;
                };

                let chat_type = chat.chat_type.as_deref().unwrap_or("unknown").to_string();
                if !chat_type.eq_ignore_ascii_case("group") {
                    tracing::warn!(
                        chat_id = %chat_id,
                        chat_type = %chat_type,
                        "Teams /topic failed locally: only group chats support topics"
                    );
                    self.status =
                        format!("ERROR: /topic works only in group chats (this chat: {chat_type})");
                    return;
                }

                if let Err(error) = validate_chat_topic(&topic) {
                    tracing::warn!(
                        chat_id = %chat_id,
                        reason = %error,
                        "Teams /topic failed local validation"
                    );
                    self.status = format!("ERROR: /topic rejected: {error}");
                    return;
                }

                let s = self.session.clone();
                let result_topic = topic.clone();
                let result_status = format!("Topic updated: {topic}");
                self.status = "updating chat topic…".into();
                tracing::debug!(
                    chat_id = %chat_id,
                    "Teams /topic PATCH starting"
                );
                self.spawn(async move {
                    chats::set_topic(&s.graph, &chat_id, &topic)
                        .await
                        .context("updating Teams chat topic")?;
                    tracing::debug!(
                        chat_id = %chat_id,
                        "Teams /topic PATCH succeeded"
                    );
                    Ok(AppMessage::ChatTopicUpdated {
                        chat_id,
                        topic: result_topic,
                        status: result_status,
                    })
                });
            }
            TeamsSlashCommand::Rename(name) => {
                let Some(chat) = self.teams.chats.iter().find(|chat| chat.id == chat_id) else {
                    tracing::warn!(
                        chat_id = %chat_id,
                        "Teams /rename failed locally: chat metadata is not loaded"
                    );
                    self.status = "ERROR: /rename requires loaded chat metadata".into();
                    return;
                };

                let chat_type = chat.chat_type.as_deref().unwrap_or("unknown").to_string();
                if !chat_type.eq_ignore_ascii_case("group") {
                    tracing::warn!(
                        chat_id = %chat_id,
                        chat_type = %chat_type,
                        "Teams /rename failed locally: Graph can rename only group chats"
                    );
                    self.status = format!(
                        "ERROR: /rename works only in group chats (this chat: {chat_type})"
                    );
                    return;
                }

                if let Err(error) = validate_chat_topic(&name) {
                    tracing::warn!(
                        chat_id = %chat_id,
                        reason = %error,
                        "Teams /rename failed local validation"
                    );
                    self.status = format!("ERROR: /rename rejected: {error}");
                    return;
                }

                self.overlay = Some(Overlay::ConfirmRename { chat_id, name });
            }
            TeamsSlashCommand::Edit(text) => {
                let Some((chat_id, message)) = self.selected_owned_chat_message("edit", "edited")
                else {
                    return;
                };
                if let Err(reason) = validate_plain_message_edit(&message) {
                    self.status = reason.into();
                    return;
                }
                let message_id = message.id;
                let s = self.session.clone();
                self.status = "saving message edit…".into();
                self.spawn(async move {
                    chats::edit_message(
                        &s.graph,
                        &chat_id,
                        &message_id,
                        &text,
                        s.config.teams_markdown,
                    )
                    .await?;
                    Ok(AppMessage::Done("message edited".into()))
                });
            }
            TeamsSlashCommand::Pin => {
                let Some(message) = self.teams.messages.get(self.teams.msg_sel) else {
                    self.status = "select a message first".into();
                    return;
                };
                if message.is_system_event() {
                    self.status = "system events cannot be pinned".into();
                    return;
                }
                self.status =
                    "Teams allows one pinned chat message; confirm replacement with y".into();
                self.overlay = Some(Overlay::ConfirmReplacePin {
                    chat_id,
                    message_id: message.id.clone(),
                });
            }
            TeamsSlashCommand::Unpin => {
                let Some(message) = self.teams.messages.get(self.teams.msg_sel) else {
                    self.status = "select a message first".into();
                    return;
                };
                if message.is_system_event() {
                    self.status = "system events cannot be unpinned".into();
                    return;
                }
                let message_id = message.id.clone();
                let s = self.session.clone();
                self.status = "unpinning message…".into();
                self.spawn(async move {
                    chats::unpin_message(&s.graph, &chat_id, &message_id).await?;
                    Ok(AppMessage::Status("message unpinned".into()))
                });
            }
            TeamsSlashCommand::Pins => {
                let s = self.session.clone();
                self.status = "loading pinned messages…".into();
                self.spawn(async move {
                    let messages = chats::list_pinned_messages(&s.graph, &chat_id).await?;
                    Ok(AppMessage::ChatPinsLoaded {
                        chat_id,
                        messages,
                        status: None,
                    })
                });
            }
            TeamsSlashCommand::Unread => {
                let tenant_id = self
                    .teams
                    .chats
                    .iter()
                    .find(|chat| chat.id == chat_id)
                    .and_then(|chat| self.selected_chat_tenant_id(chat));
                let Some(tenant_id) = tenant_id else {
                    self.status = "cannot determine tenant id for mark-unread".into();
                    return;
                };
                let known_user_id = self.me.as_ref().map(|me| me.id.clone());
                let s = self.session.clone();
                self.status = "marking chat unread…".into();
                self.spawn(async move {
                    let user_id = match known_user_id {
                        Some(id) => id,
                        None => people::me(&s.graph).await?.id,
                    };
                    chats::mark_unread(&s.graph, &chat_id, &user_id, &tenant_id).await?;
                    Ok(AppMessage::ChatMarkedUnread { chat_id })
                });
            }
            TeamsSlashCommand::Delete => {
                let Some((chat_id, message)) =
                    self.selected_owned_chat_message("delete", "deleted")
                else {
                    return;
                };
                let Some(user_id) = self.me.as_ref().map(|me| me.id.clone()) else {
                    self.status = "still loading your profile".into();
                    return;
                };
                self.overlay = Some(Overlay::ConfirmDelete {
                    chat_id,
                    message_id: message.id,
                    user_id,
                });
            }
            TeamsSlashCommand::OutOfOffice
            | TeamsSlashCommand::OutOfOfficeSet
            | TeamsSlashCommand::OutOfOfficeClear
            | TeamsSlashCommand::Help
            | TeamsSlashCommand::HelpAll => {
                unreachable!("handled above")
            }
        }
    }

    fn react_selected(&mut self, emoji: &str) {
        let Some(msg) = self.teams.messages.get(self.teams.msg_sel) else {
            return;
        };
        if msg.is_system_event() {
            self.status = "system events cannot be reacted to".into();
            return;
        }
        let message_id = msg.id.clone();
        let emoji = emoji.to_string();
        let s = self.session.clone();
        if let Some(chat_id) = self.teams.open_chat_id.clone() {
            self.spawn(async move {
                chats::set_reaction(&s.graph, &chat_id, &message_id, &emoji).await?;
                Ok(AppMessage::Done("reaction added".into()))
            });
        } else if let Some((team_id, channel_id)) = self.teams.open_channel.clone() {
            self.spawn(async move {
                channels::set_reaction(&s.graph, &team_id, &channel_id, &message_id, &emoji)
                    .await?;
                Ok(AppMessage::Done("reaction added".into()))
            });
        }
    }

    fn on_key_overlay(&mut self, key: KeyEvent) {
        if key.code == KeyCode::Esc {
            if let Some(Overlay::Compose(c)) = &mut self.overlay {
                if c.ctrl_x {
                    c.ctrl_x = false;
                    self.status.clear();
                    return;
                }
            }
            if let Some(Overlay::MailImages {
                fullscreen: Some(_),
                scroll,
            }) = &self.overlay
            {
                let scroll = *scroll;
                self.overlay = Some(Overlay::MailImages {
                    scroll,
                    fullscreen: None,
                });
                return;
            }
            self.overlay = None;
            return;
        }
        // Take the overlay out so we can mutate self freely, then put it back.
        let overlay = self.overlay.take();
        match overlay {
            Some(Overlay::TeamsCommands) => {
                let max = self.help_max_scroll.get();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.help_scroll = self.help_scroll.saturating_sub(1);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.help_scroll = self.help_scroll.saturating_add(1).min(max);
                    }
                    KeyCode::PageUp => {
                        self.help_scroll = self.help_scroll.saturating_sub(10);
                    }
                    KeyCode::PageDown => {
                        self.help_scroll = self.help_scroll.saturating_add(10).min(max);
                    }
                    KeyCode::Home => self.help_scroll = 0,
                    KeyCode::End => self.help_scroll = max,
                    _ => {}
                }
                self.overlay = Some(Overlay::TeamsCommands);
            }
            Some(Overlay::Help) => {
                let max = self.help_max_scroll.get();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.help_scroll = self.help_scroll.saturating_sub(1);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.help_scroll = self.help_scroll.saturating_add(1).min(max);
                    }
                    KeyCode::PageUp => {
                        self.help_scroll = self.help_scroll.saturating_sub(10);
                    }
                    KeyCode::PageDown => {
                        self.help_scroll = self.help_scroll.saturating_add(10).min(max);
                    }
                    KeyCode::Home => self.help_scroll = 0,
                    KeyCode::End => self.help_scroll = max,
                    _ => {}
                }
                self.overlay = Some(Overlay::Help);
            }
            Some(Overlay::Diagnostics) => {
                let max = self.diagnostics_max_scroll.get();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.diagnostics_scroll = self.diagnostics_scroll.saturating_sub(1);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.diagnostics_scroll =
                            self.diagnostics_scroll.saturating_add(1).min(max);
                    }
                    KeyCode::PageUp => {
                        self.diagnostics_scroll = self.diagnostics_scroll.saturating_sub(10);
                    }
                    KeyCode::PageDown => {
                        self.diagnostics_scroll =
                            self.diagnostics_scroll.saturating_add(10).min(max);
                    }
                    KeyCode::Home => self.diagnostics_scroll = 0,
                    KeyCode::End => self.diagnostics_scroll = max,
                    KeyCode::Char('r') => self.refresh_diagnostics(),
                    KeyCode::Char('c') => self.export_diagnostics(false),
                    KeyCode::Char('l') => self.export_diagnostics(true),
                    _ => {}
                }
                self.overlay = Some(Overlay::Diagnostics);
            }
            Some(Overlay::ContactDiagnostics) => {
                let max = self.contact_diagnostics_max_scroll.get();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.contact_diagnostics_scroll =
                            self.contact_diagnostics_scroll.saturating_sub(1);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.contact_diagnostics_scroll =
                            self.contact_diagnostics_scroll.saturating_add(1).min(max);
                    }
                    KeyCode::PageUp => {
                        self.contact_diagnostics_scroll =
                            self.contact_diagnostics_scroll.saturating_sub(10);
                    }
                    KeyCode::PageDown => {
                        self.contact_diagnostics_scroll =
                            self.contact_diagnostics_scroll.saturating_add(10).min(max);
                    }
                    KeyCode::Home => self.contact_diagnostics_scroll = 0,
                    KeyCode::End => self.contact_diagnostics_scroll = max,
                    KeyCode::Char('r') => self.refresh_contact_diagnostics(),
                    KeyCode::Char('i') => {
                        self.copy_log_contact_automation_peer();
                    }
                    KeyCode::Char('a') => {
                        self.begin_add_contact_automation_peer();
                    }
                    KeyCode::Char('c') => self.export_contact_diagnostics(false),
                    KeyCode::Char('l') => self.export_contact_diagnostics(true),
                    _ => {}
                }
                if self.overlay.is_none() {
                    self.overlay = Some(Overlay::ContactDiagnostics);
                }
            }
            Some(Overlay::ContactProfile) => {
                let max = self.contact_profile_max_scroll.get();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.contact_profile_scroll = self.contact_profile_scroll.saturating_sub(1);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.contact_profile_scroll =
                            self.contact_profile_scroll.saturating_add(1).min(max);
                    }
                    _ => {}
                }
                self.overlay = Some(Overlay::ContactProfile);
            }
            Some(Overlay::OutOfOffice(view)) => match key.code {
                KeyCode::Char('s') => {
                    self.overlay = Some(Overlay::OutOfOfficeEdit(Box::new(
                        out_of_office::Editor::from_settings(&view.settings),
                    )));
                }
                KeyCode::Char('c') => {
                    self.overlay = Some(Overlay::ConfirmOutOfOfficeClear);
                }
                KeyCode::Char('w') => {
                    self.open_out_of_office_web();
                    self.overlay = Some(Overlay::OutOfOffice(view));
                }
                _ => self.overlay = Some(Overlay::OutOfOffice(view)),
            },
            Some(Overlay::ConfirmOutOfOfficeClear) => match key.code {
                KeyCode::Char('y') | KeyCode::Char('Y') => self.clear_out_of_office(),
                _ => self.status = "Out of Office clear cancelled".into(),
            },
            Some(Overlay::OutOfOfficeEdit(mut editor)) => {
                let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
                let shifted = key.modifiers.contains(KeyModifiers::SHIFT)
                    || key.modifiers.contains(KeyModifiers::ALT);

                // Message editing is intentionally handled first and in
                // isolation. Printable characters, including an ordinary
                // space, always go into the message body while it is focused.
                if editor.is_message_field() {
                    let width = self.text_width_hint.get().max(1);
                    match key.code {
                        KeyCode::Tab => editor.next_field(),
                        KeyCode::BackTab => editor.prev_field(),
                        KeyCode::PageUp => editor.page_up(),
                        KeyCode::PageDown => editor.page_down(),
                        KeyCode::Enter if shifted => editor.message.insert('\n'),
                        KeyCode::Char('j') if ctrl => editor.message.insert('\n'),
                        KeyCode::Enter => {
                            self.save_out_of_office(*editor);
                            return;
                        }
                        KeyCode::Backspace => editor.message.backspace(),
                        KeyCode::Delete => editor.message.delete(),
                        KeyCode::Char('w') if ctrl => editor.message.delete_word_before(),
                        KeyCode::Char('u') if ctrl => editor.message.delete_to_line_start(),
                        KeyCode::Char('k') if ctrl => editor.message.delete_to_line_end(),
                        KeyCode::Left if ctrl => editor.message.word_left(),
                        KeyCode::Right if ctrl => editor.message.word_right(),
                        KeyCode::Left => editor.message.left(),
                        KeyCode::Right => editor.message.right(),
                        KeyCode::Up => editor.message.move_row(-1, width),
                        KeyCode::Down => editor.message.move_row(1, width),
                        KeyCode::Home if ctrl => editor.message.start_of_text(),
                        KeyCode::End if ctrl => editor.message.end_of_text(),
                        KeyCode::Home => editor.message.home(),
                        KeyCode::End => editor.message.end(),
                        KeyCode::Char(ch) if !ctrl => editor.message.insert(ch),
                        _ => {}
                    }
                    self.overlay = Some(Overlay::OutOfOfficeEdit(editor));
                    return;
                }

                let is_audience = editor.is_audience_field();
                let is_numeric = editor.is_numeric_field();

                match key.code {
                    KeyCode::Tab => editor.next_field(),
                    KeyCode::BackTab => editor.prev_field(),
                    KeyCode::PageUp => editor.page_up(),
                    KeyCode::PageDown => editor.page_down(),
                    KeyCode::Enter => {
                        self.save_out_of_office(*editor);
                        return;
                    }

                    // Audience uses only arrows. Space is deliberately not a
                    // toggle shortcut, so it can never conflict with text input.
                    KeyCode::Left | KeyCode::Right if is_audience => editor.audience.toggle(),

                    KeyCode::Backspace if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.backspace();
                        }
                    }
                    KeyCode::Delete if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.delete();
                        }
                    }
                    KeyCode::Left if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.left();
                        }
                    }
                    KeyCode::Right if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.right();
                        }
                    }
                    KeyCode::Home if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.home();
                        }
                    }
                    KeyCode::End if is_numeric => {
                        editor.mark_editing();
                        if let Some(input) = editor.active_mut() {
                            input.end();
                        }
                    }

                    KeyCode::Char(ch) if is_numeric && !ctrl && ch.is_ascii_digit() => {
                        editor.insert_digit(ch);
                    }
                    KeyCode::Char(_) if is_numeric && !ctrl => editor.reject_non_digit(),

                    _ => {}
                }
                self.overlay = Some(Overlay::OutOfOfficeEdit(editor));
            }
            Some(Overlay::Calendar) => {
                // Any key besides Esc closes the compact calendar overlay.
            }
            Some(Overlay::CalendarEvent) => {
                if key.code == KeyCode::Char('o') {
                    self.open_selected_calendar_meeting();
                    self.overlay = Some(Overlay::CalendarEvent);
                }
            }
            Some(Overlay::React { mut row, mut col }) => match key.code {
                KeyCode::Left | KeyCode::Char('h') => {
                    col = col.saturating_sub(1);
                    self.overlay = Some(Overlay::React { row, col });
                }
                KeyCode::Right | KeyCode::Char('l') => {
                    if col + 1 < REACTION_COLS {
                        col += 1;
                    }
                    self.overlay = Some(Overlay::React { row, col });
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    row = row.saturating_sub(1);
                    self.overlay = Some(Overlay::React { row, col });
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    if (row + 1) * REACTION_COLS + col < REACTIONS.len() {
                        row += 1;
                    }
                    self.overlay = Some(Overlay::React { row, col });
                }
                KeyCode::Enter => {
                    let index = row * REACTION_COLS + col;
                    if let Some(emoji) = REACTIONS.get(index) {
                        self.react_selected(emoji);
                    }
                }
                _ => {
                    self.overlay = Some(Overlay::React { row, col });
                }
            },
            Some(Overlay::PinnedMessages {
                chat_id,
                messages,
                mut sel,
            }) => match key.code {
                KeyCode::Up | KeyCode::Char('k') => {
                    sel = sel.saturating_sub(1);
                    self.overlay = Some(Overlay::PinnedMessages {
                        chat_id,
                        messages,
                        sel,
                    });
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    sel = sel.saturating_add(1).min(messages.len().saturating_sub(1));
                    self.overlay = Some(Overlay::PinnedMessages {
                        chat_id,
                        messages,
                        sel,
                    });
                }
                KeyCode::Enter => {
                    if self.teams.open_chat_id.as_deref() != Some(chat_id.as_str()) {
                        self.status = "that pinned message belongs to another chat".into();
                    } else if let Some(message) = messages.get(sel).cloned() {
                        let message_id = message.id.clone();

                        // /pins already fetched the full message. If it is older
                        // than the currently loaded conversation window, merge it
                        // locally before selecting it so Enter always has a real
                        // message row to scroll to.
                        if !self
                            .teams
                            .messages
                            .iter()
                            .any(|loaded| loaded.id == message_id)
                        {
                            let mut combined = self.teams.messages.clone();
                            combined.push(message);
                            let next = self.teams.messages_next.clone();
                            self.set_teams_messages(combined, next, ListUpdate::Replace);
                        }

                        if let Some(index) = self
                            .teams
                            .messages
                            .iter()
                            .position(|loaded| loaded.id == message_id)
                        {
                            self.teams.msg_sel = index;
                            self.teams.focus = TeamsFocus::Messages;
                            self.refresh_selected_teams_images();
                            self.status = "jumped to pinned message".into();
                        }
                    }
                }
                KeyCode::Char('u') => {
                    let Some(message_id) = messages.get(sel).map(|message| message.id.clone())
                    else {
                        self.status = "no pinned message selected".into();
                        return;
                    };
                    let s = self.session.clone();
                    self.status = "unpinning message…".into();
                    self.spawn(async move {
                        chats::unpin_message(&s.graph, &chat_id, &message_id).await?;
                        let messages = chats::list_pinned_messages(&s.graph, &chat_id).await?;
                        Ok(AppMessage::ChatPinsLoaded {
                            chat_id,
                            messages,
                            status: Some("message unpinned".into()),
                        })
                    });
                }
                _ => {
                    self.overlay = Some(Overlay::PinnedMessages {
                        chat_id,
                        messages,
                        sel,
                    });
                }
            },
            Some(Overlay::ConfirmAutomationPeer {
                peer_id,
                dotenv_path,
            }) => match key.code {
                KeyCode::Char('y') | KeyCode::Char('Y') => {
                    self.persist_contact_automation_peer(peer_id, dotenv_path);
                }
                _ => {
                    tracing::info!("F7 automation peer add cancelled: peer_id={peer_id}");
                    self.status = "automation peer add cancelled".into();
                }
            },
            Some(Overlay::ConfirmRename { chat_id, name }) => match key.code {
                KeyCode::Char('y') | KeyCode::Char('Y') => {
                    let s = self.session.clone();
                    let result_name = name.clone();
                    let result_status = format!("Chat renamed: {name}");
                    self.status = "renaming chat…".into();
                    tracing::debug!(
                        chat_id = %chat_id,
                        "Teams /rename PATCH starting"
                    );
                    self.spawn(async move {
                        chats::set_topic(&s.graph, &chat_id, &name)
                            .await
                            .context("renaming Teams group chat")?;
                        tracing::debug!(
                            chat_id = %chat_id,
                            "Teams /rename PATCH succeeded"
                        );
                        Ok(AppMessage::ChatTopicUpdated {
                            chat_id,
                            topic: result_name,
                            status: result_status,
                        })
                    });
                }
                _ => {
                    self.status = "chat rename cancelled".into();
                }
            },
            Some(Overlay::ConfirmReplacePin {
                chat_id,
                message_id,
            }) => match key.code {
                KeyCode::Char('y') | KeyCode::Char('Y') => {
                    let s = self.session.clone();
                    self.status = "pinning message…".into();
                    self.spawn(async move {
                        chats::pin_message(&s.graph, &chat_id, &message_id).await?;
                        Ok(AppMessage::Status(
                            "message pinned; previous pin replaced if present".into(),
                        ))
                    });
                }
                _ => {
                    self.status = "pin replacement cancelled".into();
                }
            },
            Some(Overlay::ConfirmDelete {
                chat_id,
                message_id,
                user_id,
            }) => match key.code {
                KeyCode::Char('y') | KeyCode::Char('Y') => {
                    let s = self.session.clone();
                    self.status = "deleting message…".into();
                    self.spawn(async move {
                        chats::soft_delete_message(&s.graph, &user_id, &chat_id, &message_id)
                            .await?;
                        Ok(AppMessage::Done("message deleted".into()))
                    });
                }
                _ => {
                    self.status = "message deletion cancelled".into();
                }
            },
            Some(Overlay::Attachments) => match key.code {
                KeyCode::Char(c @ '1'..='9') => {
                    self.download_attachment((c as u8 - b'1') as usize);
                }
                _ => self.overlay = Some(Overlay::Attachments),
            },
            Some(Overlay::Links) => match key.code {
                KeyCode::Char(c @ '1'..='9') => {
                    self.open_link((c as u8 - b'1') as usize);
                }
                KeyCode::Char('y') => {
                    // Copy a link instead of opening it.
                    if let Some(url) = self.focused_links().first().cloned() {
                        self.copy_to_clipboard(&url, "link");
                    }
                }
                _ => self.overlay = Some(Overlay::Links),
            },
            Some(Overlay::MailImages { scroll, fullscreen }) => {
                self.on_key_mail_images(key, scroll, fullscreen);
            }
            Some(Overlay::Presence) => match key.code {
                KeyCode::Char(c @ '1'..='6') => {
                    let idx = (c as u8 - b'1') as usize;
                    if let Some(opt) = PRESENCE_OPTIONS.get(idx) {
                        if !self.session.config.can_write_presence() {
                            self.status =
                                "setting presence needs M365_PRESENCE_WRITE=1 + Presence.ReadWrite consent".into();
                        } else {
                            self.status = format!("setting presence to {}…", opt.label);
                            self.set_presence(opt);
                        }
                    }
                }
                KeyCode::Char('m') => {
                    if !self.session.config.can_write_presence() {
                        self.status =
                            "setting status message needs M365_PRESENCE_WRITE=1 + Presence.ReadWrite consent".into();
                    } else {
                        self.open_presence_status_message_editor();
                    }
                }
                KeyCode::Char('c') | KeyCode::Char('x') => {
                    if !self.session.config.can_write_presence() {
                        self.status =
                            "clearing status message needs M365_PRESENCE_WRITE=1 + Presence.ReadWrite consent".into();
                    } else {
                        self.clear_presence_status_message();
                    }
                }
                KeyCode::Char('a') => {
                    if !self.session.config.can_write_presence() {
                        self.status =
                            "setting presence needs M365_PRESENCE_WRITE=1 + Presence.ReadWrite consent".into();
                    } else {
                        self.status = "clearing presence…".into();
                        self.clear_presence();
                    }
                }
                _ => self.overlay = Some(Overlay::Presence), // ignore other keys
            },
            Some(Overlay::PresenceMessage(mut input)) => {
                let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
                match key.code {
                    KeyCode::Enter => self.set_presence_status_message(input.text()),
                    KeyCode::Backspace => {
                        input.backspace();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Delete => {
                        input.delete();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Left if ctrl => {
                        input.word_left();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Right if ctrl => {
                        input.word_right();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Left => {
                        input.left();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Right => {
                        input.right();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Home if ctrl => {
                        input.start_of_text();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::End if ctrl => {
                        input.end_of_text();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Home => {
                        input.home();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::End => {
                        input.end();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Char('w') if ctrl => {
                        input.delete_word_before();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Char('u') if ctrl => {
                        input.delete_to_line_start();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Char('k') if ctrl => {
                        input.delete_to_line_end();
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    KeyCode::Char(ch) if !ctrl => {
                        input.insert(ch);
                        self.overlay = Some(Overlay::PresenceMessage(input));
                    }
                    _ => self.overlay = Some(Overlay::PresenceMessage(input)),
                }
            }
            Some(Overlay::NtfySnooze { mut sel }) => {
                let max = NTFY_SNOOZE_HOURS.len();
                match key.code {
                    KeyCode::Up | KeyCode::Char('k') => {
                        sel = sel.saturating_sub(1);
                        self.overlay = Some(Overlay::NtfySnooze { sel });
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        sel = sel.saturating_add(1).min(max);
                        self.overlay = Some(Overlay::NtfySnooze { sel });
                    }
                    KeyCode::Enter => {
                        if let Some(&hours) = NTFY_SNOOZE_HOURS.get(sel) {
                            self.set_ntfy_snooze(hours);
                        } else if sel == NTFY_SNOOZE_HOURS.len() {
                            self.clear_ntfy_snooze();
                        }
                    }
                    KeyCode::Char('c') | KeyCode::Char('0') => self.clear_ntfy_snooze(),
                    _ => self.overlay = Some(Overlay::NtfySnooze { sel }),
                }
            }
            Some(Overlay::Search { mut query }) => match key.code {
                KeyCode::Enter => {
                    let q = query.clone();
                    self.run_search(q);
                }
                KeyCode::Backspace => {
                    query.pop();
                    self.overlay = Some(Overlay::Search { query });
                }
                KeyCode::Char(c) => {
                    query.push(c);
                    self.overlay = Some(Overlay::Search { query });
                }
                _ => self.overlay = Some(Overlay::Search { query }),
            },
            Some(Overlay::FolderSearch { mut query, mut sel }) => {
                let matches = filter_folders(&self.outlook.folders, &query);
                match key.code {
                    KeyCode::Enter => {
                        if let Some(&idx) = matches.get(sel) {
                            self.outlook.folder_sel = idx;
                            self.outlook_enter();
                        }
                    }
                    KeyCode::Up => {
                        sel = sel.saturating_sub(1);
                        self.overlay = Some(Overlay::FolderSearch { query, sel });
                    }
                    KeyCode::Down => {
                        if sel + 1 < matches.len() {
                            sel += 1;
                        }
                        self.overlay = Some(Overlay::FolderSearch { query, sel });
                    }
                    KeyCode::Backspace => {
                        query.pop();
                        self.overlay = Some(Overlay::FolderSearch { query, sel: 0 });
                    }
                    KeyCode::Char(c) => {
                        query.push(c);
                        self.overlay = Some(Overlay::FolderSearch { query, sel: 0 });
                    }
                    _ => self.overlay = Some(Overlay::FolderSearch { query, sel }),
                }
            }
            Some(Overlay::NewChat {
                mut query,
                mut results,
                mut sel,
                mut loading,
            }) => {
                match key.code {
                    KeyCode::Enter => {
                        if let Some(user) = results.get(sel) {
                            self.open_chat_with_user(user.id.clone());
                        }
                    }
                    KeyCode::Up => sel = sel.saturating_sub(1),
                    KeyCode::Down => sel = (sel + 1).min(results.len().saturating_sub(1)),
                    KeyCode::Backspace => {
                        query.pop();
                        results.clear();
                        sel = 0;
                        loading = !query.trim().is_empty();
                        self.search_chat_people(query.clone());
                    }
                    KeyCode::Char(c) if !key.modifiers.contains(KeyModifiers::CONTROL) => {
                        query.push(c);
                        results.clear();
                        sel = 0;
                        loading = true;
                        self.search_chat_people(query.clone());
                    }
                    _ => {}
                }
                self.overlay = Some(Overlay::NewChat {
                    query,
                    results,
                    sel,
                    loading,
                });
            }
            Some(Overlay::Palette { mut query, mut sel }) => {
                let matches = filter_commands(&query);
                match key.code {
                    KeyCode::Enter => {
                        if let Some((id, _)) = matches.get(sel) {
                            self.run_command(id);
                        }
                    }
                    KeyCode::Up => {
                        sel = sel.saturating_sub(1);
                        self.overlay = Some(Overlay::Palette { query, sel });
                    }
                    KeyCode::Down => {
                        if sel + 1 < matches.len() {
                            sel += 1;
                        }
                        self.overlay = Some(Overlay::Palette { query, sel });
                    }
                    KeyCode::Backspace => {
                        query.pop();
                        self.overlay = Some(Overlay::Palette { query, sel: 0 });
                    }
                    KeyCode::Char(c) => {
                        query.push(c);
                        self.overlay = Some(Overlay::Palette { query, sel: 0 });
                    }
                    _ => self.overlay = Some(Overlay::Palette { query, sel }),
                }
            }
            Some(Overlay::Compose(mut c)) => self.on_key_compose(key, &mut c),
            Some(Overlay::MoveMail { mut sel }) => match key.code {
                KeyCode::Enter => {
                    if let Some(folder) = self.outlook.folders.get(sel) {
                        let dest_id = folder.id.clone();
                        let dest_label = folder
                            .display_name
                            .clone()
                            .unwrap_or_else(|| dest_id.clone());
                        self.move_current_mail(dest_id, dest_label);
                    }
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    sel = sel.saturating_sub(1);
                    self.overlay = Some(Overlay::MoveMail { sel });
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    if sel + 1 < self.outlook.folders.len() {
                        sel += 1;
                    }
                    self.overlay = Some(Overlay::MoveMail { sel });
                }
                KeyCode::Home | KeyCode::Char('g')
                    if !key.modifiers.contains(KeyModifiers::SHIFT) =>
                {
                    self.overlay = Some(Overlay::MoveMail { sel: 0 });
                }
                KeyCode::End | KeyCode::Char('G') => {
                    let sel = self.outlook.folders.len().saturating_sub(1);
                    self.overlay = Some(Overlay::MoveMail { sel });
                }
                KeyCode::Char('g') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                    let sel = self.outlook.folders.len().saturating_sub(1);
                    self.overlay = Some(Overlay::MoveMail { sel });
                }
                _ => self.overlay = Some(Overlay::MoveMail { sel }),
            },
            Some(Overlay::ViewImage { keys, mut sel }) => match key.code {
                KeyCode::Char('j') | KeyCode::Down | KeyCode::Char('l') | KeyCode::Right => {
                    if !keys.is_empty() {
                        sel = (sel + 1) % keys.len();
                    }
                    self.overlay = Some(Overlay::ViewImage { keys, sel });
                }
                KeyCode::Char('k') | KeyCode::Up | KeyCode::Char('h') | KeyCode::Left => {
                    if !keys.is_empty() {
                        sel = (sel + keys.len() - 1) % keys.len();
                    }
                    self.overlay = Some(Overlay::ViewImage { keys, sel });
                }
                _ => self.overlay = Some(Overlay::ViewImage { keys, sel }),
            },
            Some(Overlay::Settings { mut sel }) => match key.code {
                KeyCode::Up | KeyCode::Char('k') => {
                    sel = sel.saturating_sub(1);
                    self.overlay = Some(Overlay::Settings { sel });
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    self.overlay = Some(Overlay::Settings { sel });
                }
                KeyCode::Enter | KeyCode::Char(' ') => {
                    self.settings.preview_mail_on_hover = !self.settings.preview_mail_on_hover;
                    self.status = match self.settings.save() {
                        Ok(()) => "settings saved".into(),
                        Err(error) => format!("could not save settings: {error:#}"),
                    };
                    if self.settings.preview_mail_on_hover
                        && self.outlook_focus == OutlookFocus::Messages
                    {
                        self.open_current_mail();
                    }
                    self.overlay = Some(Overlay::Settings { sel });
                }
                _ => self.overlay = Some(Overlay::Settings { sel }),
            },
            Some(Overlay::Notice(_)) => {}
            None => {}
        }
    }

    fn on_key_compose(&mut self, key: KeyEvent, c: &mut Compose) {
        let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
        let is_body = c.field == COMPOSE_BODY;
        let suggestions = self.compose_suggestions(c);
        let has_suggestions = !suggestions.is_empty();
        // Width the body is laid out at, so Up/Down follow what's on screen.
        let width = self.text_width_hint.get();

        if c.ctrl_x {
            c.ctrl_x = false;
            match key.code {
                KeyCode::Char('e') => {
                    self.pending_external_edit = true;
                    self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                    return;
                }
                KeyCode::Char('x') => {
                    if let Some((p, _)) = c.attachments.pop() {
                        self.status = format!("removed {}", p.display());
                    }
                    self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                    return;
                }
                _ => {}
            }
        }

        match key.code {
            // -- actions --
            KeyCode::Char('r') if ctrl => {
                c.kind = match &c.kind {
                    ComposeKind::ReplyMail { id } => ComposeKind::ReplyAllMail { id: id.clone() },
                    ComposeKind::ReplyAllMail { id } => ComposeKind::ReplyMail { id: id.clone() },
                    _ => {
                        self.overlay =
                            Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                        return;
                    }
                };
                if !c.edit_recipients {
                    self.populate_reply_recipients(c);
                }
                self.status = format!("compose mode: {}", c.kind.title());
            }
            KeyCode::Char('e')
                if ctrl
                    && matches!(
                        c.kind,
                        ComposeKind::ReplyMail { .. } | ComposeKind::ReplyAllMail { .. }
                    ) =>
            {
                c.edit_recipients = true;
                c.field = COMPOSE_TO;
                self.status = "reply recipients editable".into();
            }
            KeyCode::Char('s') if ctrl => {
                self.submit_compose(c);
                return;
            }
            KeyCode::Tab if self.complete_mail_path(c) => {}
            KeyCode::Tab if has_suggestions => {
                let selected = c.suggestion_sel.min(suggestions.len() - 1);
                accept_recipient_suggestion(c, &suggestions[selected]);
            }
            KeyCode::Tab => {
                c.field = next_field(c.fields(), c.field);
                c.suggestion_sel = 0;
            }
            KeyCode::BackTab => c.field = prev_field(c.fields(), c.field),
            KeyCode::Enter => {
                if has_suggestions {
                    let selected = c.suggestion_sel.min(suggestions.len() - 1);
                    accept_recipient_suggestion(c, &suggestions[selected]);
                } else if is_body {
                    c.active_mut().insert('\n');
                } else if c.field == COMPOSE_ATTACH {
                    let msg = c.stage_attachment();
                    if !msg.is_empty() {
                        self.status = msg;
                    }
                } else {
                    c.field = next_field(c.fields(), c.field);
                }
            }
            // Ctrl+X is a prefix: e opens $EDITOR, x unstages the last file.
            KeyCode::Char('x') if ctrl => {
                c.ctrl_x = true;
                self.status = "Ctrl+X — e $EDITOR · x unstage last attachment".into();
            }

            // -- deletion --
            KeyCode::Backspace => {
                c.active_mut().backspace();
                c.suggestion_sel = 0;
            }
            KeyCode::Delete => {
                c.active_mut().delete();
                c.suggestion_sel = 0;
            }
            KeyCode::Char('w') if ctrl => {
                c.active_mut().delete_word_before();
                c.suggestion_sel = 0;
            }
            KeyCode::Char('u') if ctrl => {
                c.active_mut().delete_to_line_start();
                c.suggestion_sel = 0;
            }
            KeyCode::Char('k') if ctrl => {
                c.active_mut().delete_to_line_end();
                c.suggestion_sel = 0;
            }

            // -- movement --
            KeyCode::Left if ctrl => c.active_mut().word_left(),
            KeyCode::Right if ctrl => c.active_mut().word_right(),
            KeyCode::Left => c.active_mut().left(),
            KeyCode::Right => c.active_mut().right(),
            KeyCode::Up if has_suggestions => c.suggestion_sel = c.suggestion_sel.saturating_sub(1),
            KeyCode::Down if has_suggestions => {
                c.suggestion_sel = (c.suggestion_sel + 1).min(suggestions.len() - 1)
            }
            KeyCode::Up if is_body => c.active_mut().move_row(-1, width),
            KeyCode::Down if is_body => c.active_mut().move_row(1, width),
            KeyCode::Up => c.field = prev_field(c.fields(), c.field),
            KeyCode::Down => c.field = next_field(c.fields(), c.field),
            KeyCode::Home if ctrl => c.active_mut().start_of_text(),
            KeyCode::End if ctrl => c.active_mut().end_of_text(),
            KeyCode::Char('a') if ctrl => c.active_mut().home(),
            KeyCode::Char('e') if ctrl => c.active_mut().end(),
            KeyCode::Home => c.active_mut().home(),
            KeyCode::End => c.active_mut().end(),
            KeyCode::PageUp => c.active_mut().move_row(-10, width),
            KeyCode::PageDown => c.active_mut().move_row(10, width),

            // -- typing (ignore other Ctrl chords so they can't insert junk) --
            KeyCode::Char(ch) if !ctrl => {
                c.active_mut().insert(ch);
                c.suggestion_sel = 0;
            }
            _ => {}
        }
        // Put the (mutated) compose overlay back.
        self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
    }

    pub fn on_paste(&mut self, text: String) {
        if self.overlay.is_none()
            && self.screen == Screen::Teams
            && self.teams.focus == TeamsFocus::Composer
        {
            match crate::images::from_data_uri(text.trim()) {
                Some(Ok(mut img)) => {
                    let ext = crate::images::image_extension(&img.bytes).unwrap_or("png");
                    img.name = crate::images::unique_paste_name(&self.teams.images, ext);
                    self.stage_teams_image(img);
                    return;
                }
                Some(Err(e)) => {
                    self.status = e;
                    return;
                }
                None => {}
            }
        }

        self.note_presence_activity();
        match self.overlay.take() {
            Some(Overlay::Compose(mut c)) => {
                c.active_mut().insert_str(&text);
                self.overlay = Some(Overlay::Compose(c));
            }
            Some(Overlay::Search { mut query }) => {
                query.push_str(text.trim());
                self.overlay = Some(Overlay::Search { query });
            }
            Some(Overlay::PresenceMessage(mut input)) => {
                let text = text.replace('\r', "").replace('\n', " ");
                input.insert_str(&text);
                self.overlay = Some(Overlay::PresenceMessage(input));
            }
            Some(Overlay::OutOfOfficeEdit(mut editor)) => {
                editor.paste(&text);
                self.overlay = Some(Overlay::OutOfOfficeEdit(editor));
            }
            other => {
                self.overlay = other;
                if self.overlay.is_none()
                    && self.screen == Screen::Teams
                    && self.teams.focus == TeamsFocus::Composer
                {
                    self.teams.composer.insert_str(&text);
                }
            }
        }
    }

    fn submit_compose(&mut self, c: &mut Compose) {
        use m365_core::mail::Outgoing;

        // Recipients are required for a new mail and for a forward.
        let kind = match &c.kind {
            ComposeKind::NewMail => {
                let to = parse_recipients(&c.to.text());
                let cc = parse_recipients(&c.cc.text());
                let bcc = parse_recipients(&c.bcc.text());
                if to.is_empty() {
                    self.status = "add at least one recipient".into();
                    self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                    return;
                }
                Outgoing::New {
                    to,
                    cc,
                    bcc,
                    subject: c.subject.text(),
                }
            }
            ComposeKind::ReplyMail { id } => Outgoing::Reply {
                id: id.clone(),
                recipients: c
                    .edit_recipients
                    .then(|| m365_core::mail::OutgoingRecipients {
                        to: parse_recipients(&c.to.text()),
                        cc: parse_recipients(&c.cc.text()),
                        bcc: parse_recipients(&c.bcc.text()),
                    }),
            },
            ComposeKind::ReplyAllMail { id } => Outgoing::ReplyAll {
                id: id.clone(),
                recipients: c
                    .edit_recipients
                    .then(|| m365_core::mail::OutgoingRecipients {
                        to: parse_recipients(&c.to.text()),
                        cc: parse_recipients(&c.cc.text()),
                        bcc: parse_recipients(&c.bcc.text()),
                    }),
            },
            ComposeKind::ForwardMail { id } => {
                let to = parse_recipients(&c.to.text());
                let cc = parse_recipients(&c.cc.text());
                let bcc = parse_recipients(&c.bcc.text());
                if to.is_empty() {
                    self.status = "add at least one recipient to forward to".into();
                    self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                    return;
                }
                Outgoing::Forward {
                    id: id.clone(),
                    to,
                    cc,
                    bcc,
                }
            }
        };

        let paths = c.attachments.clone();
        let body_text = c.body.text();
        let prepared_mail = match crate::images::prepare_mail(&body_text) {
            Ok(prepared) => prepared,
            Err(error) => {
                self.status = error;
                self.overlay = Some(Overlay::Compose(std::mem::replace(c, empty_compose())));
                return;
            }
        };
        let body_is_html = prepared_mail.html.is_some();
        let body = prepared_mail.html.unwrap_or(body_text);
        let inline_images = prepared_mail.images;
        let label = match paths.len() {
            0 => "sending…".to_string(),
            1 => "sending with 1 attachment…".to_string(),
            n => format!("sending with {n} attachments…"),
        };
        self.status = label;
        self.overlay = None;

        let s = self.session.clone();
        self.spawn(async move {
            // Read the staged files here, off the UI thread.
            let mut attachments = Vec::with_capacity(paths.len() + inline_images.len());
            for (path, _) in &paths {
                let bytes =
                    std::fs::read(path).with_context(|| format!("reading {}", path.display()))?;
                let name = path
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_else(|| "attachment".to_string());
                attachments.push(m365_core::mail::OutgoingAttachment {
                    name,
                    bytes,
                    content_type: None,
                    content_id: None,
                    is_inline: false,
                });
            }
            for (image, content_id) in inline_images {
                attachments.push(m365_core::mail::OutgoingAttachment {
                    name: image.name,
                    bytes: image.bytes,
                    content_type: Some(image.mime.to_string()),
                    content_id: Some(content_id),
                    is_inline: true,
                });
            }
            let count = attachments.len();
            mail::send_message(&s.graph, kind, &body, body_is_html, attachments).await?;
            Ok(AppMessage::MailSent(match count {
                0 => "sent".to_string(),
                1 => "sent with 1 attachment".to_string(),
                n => format!("sent with {n} attachments"),
            }))
        });
    }

    fn run_search(&mut self, query: String) {
        self.overlay = None;
        if query.trim().is_empty() {
            return;
        }
        let s = self.session.clone();
        self.status = format!("searching \"{query}\"…");
        self.outlook.msg_sel = 0;
        self.spawn(async move {
            let results = mail::search(&s.graph, &query, 50).await?;
            Ok(AppMessage::Messages {
                items: results,
                next: None,
                mode: ListUpdate::Replace,
            })
        });
    }

    pub fn set_graphics(&mut self, graphics: Option<crate::termimg::Graphics>) {
        self.graphics = graphics;
        self.fetch_message_images();
        self.fetch_user_photos();
    }

    pub fn image_display_rows(&self, key: &str, width: u16) -> u16 {
        self.image_cache
            .get(key)
            .map(|image| image.rows_for_width(width))
            .unwrap_or(1)
    }

    pub fn image_is_pending(&self, key: &str) -> bool {
        self.image_loading.contains(key)
            || self.outlook.reading_images.iter().any(|image| {
                self.outlook.attachment_loading.contains(&image.message_id)
                    && crate::termimg::mail_cache_key(&image.message_id, &image.image.src) == key
            })
    }

    pub fn avatar_cell_size(&self, key: &str) -> Option<(u16, u16)> {
        if self.image_failed.contains(key) {
            return None;
        }
        let rows = crate::termimg::AVATAR_IMAGE_ROWS;
        if let Some(img) = self.image_cache.get(key) {
            return Some((img.cols_for_rows(rows).clamp(1, 8), rows));
        }
        if self.image_loading.contains(key) {
            return Some((rows.saturating_mul(2).max(1), rows));
        }
        None
    }

    fn load_mail_thread(&mut self, conversation_id: String) {
        self.status = "loading thread…".into();
        let s = self.session.clone();
        self.spawn(async move {
            let (items, truncated) =
                mail::list_conversation(&s.graph, &conversation_id, PAGE_SIZE).await?;
            Ok(AppMessage::MailThread {
                conversation_id,
                items,
                truncated,
            })
        });
    }

    fn search_chat_people(&self, query: String) {
        if query.trim().is_empty() {
            return;
        }
        let s = self.session.clone();
        self.spawn(async move {
            tokio::time::sleep(Duration::from_millis(250)).await;
            let result = people::search_users(&s.graph, &query)
                .await
                .map_err(|e| m365_core::util::graph_error_summary(&e.to_string()));
            Ok(AppMessage::PeopleSearch { query, result })
        });
    }

    fn open_chat_with_user(&mut self, user_id: String) {
        let Some(my_id) = self.me.as_ref().map(|me| me.id.clone()) else {
            self.status = "still loading your account; try again shortly".into();
            return;
        };
        self.status = "opening chat…".into();
        let s = self.session.clone();
        self.spawn(async move {
            let chat = chats::create_one_on_one(&s.graph, &my_id, &user_id).await?;
            Ok(AppMessage::OpenChat(Some(chat.id)))
        });
    }

    fn show_new_chat(&mut self) {
        if !self.session.config.can_search_people() {
            self.status = "directory search needs User.ReadBasic.All — grant it, set M365_PEOPLE_SEARCH=1 and sign in again".into();
            return;
        }
        self.overlay = Some(Overlay::NewChat {
            query: String::new(),
            results: Vec::new(),
            sel: 0,
            loading: false,
        });
    }

    fn toggle_pane_layout(&mut self) {
        self.panes_vertical = !self.panes_vertical;
        self.status = if self.panes_vertical {
            "panes stacked vertically".into()
        } else {
            "panes stacked horizontally".into()
        };
    }

    fn resize_folder_panel(&mut self, delta: i16) {
        if self.panes_vertical {
            let height = self
                .outlook
                .folder_height
                .unwrap_or_else(|| folder_panel_height(&self.outlook.folders));
            let height = resize_panel_extent(height, delta, MIN_FOLDER_PANEL_HEIGHT);
            self.outlook.folder_height = Some(height);
            self.status = format!("Folders height: {height}");
            return;
        }
        let width = self
            .outlook
            .folder_width
            .unwrap_or(DEFAULT_FOLDER_PANEL_WIDTH);
        let width = resize_panel_extent(width, delta, MIN_FOLDER_PANEL_WIDTH);
        self.outlook.folder_width = Some(width);
        self.status = format!("Folders width: {width}");
    }

    fn outlook_jump(&mut self, to_end: bool) {
        match self.outlook_focus {
            OutlookFocus::Folders => {
                self.outlook.folder_sel = if to_end {
                    self.outlook.folders.len().saturating_sub(1)
                } else {
                    0
                };
            }
            OutlookFocus::Messages => {
                self.outlook.msg_sel = if to_end {
                    self.outlook.message_rows.len().saturating_sub(1)
                } else {
                    0
                };
            }
            OutlookFocus::Reading => {
                self.outlook.reading_scroll = if to_end {
                    self.reading_max_scroll.get()
                } else {
                    0
                };
            }
        }
        if self.outlook_focus == OutlookFocus::Messages && self.settings.preview_mail_on_hover {
            self.open_current_mail();
        }
    }

    fn selection_reading_key(&self) -> Option<ReadingKey> {
        self.current_mail()
            .map(|message| reading_key(message, self.outlook.threaded))
    }

    fn open_current_mail(&mut self) {
        self.open_selected_mail(false);
    }

    fn open_current_mail_now(&mut self) {
        self.open_selected_mail(true);
    }

    fn open_selected_mail(&mut self, immediate: bool) {
        let Some(key) = self.selection_reading_key() else {
            return;
        };
        if self.show_cached_mail(&key) {
            self.outlook.mail_fetch_generation = self.outlook.mail_fetch_generation.wrapping_add(1);
            self.outlook.pending_reading = None;
            return;
        }
        if immediate {
            self.fetch_selected_mail(key);
        } else if self.outlook.pending_reading.as_ref() != Some(&key) {
            self.schedule_mail_fetch();
        }
    }

    fn show_cached_mail(&mut self, key: &ReadingKey) -> bool {
        match key {
            ReadingKey::Thread(conversation_id) => {
                let Some((items, truncated)) =
                    self.outlook.thread_cache.get(conversation_id).cloned()
                else {
                    return false;
                };
                self.show_thread(items, truncated);
                true
            }
            ReadingKey::Message(id) => {
                let Some(cached) = self.outlook.message_cache.get(id).cloned() else {
                    return false;
                };
                self.show_message(cached);
                true
            }
        }
    }

    fn schedule_mail_fetch(&mut self) {
        self.outlook.mail_fetch_generation = self.outlook.mail_fetch_generation.wrapping_add(1);
        let generation = self.outlook.mail_fetch_generation;
        let tx = self.tx.clone();
        tokio::spawn(async move {
            tokio::time::sleep(MAIL_FETCH_DELAY).await;
            let _ = tx.send(AppMessage::OpenMailDue { generation }).await;
        });
    }

    fn fetch_selected_mail(&mut self, key: ReadingKey) {
        self.outlook.mail_fetch_generation = self.outlook.mail_fetch_generation.wrapping_add(1);
        if self.outlook.pending_reading.as_ref() == Some(&key) {
            return;
        }
        self.outlook.pending_reading = Some(key.clone());
        match key {
            ReadingKey::Thread(conversation_id) => self.load_mail_thread(conversation_id),
            ReadingKey::Message(id) => {
                self.status = "loading message…".into();
                self.load_body(id);
            }
        }
    }

    fn show_message(&mut self, message: MailMessage) {
        let rendered = render_mail_body(&message);
        self.outlook.reading_links = rendered.links;
        self.outlook.reading_images = rendered
            .images
            .into_iter()
            .map(|image| MailBodyImage {
                message_id: message.id.clone(),
                image,
            })
            .collect();
        self.outlook.reading_body = Some(rendered.text);
        self.outlook.reading_thread.clear();
        self.outlook.reading_thread_bodies.clear();
        self.show_cached_attachments(&message);
        if self.graphics.is_some()
            && self
                .outlook
                .reading_images
                .iter()
                .any(|image| mail_content_id(&image.image.src).is_some())
        {
            self.ensure_attachment_list(message.id.clone());
        }
        self.outlook.reading_scroll = 0;
        self.outlook.reading = Some(message);
        self.schedule_current_read_timer();
        self.fetch_mail_images();
    }

    fn show_thread(&mut self, items: Vec<MailMessage>, truncated: bool) {
        let Some(latest) = items.first().cloned() else {
            self.status = "thread has no messages".into();
            return;
        };
        let rendered: Vec<content::RenderedBody> = items.iter().map(render_mail_body).collect();
        self.outlook.reading_links = rendered
            .first()
            .map(|body| body.links.clone())
            .unwrap_or_default();
        self.outlook.reading_body = None;
        self.outlook.reading_images = items
            .iter()
            .zip(&rendered)
            .flat_map(|(message, body)| {
                body.images.iter().cloned().map(|image| MailBodyImage {
                    message_id: message.id.clone(),
                    image,
                })
            })
            .collect();
        self.outlook.reading_thread_bodies = rendered.into_iter().map(|body| body.text).collect();
        self.outlook.reading_thread = items;
        self.show_cached_attachments(&latest);
        if self.graphics.is_some() {
            let image_message_ids: std::collections::HashSet<String> = self
                .outlook
                .reading_images
                .iter()
                .filter(|image| mail_content_id(&image.image.src).is_some())
                .map(|image| image.message_id.clone())
                .collect();
            for message_id in image_message_ids {
                self.ensure_attachment_list(message_id);
            }
        }
        self.outlook.reading_scroll = 0;
        self.outlook.reading = Some(latest);
        self.schedule_current_read_timer();
        self.fetch_mail_images();
        self.status = if truncated {
            format!(
                "{} thread messages loaded (more not shown)",
                self.outlook.reading_thread.len()
            )
        } else {
            format!(
                "{} thread messages loaded",
                self.outlook.reading_thread.len()
            )
        };
    }

    fn show_cached_attachments(&mut self, message: &MailMessage) {
        self.outlook.reading_attachments = self
            .outlook
            .attachment_cache
            .get(&message.id)
            .cloned()
            .unwrap_or_default();
        if message.has_attachments.unwrap_or(false) {
            self.ensure_attachment_list(message.id.clone());
        }
    }

    fn ensure_attachment_list(&mut self, message_id: String) {
        if !self.outlook.attachment_cache.contains_key(&message_id)
            && self.outlook.attachment_loading.insert(message_id.clone())
        {
            self.load_attachments(message_id);
        }
    }

    fn cache_mail_contacts(&mut self, messages: &[MailMessage]) {
        for message in messages {
            for recipient in message
                .from
                .iter()
                .chain(message.to_recipients.iter())
                .chain(message.cc_recipients.iter())
                .chain(message.bcc_recipients.iter())
            {
                cache_mail_contact(&mut self.mail_contacts, recipient);
            }
        }
    }

    pub(crate) fn compose_suggestions(&self, compose: &Compose) -> Vec<MailContact> {
        let Some(input) = compose.active_recipient() else {
            return Vec::new();
        };
        let (_, _, query) = recipient_token(input);
        let query = query.to_lowercase();
        if query.is_empty() {
            return Vec::new();
        }

        let mut matches: Vec<(bool, MailContact)> = self
            .mail_contacts
            .values()
            .filter_map(|contact| {
                let address = contact.address.to_lowercase();
                let name = contact.name.as_deref().unwrap_or_default().to_lowercase();
                (address.contains(&query) || name.contains(&query)).then(|| {
                    (
                        address.starts_with(&query) || name.starts_with(&query),
                        contact.clone(),
                    )
                })
            })
            .collect();
        matches.sort_by(|(a_prefix, a), (b_prefix, b)| {
            b_prefix
                .cmp(a_prefix)
                .then_with(|| a.name.cmp(&b.name))
                .then_with(|| a.address.cmp(&b.address))
        });
        matches
            .into_iter()
            .map(|(_, contact)| contact)
            .take(5)
            .collect()
    }

    fn mail_selection_key(&self, message: &MailMessage) -> String {
        if self.outlook.threaded {
            message
                .conversation_id
                .clone()
                .unwrap_or_else(|| message.id.clone())
        } else {
            message.id.clone()
        }
    }

    fn rebuild_mail_rows(&mut self) {
        self.outlook.message_rows = mail_row_indices(&self.outlook.messages, self.outlook.threaded);
    }

    fn toggle_mail_threads(&mut self) {
        let selected = self.current_mail().cloned();
        self.outlook.threaded = !self.outlook.threaded;
        self.rebuild_mail_rows();
        self.outlook.msg_sel = selected
            .as_ref()
            .and_then(|selected| {
                self.outlook.message_rows.iter().position(|&i| {
                    self.outlook.messages.get(i).is_some_and(|message| {
                        if self.outlook.threaded {
                            mail_conversation_key(message) == mail_conversation_key(selected)
                        } else {
                            message.id == selected.id
                        }
                    })
                })
            })
            .unwrap_or(0)
            .min(self.outlook.message_rows.len().saturating_sub(1));
        self.status = if self.outlook.threaded {
            "mail: threads".into()
        } else {
            "mail: individual messages".into()
        };
        if self.settings.preview_mail_on_hover && self.outlook_focus == OutlookFocus::Messages {
            self.open_current_mail();
        }
    }

    fn toggle_mail_read(&mut self) {
        let message = if self.outlook_focus == OutlookFocus::Reading {
            self.outlook
                .reading
                .as_ref()
                .or_else(|| self.current_mail())
        } else {
            self.current_mail()
        };
        let Some(m) = message else {
            self.status = "select a message first".into();
            return;
        };
        let id = m.id.clone();
        let read = !m.is_read.unwrap_or(false);
        self.set_mail_read(&id, read);
    }

    fn open_move_mail(&mut self) {
        if self.current_mail().is_none() {
            self.status = "select a message first".into();
            return;
        }
        if self.outlook.folders.is_empty() {
            self.status = "folders still loading".into();
            return;
        }
        self.overlay = Some(Overlay::MoveMail {
            sel: self.outlook.folder_sel,
        });
    }

    fn trash_current_mail(&mut self) {
        let (id, label) = self
            .outlook
            .folders
            .iter()
            .find(|f| {
                matches!(
                    f.display_name.as_deref(),
                    Some("Deleted Items") | Some("Trash")
                )
            })
            .map(|f| {
                (
                    f.id.clone(),
                    f.display_name
                        .clone()
                        .unwrap_or_else(|| "Deleted Items".into()),
                )
            })
            .unwrap_or_else(|| ("deleteditems".into(), "Deleted Items".into()));
        self.move_current_mail(id, label);
    }

    fn move_current_mail(&mut self, destination_id: String, dest_label: String) {
        let Some(m) = self.current_mail() else {
            self.status = "select a message first".into();
            return;
        };
        if self
            .outlook
            .folders
            .get(self.outlook.folder_sel)
            .is_some_and(|f| f.id == destination_id)
        {
            self.status = format!("already in {dest_label}");
            return;
        }
        let id = m.id.clone();
        let sel = self.outlook.msg_sel;
        if let Some(raw) = self.outlook.message_rows.get(sel).copied() {
            self.outlook.messages.remove(raw);
        }
        self.rebuild_mail_rows();
        self.outlook.msg_sel = sel.min(self.outlook.message_rows.len().saturating_sub(1));
        if self.outlook.reading.as_ref().is_some_and(|r| r.id == id) {
            self.outlook.reading = None;
            self.outlook.reading_body = None;
            self.outlook.reading_thread.clear();
            self.outlook.reading_thread_bodies.clear();
            self.outlook.reading_links.clear();
            self.outlook.reading_attachments.clear();
            self.outlook.reading_scroll = 0;
            if self.outlook_focus == OutlookFocus::Reading {
                self.outlook_focus = OutlookFocus::Messages;
            }
        }
        self.status = format!("moving to {dest_label}…");
        let s = self.session.clone();
        self.spawn(async move {
            mail::move_message(&s.graph, &id, &destination_id).await?;
            Ok(AppMessage::Done(format!("moved to {dest_label}")))
        });
    }

    fn populate_reply_recipients(&self, compose: &mut Compose) {
        if !matches!(
            compose.kind,
            ComposeKind::ReplyMail { .. } | ComposeKind::ReplyAllMail { .. }
        ) {
            return;
        }
        let source = self
            .outlook
            .reading
            .as_ref()
            .or_else(|| self.current_mail());
        let Some(source) = source else { return };
        let me = self
            .me
            .as_ref()
            .and_then(User::best_email)
            .unwrap_or_default();
        let sender = source.sender_address().unwrap_or_default();
        let mut to = vec![sender];
        let mut cc = Vec::new();
        if matches!(compose.kind, ComposeKind::ReplyAllMail { .. }) {
            to.extend(recipient_addresses(&source.to_recipients));
            cc.extend(recipient_addresses(&source.cc_recipients));
        }
        dedup_addresses(&mut to, me);
        dedup_addresses(&mut cc, me);
        compose.to = TextInput::from(to.join(", ").as_str());
        compose.cc = TextInput::from(cc.join(", ").as_str());
        compose.bcc.clear();
    }

    fn teams_jump_list(&mut self, to_end: bool) {
        match self.teams.mode {
            TeamsMode::Chats => {
                self.teams.chat_sel = if to_end {
                    self.teams_chat_row_count().saturating_sub(1)
                } else {
                    0
                };
                self.preview_selected_teams_chat();
            }
            TeamsMode::Channels => {
                if self.teams.channels.is_empty() {
                    self.teams.team_sel = if to_end {
                        self.teams.teams.len().saturating_sub(1)
                    } else {
                        0
                    };
                } else {
                    self.teams.channel_sel = if to_end {
                        self.teams.channels.len().saturating_sub(1)
                    } else {
                        0
                    };
                }
            }
        }
    }

    fn teams_update(&mut self, id: String, prepared: crate::images::Prepared) {
        let Some(original) = self.teams.messages.iter().find(|m| m.id == id).cloned() else {
            self.restore_teams_composer();
            self.status = "message to edit is no longer loaded".into();
            return;
        };
        let s = self.session.clone();
        self.status = "saving edit…".into();
        match self.teams.mode {
            TeamsMode::Chats => {
                let Some(chat_id) = self.teams.open_chat_id.clone() else {
                    self.restore_teams_composer();
                    return;
                };
                self.spawn(async move {
                    let body = prepared_body(&prepared);
                    match chats::update_message_with_body(&s.graph, &chat_id, &original, body).await
                    {
                        Ok(()) => Ok(AppMessage::TeamsSent("message updated".into())),
                        Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
                    }
                });
            }
            TeamsMode::Channels => {
                let Some((team_id, channel_id)) = self.teams.open_channel.clone() else {
                    self.restore_teams_composer();
                    return;
                };
                self.spawn(async move {
                    let body = prepared_body(&prepared);
                    match channels::update_message_with_body(
                        &s.graph,
                        &team_id,
                        &channel_id,
                        &original,
                        body,
                    )
                    .await
                    {
                        Ok(()) => Ok(AppMessage::TeamsSent("message updated".into())),
                        Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
                    }
                });
            }
        }
    }

    fn teams_edit_selected(&mut self) {
        self.teams_edit_index(self.teams.msg_sel);
    }

    fn teams_edit_last_own(&mut self) {
        let my_id = self.me.as_ref().map(|m| m.id.as_str());
        let Some(idx) = self.teams.messages.iter().rposition(|m| m.can_edit(my_id)) else {
            self.status = "no message of yours to edit".into();
            return;
        };
        self.teams.msg_sel = idx;
        self.teams_edit_index(idx);
    }

    fn teams_edit_index(&mut self, idx: usize) {
        let Some(m) = self.teams.messages.get(idx) else {
            return;
        };
        let my_id = self.me.as_ref().map(|u| u.id.as_str());
        if my_id.is_none() {
            self.status = "still loading your profile".into();
            return;
        }
        if !m.can_edit(my_id) {
            self.status = if m.deleted_date_time.is_some() {
                "can't edit a deleted message".into()
            } else if m.author_id() != my_id {
                "you can only edit your own messages".into()
            } else {
                "can't edit this message".into()
            };
            return;
        }
        if self.teams.sending {
            self.status = "already sending…".into();
            return;
        }
        if self.teams.editing_message.is_none()
            && (!self.teams.composer.is_empty() || !self.teams.images.is_empty())
        {
            self.status = "composer has unsent text — Esc first".into();
            return;
        }
        let text = m.editable_text();
        self.teams.editing_message = self
            .teams
            .open_chat_id
            .clone()
            .map(|chat| (chat, m.id.clone()));
        self.teams.replying_to = None;
        self.teams.composer = TextInput::from(text.as_str());
        self.teams.focus = TeamsFocus::Composer;
        self.status = "editing message — Enter to save · Esc to cancel".into();
    }

    fn restore_teams_composer(&mut self) {
        self.teams.sending = false;
        if let Some((text, images)) = self.teams.pending_restore.take() {
            self.teams.composer = TextInput::from(text.as_str());
            self.teams.images = images;
            self.rebuild_composer_previews();
        }
    }

    fn complete_teams_path(&mut self) -> bool {
        let Some(token) =
            crate::images::token_at(self.teams.composer.chars(), self.teams.composer.cursor())
        else {
            return false;
        };
        let Some(completion) = crate::images::complete(&token) else {
            return false;
        };
        self.teams
            .composer
            .replace_range(token.start, token.end, &completion.replacement);
        self.status = completion.status;
        true
    }

    fn complete_mail_path(&mut self, compose: &mut Compose) -> bool {
        let completion = if compose.field == COMPOSE_BODY {
            let Some(token) = crate::images::token_at(compose.body.chars(), compose.body.cursor())
            else {
                return false;
            };
            let Some(completion) = crate::images::complete(&token) else {
                return false;
            };
            compose
                .body
                .replace_range(token.start, token.end, &completion.replacement);
            completion
        } else if compose.field == COMPOSE_ATTACH {
            let Some(completion) =
                crate::images::complete_attachment(compose.attach.chars(), compose.attach.cursor())
            else {
                return false;
            };
            compose
                .attach
                .replace_range(0, compose.attach.chars().len(), &completion.replacement);
            completion
        } else {
            return false;
        };
        self.status = completion.status;
        true
    }

    fn paste_clipboard_image(&mut self) {
        match crate::clipboard::image_bytes() {
            Ok((_mime, bytes)) => {
                let ext = crate::images::image_extension(&bytes).unwrap_or("png");
                let name = crate::images::unique_paste_name(&self.teams.images, ext);
                match crate::images::InlineImage::from_bytes(name, bytes) {
                    Ok(img) => self.stage_teams_image(img),
                    Err(e) => self.status = e,
                }
            }
            Err(crate::clipboard::ImagePasteError::NoHelper) => {
                self.status =
                    "clipboard image paste needs wl-paste or xclip on this machine".into();
            }
            Err(crate::clipboard::ImagePasteError::NoImage) => {
                self.status =
                    "clipboard has no image — copy an image, or paste text with the terminal paste"
                        .into();
            }
            Err(crate::clipboard::ImagePasteError::Failed(e)) => {
                self.status = format!("clipboard: {e}");
            }
        }
    }

    fn stage_teams_image(&mut self, img: crate::images::InlineImage) {
        self.status = format!(
            "attached {} ({})",
            img.name,
            crate::images::human_size(img.bytes.len() as u64)
        );
        let preview = self.graphics.as_ref().and_then(|g| {
            g.decode(&img.bytes, crate::termimg::COMPOSER_IMAGE_ROWS)
                .ok()
        });
        self.teams.images.push(img);
        if self.graphics.is_some() {
            self.teams.composer_previews.push(preview);
        }
    }

    fn unstage_teams_image(&mut self) {
        if let Some(img) = self.teams.images.pop() {
            self.teams.composer_previews.pop();
            self.status = format!("removed {}", img.name);
        }
    }

    fn rebuild_composer_previews(&mut self) {
        self.teams.composer_previews = match &self.graphics {
            Some(g) => self
                .teams
                .images
                .iter()
                .map(|img| {
                    g.decode(&img.bytes, crate::termimg::COMPOSER_IMAGE_ROWS)
                        .ok()
                })
                .collect(),
            None => Vec::new(),
        };
    }

    fn fetch_message_images(&mut self) {
        if self.graphics.is_none() {
            return;
        }
        let chat = self.teams.open_chat_id.clone();
        let channel = self.teams.open_channel.clone();
        let jobs: Vec<(String, String)> = self
            .teams
            .messages
            .iter()
            .zip(self.teams.messages_images.iter())
            .flat_map(|(msg, imgs)| {
                let chat = chat.clone();
                let channel = channel.clone();
                imgs.iter().filter_map(move |img| {
                    let hosted = crate::termimg::hosted_content_id(&img.src);
                    let key = crate::termimg::cache_key(&img.src, hosted);
                    let path = crate::termimg::fetch_path(
                        &img.src,
                        hosted,
                        &msg.id,
                        chat.as_deref(),
                        channel.as_ref(),
                    )?;
                    Some((key, path))
                })
            })
            .collect();
        for (key, path) in jobs {
            if self.image_cache.contains_key(&key)
                || self.image_failed.contains(&key)
                || self.image_loading.contains(&key)
            {
                continue;
            }
            if let Some(root) = self.session.config.cache_dir.clone() {
                if let Some(bytes) = crate::teams_cache::load_image_bytes(&root, &key) {
                    self.image_loading.insert(key.clone());
                    let tx = self.tx.clone();
                    tokio::spawn(async move {
                        let _ = tx.send(AppMessage::HostedImage { key, bytes }).await;
                    });
                    continue;
                }
            }
            self.image_loading.insert(key.clone());
            let s = self.session.clone();
            self.spawn(async move {
                match s.graph.get_bytes(&path).await {
                    Ok(bytes) => Ok(AppMessage::HostedImage { key, bytes }),
                    Err(_) => Ok(AppMessage::HostedImageFailed { key }),
                }
            });
        }
    }

    fn fetch_user_photos(&mut self) {
        if self.graphics.is_none() {
            return;
        }
        let mut ids: Vec<String> = self
            .teams
            .messages
            .iter()
            .filter_map(|m| m.author_id().map(str::to_string))
            .filter(|id| !id.is_empty())
            .collect();
        ids.sort();
        ids.dedup();
        for id in ids {
            let key = crate::termimg::photo_cache_key(&id);
            if self.image_cache.contains_key(&key)
                || self.image_failed.contains(&key)
                || self.image_loading.contains(&key)
            {
                continue;
            }
            self.image_loading.insert(key.clone());
            let s = self.session.clone();
            self.spawn(async move {
                match people::photo(&s.graph, &id).await {
                    Ok(bytes) => Ok(AppMessage::HostedImage { key, bytes }),
                    Err(_) => Ok(AppMessage::HostedImageFailed { key }),
                }
            });
        }
    }

    fn selected_message_image_keys(&self) -> Vec<String> {
        let Some(imgs) = self.teams.messages_images.get(self.teams.msg_sel) else {
            return Vec::new();
        };
        imgs.iter()
            .map(|img| {
                let hosted = crate::termimg::hosted_content_id(&img.src);
                crate::termimg::cache_key(&img.src, hosted)
            })
            .collect()
    }

    fn open_full_image(&mut self) {
        if self.graphics.is_none() {
            self.status = "this terminal cannot display images".into();
            return;
        }
        let keys = self.selected_message_image_keys();
        if keys.is_empty() {
            self.status = "no image in this message".into();
            return;
        }
        self.overlay = Some(Overlay::ViewImage { keys, sel: 0 });
    }

    pub fn reading_image_entries(&self) -> Vec<(String, String)> {
        reading_mail_image_entries(&self.outlook.reading_images)
    }

    fn reading_image_keys(&self) -> Vec<String> {
        self.reading_image_entries()
            .into_iter()
            .map(|(key, _)| key)
            .collect()
    }

    fn open_mail_images(&mut self) {
        if self.graphics.is_none() {
            self.status = "this terminal cannot display images".into();
            return;
        }
        if self.reading_image_keys().is_empty() {
            self.status = "no images in this message".into();
            return;
        }
        self.overlay = Some(Overlay::MailImages {
            scroll: 0,
            fullscreen: None,
        });
    }

    fn fetch_mail_images(&mut self) {
        if self.graphics.is_none() {
            return;
        }
        let images = self.outlook.reading_images.clone();
        for mail_image in images {
            let key = crate::termimg::mail_cache_key(&mail_image.message_id, &mail_image.image.src);
            if self.image_cache.contains_key(&key)
                || self.image_failed.contains(&key)
                || self.image_loading.contains(&key)
            {
                continue;
            }

            if let Some(decoded) = crate::images::from_data_uri(&mail_image.image.src) {
                match decoded.ok().and_then(|image| {
                    self.graphics.as_ref().and_then(|graphics| {
                        graphics
                            .decode(&image.bytes, crate::termimg::MAIL_IMAGE_ROWS)
                            .ok()
                    })
                }) {
                    Some(image) => {
                        self.image_cache.insert(key, image);
                    }
                    None => {
                        self.image_failed.insert(key);
                    }
                }
                continue;
            }

            if let Some(path) = graph_attachment_path(&mail_image.image.src) {
                self.image_loading.insert(key.clone());
                let session = self.session.clone();
                self.spawn(async move {
                    match session.graph.get_bytes(&path).await {
                        Ok(bytes) => Ok(AppMessage::HostedImage { key, bytes }),
                        Err(_) => Ok(AppMessage::HostedImageFailed { key }),
                    }
                });
                continue;
            }

            let Some(content_id) = mail_content_id(&mail_image.image.src) else {
                // Remote tracking pixels / third-party images: do not fetch, and
                // do not reserve a slot in the reading pane.
                continue;
            };
            let Some(attachments) = self.outlook.attachment_cache.get(&mail_image.message_id)
            else {
                continue;
            };
            let attachment = attachments
                .iter()
                .find(|attachment| cid_matches(attachment, content_id));
            let Some(attachment) = attachment else {
                self.image_failed.insert(key);
                continue;
            };
            if let Some(bytes) = attachment
                .content_bytes
                .as_deref()
                .and_then(|b| m365_core::util::base64_decode(b).ok())
            {
                match self
                    .graphics
                    .as_ref()
                    .and_then(|g| g.decode(&bytes, crate::termimg::MAIL_IMAGE_ROWS).ok())
                {
                    Some(image) => {
                        self.image_cache.insert(key, image);
                    }
                    None => {
                        self.image_failed.insert(key);
                    }
                }
                continue;
            }
            let message_id = mail_image.message_id;
            let attachment_id = attachment.id.clone();
            self.image_loading.insert(key.clone());
            let session = self.session.clone();
            self.spawn(async move {
                match mail::download_attachment(&session.graph, &message_id, &attachment_id).await {
                    Ok(bytes) => Ok(AppMessage::HostedImage { key, bytes }),
                    Err(_) => Ok(AppMessage::HostedImageFailed { key }),
                }
            });
        }
    }

    fn on_key_mail_images(&mut self, key: KeyEvent, mut scroll: u16, fullscreen: Option<usize>) {
        let keys = self.reading_image_keys();
        let n = keys.len();
        let put = |scroll, fullscreen| Overlay::MailImages { scroll, fullscreen };
        if let Some(sel) = fullscreen {
            match key.code {
                KeyCode::Char(c @ '1'..='9') => {
                    let idx = (c as u8 - b'1') as usize;
                    self.overlay = Some(put(scroll, if idx < n { Some(idx) } else { Some(sel) }));
                }
                KeyCode::Char('j') | KeyCode::Down | KeyCode::Char('l') | KeyCode::Right => {
                    let sel = if n == 0 { 0 } else { (sel + 1) % n };
                    self.overlay = Some(put(scroll, Some(sel)));
                }
                KeyCode::Char('k') | KeyCode::Up | KeyCode::Char('h') | KeyCode::Left => {
                    let sel = if n == 0 { 0 } else { (sel + n - 1) % n };
                    self.overlay = Some(put(scroll, Some(sel)));
                }
                _ => self.overlay = Some(put(scroll, Some(sel))),
            }
            return;
        }
        match key.code {
            KeyCode::Char(c @ '1'..='9') => {
                let idx = (c as u8 - b'1') as usize;
                self.overlay = Some(put(scroll, if idx < n { Some(idx) } else { None }));
            }
            KeyCode::Char('j') | KeyCode::Down => {
                scroll = scroll.saturating_add(1).min(self.images_max_scroll.get());
                self.overlay = Some(put(scroll, None));
            }
            KeyCode::Char('k') | KeyCode::Up => {
                scroll = scroll.saturating_sub(1);
                self.overlay = Some(put(scroll, None));
            }
            KeyCode::PageDown => {
                scroll = scroll.saturating_add(10).min(self.images_max_scroll.get());
                self.overlay = Some(put(scroll, None));
            }
            KeyCode::PageUp => {
                scroll = scroll.saturating_sub(10);
                self.overlay = Some(put(scroll, None));
            }
            KeyCode::Home | KeyCode::Char('g') if !key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.overlay = Some(put(0, None));
            }
            KeyCode::End | KeyCode::Char('G') => {
                self.overlay = Some(put(self.images_max_scroll.get(), None));
            }
            KeyCode::Char('g') if key.modifiers.contains(KeyModifiers::SHIFT) => {
                self.overlay = Some(put(self.images_max_scroll.get(), None));
            }
            _ => self.overlay = Some(put(scroll, None)),
        }
    }

    fn run_command(&mut self, id: &str) {
        self.overlay = None;
        match id {
            "outlook" => self.switch_screen(Screen::Outlook),
            "teams" => self.switch_screen(Screen::Teams),
            "calendar" => self.switch_screen(Screen::Calendar),
            "compose" => {
                self.overlay = Some(Overlay::Compose(empty_compose()));
            }
            "reply" => self.open_reply(ReplyMode::Reply),
            "reply-all" => self.open_reply(ReplyMode::ReplyAll),
            "forward" => self.open_reply(ReplyMode::Forward),
            "mark-read" => self.toggle_mail_read(),
            "move-mail" => self.open_move_mail(),
            "trash" => self.trash_current_mail(),
            "layout" => self.toggle_pane_layout(),
            "threads" => self.toggle_mail_threads(),
            "chat-sender" => {
                if let Some(addr) = self.current_mail().and_then(|m| m.sender_address()) {
                    self.status = format!("opening chat with {addr}…");
                    self.open_chat_with_email(addr);
                } else {
                    self.status = "select an email first".into();
                }
            }
            "new-chat" => {
                self.screen = Screen::Teams;
                self.teams.mode = TeamsMode::Chats;
                self.show_new_chat();
            }
            "refresh" => self.refresh_current(),
            "diagnostics" => self.open_diagnostics(),
            "help" => {
                self.help_scroll = 0;
                self.overlay = Some(Overlay::Help);
            }
            "quit" => self.should_quit = true,
            _ => {}
        }
    }

    pub fn compose_edit_snapshot(&self) -> Option<(String, String)> {
        match &self.overlay {
            Some(Overlay::Compose(c)) => Some((c.subject.text(), c.body.text())),
            _ => None,
        }
    }

    pub fn apply_compose_edit(&mut self, subject: Option<String>, body: String) {
        let Some(Overlay::Compose(c)) = &mut self.overlay else {
            return;
        };
        if let Some(subject) = subject {
            if matches!(c.kind, ComposeKind::NewMail) {
                c.subject = TextInput::from(subject.as_str());
            }
        }
        c.body = TextInput::from(body.as_str());
        c.field = COMPOSE_BODY;
        c.ctrl_x = false;
        self.status = "edited in $EDITOR".into();
    }
    fn teams_send_prepared(&mut self) {
        if self.teams.sending {
            self.status = "already sending…".into();
            return;
        }
        let text = self.teams.composer.text();
        let staged = self.teams.images.clone();
        let prepared = match crate::images::prepare_message(
            &text,
            &staged,
            self.session.config.teams_markdown,
        ) {
            Ok(Some(p)) => p,
            Ok(None) => return,
            Err(e) => {
                self.status = e;
                return;
            }
        };

        let n_images = match &prepared {
            crate::images::Prepared::Text(_) => 0,
            crate::images::Prepared::Html { images, .. } => images.len(),
        };
        self.teams.pending_restore = Some((text, staged));
        self.teams.composer.clear();
        self.teams.images.clear();
        self.teams.composer_previews.clear();
        self.teams.sending = true;
        if let Some(id) = self.teams.open_chat_id.clone() {
            self.mark_teams_chat_active(&id);
        }
        let editing = self
            .teams
            .editing_message
            .as_ref()
            .map(|(_, id)| id.clone());
        let replying_to = self.teams.replying_to.take();
        self.status = match n_images {
            0 => "sending…".into(),
            1 => "sending image…".into(),
            n => format!("sending {n} images…"),
        };

        if let Some(id) = editing {
            self.teams_update(id, prepared);
            return;
        }

        match (replying_to, self.teams.mode) {
            // Replying in a chat: quote the original in the body, which is how
            // Teams itself represents a chat reply.
            (Some(idx), TeamsMode::Chats) => {
                let (Some(chat_id), Some(original)) = (
                    self.teams.open_chat_id.clone(),
                    self.teams.messages.get(idx),
                ) else {
                    self.restore_teams_composer();
                    return;
                };
                let author = original.author();
                let original = original.clone();
                let s = self.session.clone();
                self.status = format!("replying to {author}…");
                self.spawn(async move {
                    let body = prepared_body(&prepared);
                    match chats::send_reply_with_body(&s.graph, &chat_id, &original, body).await {
                        Ok(message) => Ok(AppMessage::ChatMessageSent {
                            chat_id,
                            message_id: message.id,
                            status: "reply sent".into(),
                        }),
                        Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
                    }
                });
            }
            // Channels have a real replies collection, so the reply threads.
            (Some(idx), TeamsMode::Channels) => {
                let (Some((team_id, channel_id)), Some(original)) = (
                    self.teams.open_channel.clone(),
                    self.teams.messages.get(idx),
                ) else {
                    self.restore_teams_composer();
                    return;
                };
                let message_id = original.id.clone();
                let s = self.session.clone();
                self.status = "replying…".into();
                self.spawn(async move {
                    let body = prepared_body(&prepared);
                    match channels::send_reply_with_body(
                        &s.graph,
                        &team_id,
                        &channel_id,
                        &message_id,
                        body,
                    )
                    .await
                    {
                        Ok(_) => Ok(AppMessage::TeamsSent("reply sent".into())),
                        Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
                    }
                });
            }
            (None, TeamsMode::Chats) => {
                if let Some(id) = self.teams.open_chat_id.clone() {
                    self.send_chat_message_prepared(id, prepared);
                } else {
                    self.restore_teams_composer();
                }
            }
            (None, TeamsMode::Channels) => {
                if let Some((t, c)) = self.teams.open_channel.clone() {
                    self.send_channel_message(t, c, prepared);
                } else {
                    self.restore_teams_composer();
                }
            }
        }
    }

    fn send_chat_message_prepared(&self, chat_id: String, prepared: crate::images::Prepared) {
        let s = self.session.clone();
        self.spawn(async move {
            let body = prepared_body(&prepared);
            match chats::send_message_with_body(&s.graph, &chat_id, body).await {
                Ok(message) => Ok(AppMessage::ChatMessageSent {
                    chat_id,
                    message_id: message.id,
                    status: "message sent".into(),
                }),
                Err(e) => Ok(AppMessage::TeamsSendFailed(format!("{e:#}"))),
            }
        });
    }
}

/// Fold a freshly-fetched newest page into an existing newest-first list:
/// the fresh page wins, older already-loaded entries are kept, duplicates go.
fn merge_newest_first<T, F, K>(fresh: Vec<T>, existing: Vec<T>, key: F) -> Vec<T>
where
    F: Fn(&T) -> K,
    K: std::hash::Hash + Eq,
{
    let seen: std::collections::HashSet<K> = fresh.iter().map(&key).collect();
    fresh
        .into_iter()
        .chain(existing.into_iter().filter(|e| !seen.contains(&key(e))))
        .collect()
}

/// Flatten rendered lines to plain text for the clipboard.
fn lines_to_plain(lines: &[ratatui::text::Line<'_>]) -> String {
    lines
        .iter()
        .map(|l| {
            l.spans
                .iter()
                .map(|s| s.content.as_ref())
                .collect::<String>()
                .trim_end()
                .to_string()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Shorten a URL for the status line.
fn truncate_url(url: &str) -> String {
    if url.chars().count() <= 60 {
        return url.to_string();
    }
    let head: String = url.chars().take(57).collect();
    format!("{head}…")
}

fn write_dotenv_assignment(path: &std::path::Path, key: &str, value: &str) -> anyhow::Result<()> {
    let original =
        std::fs::read_to_string(path).with_context(|| format!("reading {}", path.display()))?;
    let had_trailing_newline = original.ends_with('\n');
    let mut output = Vec::new();
    let mut replaced = false;

    for line in original.lines() {
        let trimmed = line.trim_start();
        let assignment = trimmed.strip_prefix("export ").unwrap_or(trimmed);
        let matches_key = assignment
            .split_once('=')
            .is_some_and(|(name, _)| name.trim() == key);

        if !matches_key {
            output.push(line.to_string());
            continue;
        }
        if replaced {
            continue;
        }

        let indent_len = line.len().saturating_sub(trimmed.len());
        let indent = &line[..indent_len];
        let export_prefix = if trimmed.starts_with("export ") {
            "export "
        } else {
            ""
        };
        output.push(format!("{indent}{export_prefix}{key}={value}"));
        replaced = true;
    }

    if !replaced {
        if output.last().is_some_and(|line| !line.is_empty()) {
            output.push(String::new());
        }
        output.push(format!("{key}={value}"));
    }

    let mut rendered = output.join("\n");
    if had_trailing_newline || !rendered.is_empty() {
        rendered.push('\n');
    }

    std::fs::write(path, rendered).with_context(|| format!("writing {}", path.display()))
}

/// Expand a leading `~` to the user's home directory.
fn expand_tilde(p: &str) -> std::path::PathBuf {
    if let Some(rest) = p.strip_prefix("~/") {
        if let Some(home) = std::env::var_os("HOME") {
            return std::path::PathBuf::from(home).join(rest);
        }
    }
    std::path::PathBuf::from(p)
}

fn human_size(bytes: u64) -> String {
    if bytes >= 1024 * 1024 {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    } else if bytes >= 1024 {
        format!("{:.0} KB", bytes as f64 / 1024.0)
    } else {
        format!("{bytes} B")
    }
}

/// Resident set size in KiB, from `/proc/self/status` (Linux).
fn read_rss_kb() -> Option<u64> {
    parse_vmrss(&std::fs::read_to_string("/proc/self/status").ok()?)
}

fn parse_vmrss(status: &str) -> Option<u64> {
    status
        .lines()
        .find_map(|l| l.strip_prefix("VmRSS:"))?
        .split_whitespace()
        .next()?
        .parse()
        .ok()
}

/// Prefer the learned 1:1 contact name for notification titles.
///
/// Automation messages posted "as user" carry our own identity as the sender,
/// so `Chat::label()` can otherwise fall all the way back to the raw chat type
/// (`oneOnOne`). The learned contact name is the actual peer/notifier name.
fn teams_notification_label(
    learned_name: Option<&str>,
    chat: &Chat,
    me_id: Option<&str>,
) -> String {
    learned_name
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .map(str::to_string)
        .unwrap_or_else(|| chat.label(me_id))
}

/// Return the first non-empty displayName found on any message from the peer.
///
/// Some external / Microsoft personal-account messages omit displayName while
/// later messages from the same peer include it. Do not let the first peer
/// message without a name prevent the application from learning a usable name.
fn peer_display_name_from_messages<'a>(
    messages: &'a [ChatMessage],
    me_id: Option<&str>,
    expected_peer_id: Option<&str>,
) -> Option<&'a str> {
    let me_id = me_id?;

    messages
        .iter()
        .filter_map(|message| message.from.as_ref()?.user.as_ref())
        .filter(|user| {
            if let Some(expected_peer_id) = expected_peer_id {
                user.id.as_deref() == Some(expected_peer_id)
            } else {
                user.id.as_deref() != Some(me_id)
            }
        })
        .filter_map(|user| user.display_name.as_deref())
        .map(str::trim)
        .find(|name| !name.is_empty())
}

fn parse_teams_slash_command(text: &str) -> Result<Option<TeamsSlashCommand>, String> {
    let text = text.trim();
    if !text.starts_with('/') {
        return Ok(None);
    }

    let split = text.find(char::is_whitespace).unwrap_or(text.len());
    let command = text[..split].to_ascii_lowercase();
    let argument = text[split..].trim();

    let need_argument = |usage: &str| {
        if argument.is_empty() {
            Err(format!("usage: {usage}"))
        } else {
            Ok(argument.to_string())
        }
    };

    match command.as_str() {
        "/topic" => Ok(Some(TeamsSlashCommand::Topic(need_argument(
            "/topic <chat topic>",
        )?))),
        "/rename" => Ok(Some(TeamsSlashCommand::Rename(need_argument(
            "/rename <new chat name>",
        )?))),
        "/edit" => Ok(Some(TeamsSlashCommand::Edit(need_argument(
            "/edit <new message text>",
        )?))),
        "/pin" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Pin)),
        "/unpin" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Unpin)),
        "/pins" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Pins)),
        "/unread" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Unread)),
        "/delete" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Delete)),
        "/ooo" if argument.is_empty() => Ok(Some(TeamsSlashCommand::OutOfOffice)),
        "/ooo" if argument.eq_ignore_ascii_case("set") => {
            Ok(Some(TeamsSlashCommand::OutOfOfficeSet))
        }
        "/ooo" if argument.eq_ignore_ascii_case("clear") => {
            Ok(Some(TeamsSlashCommand::OutOfOfficeClear))
        }
        "/help" if argument.is_empty() => Ok(Some(TeamsSlashCommand::Help)),
        "/help" if argument.eq_ignore_ascii_case("all") => Ok(Some(TeamsSlashCommand::HelpAll)),
        "/pin" | "/unpin" | "/pins" | "/unread" | "/delete" => {
            Err(format!("{command} does not take arguments"))
        }
        "/ooo" => Err("usage: /ooo [set|clear]".into()),
        "/help" => Err("usage: /help [all]".into()),
        _ => Err(format!("unknown Teams command {command}; use /help")),
    }
}

fn validate_chat_topic(topic: &str) -> Result<(), &'static str> {
    if topic.is_empty() {
        return Err("chat topic cannot be empty");
    }
    if topic.chars().count() > 250 {
        return Err("chat topic cannot exceed 250 characters");
    }
    if topic.contains(':') {
        return Err("chat topic cannot contain ':'");
    }
    Ok(())
}

fn validate_plain_message_edit(message: &ChatMessage) -> Result<(), &'static str> {
    if message.is_system_event() {
        return Err("system events cannot be edited");
    }
    if message.deleted_date_time.is_some() {
        return Err("deleted messages cannot be edited");
    }
    if !message.attachments.is_empty() || !message.mentions.is_empty() {
        return Err("editing messages with attachments or mentions is not supported yet");
    }
    Ok(())
}

/// Strict chronological key used when deciding whether a cached message falls
/// inside a freshly fetched authoritative window. Unlike `sort_key`, missing or
/// malformed creation times return None so uncertain old cache entries are kept
/// conservatively rather than deleted by guesswork.
fn authoritative_message_key(message: &ChatMessage) -> Option<(i64, u64)> {
    let created = message
        .created_date_time
        .as_deref()
        .and_then(|value| chrono::DateTime::parse_from_rfc3339(value).ok())?
        .timestamp_millis();
    let id = message.id.parse::<u64>().unwrap_or(0);
    Some((created, id))
}

/// Merge a freshly fetched newest server window with previously cached history.
///
/// The fresh page is authoritative for every message at or newer than its oldest
/// trustworthy key. Cached messages in that range that are absent from the fresh
/// page are stale and are dropped. Cached messages older than the boundary are
/// retained because they may belong to pages that were not fetched this time.
/// Fresh deleted tombstones remove matching cached messages but are not retained.
fn reconcile_authoritative_message_window(
    fresh: Vec<ChatMessage>,
    existing: Vec<ChatMessage>,
) -> Vec<ChatMessage> {
    if fresh.is_empty() {
        return Vec::new();
    }

    let fresh_ids: std::collections::HashSet<String> =
        fresh.iter().map(|message| message.id.clone()).collect();
    let oldest_fresh = fresh.iter().filter_map(authoritative_message_key).min();

    let mut merged: Vec<ChatMessage> = existing
        .into_iter()
        .filter(|message| {
            if message.deleted_date_time.is_some() || fresh_ids.contains(&message.id) {
                return false;
            }

            match (oldest_fresh, authoritative_message_key(message)) {
                (Some(boundary), Some(key)) => key < boundary,
                _ => true,
            }
        })
        .collect();

    merged.extend(
        fresh
            .into_iter()
            .filter(|message| message.deleted_date_time.is_none()),
    );
    merged.sort_by_key(sort_key);
    merged
}

/// Chronological sort key for a Teams message: creation time, then id.
fn sort_key(m: &ChatMessage) -> (i64, u64) {
    let at = m
        .created_date_time
        .as_deref()
        .and_then(|t| chrono::DateTime::parse_from_rfc3339(t).ok())
        .map(|d| d.timestamp_millis())
        .unwrap_or(0);
    (at, m.id.parse::<u64>().unwrap_or(0))
}

/// Split a comma/semicolon/space-separated recipients string into addresses.
fn parse_recipients(s: &str) -> Vec<String> {
    s.split([',', ';', ' '])
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect()
}

fn recipient_addresses(recipients: &[Recipient]) -> Vec<String> {
    recipients
        .iter()
        .filter_map(|recipient| recipient.email_address.as_ref()?.address.clone())
        .collect()
}

fn dedup_addresses(addresses: &mut Vec<String>, excluded: &str) {
    let mut seen = std::collections::HashSet::new();
    addresses.retain(|address| {
        let key = address.trim().to_lowercase();
        !key.is_empty() && !key.eq_ignore_ascii_case(excluded) && seen.insert(key)
    });
}

fn cache_mail_contact(
    contacts: &mut std::collections::BTreeMap<String, MailContact>,
    recipient: &Recipient,
) {
    let Some(email) = recipient.email_address.as_ref() else {
        return;
    };
    let Some(address) = email
        .address
        .as_deref()
        .map(str::trim)
        .filter(|s| !s.is_empty())
    else {
        return;
    };
    let name = email
        .name
        .as_deref()
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .map(str::to_string);
    let key = address.to_lowercase();
    contacts
        .entry(key)
        .and_modify(|contact| {
            if contact.name.is_none() && name.is_some() {
                contact.name = name.clone();
            }
        })
        .or_insert_with(|| MailContact {
            name,
            address: address.to_string(),
        });
}

fn recipient_token(input: &TextInput) -> (usize, usize, String) {
    let chars = input.chars();
    let cursor = input.cursor();
    let delimiter = |character: &char| matches!(character, ',' | ';');
    let raw_start = chars[..cursor]
        .iter()
        .rposition(delimiter)
        .map(|index| index + 1)
        .unwrap_or(0);
    let start = raw_start
        + chars[raw_start..cursor]
            .iter()
            .take_while(|character| character.is_whitespace())
            .count();
    let end = chars[cursor..]
        .iter()
        .position(delimiter)
        .map(|index| cursor + index)
        .unwrap_or(chars.len());
    let query = chars[start..cursor].iter().collect::<String>();
    (start, end, query.trim().to_string())
}

fn accept_recipient_suggestion(compose: &mut Compose, contact: &MailContact) {
    let Some(input) = compose.active_recipient() else {
        return;
    };
    let (start, end, _) = recipient_token(input);
    let input = compose.active_mut();
    input.replace_range(start, end, &format!("{}, ", contact.address));
    compose.suggestion_sel = 0;
}

/// Given the allowed field ids and the current one, return the next (wrapping).
fn next_field(fields: &[usize], current: usize) -> usize {
    let pos = fields.iter().position(|&f| f == current).unwrap_or(0);
    fields[(pos + 1) % fields.len()]
}

/// The previous allowed field (wrapping).
fn prev_field(fields: &[usize], current: usize) -> usize {
    let pos = fields.iter().position(|&f| f == current).unwrap_or(0);
    fields[(pos + fields.len() - 1) % fields.len()]
}

fn empty_compose() -> Compose {
    Compose::new(ComposeKind::NewMail, 0)
}

/// Temp-file format used when composing in `$EDITOR`: a `Subject:` header,
/// a blank line, then the body.
pub fn format_compose_file(subject: &str, body: &str) -> String {
    format!("Subject: {subject}\n\n{body}")
}

/// Parse the file written by [`format_compose_file`]. If the subject header is
/// missing, the whole file is the body and the subject is left unchanged.
pub fn parse_compose_file(s: &str) -> (Option<String>, String) {
    let s = s.replace('\r', "");
    let Some(rest) = s.strip_prefix("Subject:") else {
        return (None, s);
    };
    let (subject, rest) = match rest.split_once('\n') {
        Some((line, rest)) => (line.trim().to_string(), rest),
        None => (rest.trim().to_string(), ""),
    };
    let body = rest.strip_prefix('\n').unwrap_or(rest).to_string();
    (Some(subject), body)
}

fn prepared_body(prepared: &crate::images::Prepared) -> OutgoingBody<'_> {
    match prepared {
        crate::images::Prepared::Text(text) => OutgoingBody::Text(text),
        crate::images::Prepared::Html { html, images } => OutgoingBody::Html { html, images },
    }
}

fn mail_content_id(src: &str) -> Option<&str> {
    let src = src.trim();
    src.get(..4)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("cid:"))
        .then(|| src[4..].trim())
        .filter(|id| !id.is_empty())
}

fn normalize_content_id(id: &str) -> &str {
    id.trim().trim_start_matches('<').trim_end_matches('>')
}

fn cid_local_part(id: &str) -> &str {
    normalize_content_id(id).split('@').next().unwrap_or(id)
}

fn cid_matches(attachment: &Attachment, content_id: &str) -> bool {
    let want = normalize_content_id(content_id);
    let want_local = cid_local_part(content_id);
    if let Some(id) = attachment.content_id.as_deref() {
        let got = normalize_content_id(id);
        if got.eq_ignore_ascii_case(want) || cid_local_part(id).eq_ignore_ascii_case(want_local) {
            return true;
        }
    }
    attachment.name.as_deref().is_some_and(|name| {
        name.eq_ignore_ascii_case(want) || name.eq_ignore_ascii_case(want_local)
    })
}

/// Graph-hosted attachment `$value` URLs are safe to fetch with our token.
fn graph_attachment_path(src: &str) -> Option<String> {
    let src = src.trim();
    let rest = src
        .strip_prefix("https://graph.microsoft.com/v1.0/")
        .or_else(|| src.strip_prefix("https://graph.microsoft.com/beta/"))?;
    let path = rest.split('?').next().unwrap_or(rest);
    if path.contains("/attachments/") && path.contains("/$value") {
        Some(path.to_string())
    } else {
        None
    }
}

pub(crate) fn mail_image_is_local(src: &str) -> bool {
    let src = src.trim();
    mail_content_id(src).is_some()
        || src.to_ascii_lowercase().starts_with("data:image/")
        || graph_attachment_path(src).is_some()
}

pub(crate) fn reading_mail_image_entries(images: &[MailBodyImage]) -> Vec<(String, String)> {
    images
        .iter()
        .filter(|img| mail_image_is_local(&img.image.src))
        .map(|img| {
            (
                crate::termimg::mail_cache_key(&img.message_id, &img.image.src),
                img.image.alt.clone(),
            )
        })
        .collect()
}

fn render_mail_body(message: &MailMessage) -> content::RenderedBody {
    let (content_type, raw) = match &message.body {
        Some(body) => (
            body.content_type.as_deref(),
            body.content.clone().unwrap_or_default(),
        ),
        None => (
            Some("text"),
            message.body_preview.clone().unwrap_or_default(),
        ),
    };
    content::render_body(content_type, &raw)
}

fn mail_conversation_key(message: &MailMessage) -> &str {
    message.conversation_id.as_deref().unwrap_or(&message.id)
}

fn reading_key(message: &MailMessage, threaded: bool) -> ReadingKey {
    if threaded {
        ReadingKey::Thread(mail_conversation_key(message).to_string())
    } else {
        ReadingKey::Message(message.id.clone())
    }
}

/// Build the visible mail rows from a newest-first folder page.
fn mail_row_indices(messages: &[MailMessage], threaded: bool) -> Vec<usize> {
    if !threaded {
        return (0..messages.len()).collect();
    }
    let mut seen = std::collections::HashSet::new();
    messages
        .iter()
        .enumerate()
        .filter_map(|(i, message)| seen.insert(mail_conversation_key(message)).then_some(i))
        .collect()
}

pub(crate) fn mail_same_thread(a: &MailMessage, b: &MailMessage) -> bool {
    mail_conversation_key(a) == mail_conversation_key(b)
}

pub(crate) fn mail_thread_size(messages: &[MailMessage], message: &MailMessage) -> usize {
    let key = mail_conversation_key(message);
    messages
        .iter()
        .filter(|candidate| mail_conversation_key(candidate) == key)
        .count()
}

pub(crate) fn mail_row_unread(
    messages: &[MailMessage],
    message: &MailMessage,
    threaded: bool,
) -> bool {
    messages.iter().any(|candidate| {
        let belongs = if threaded {
            mail_same_thread(candidate, message)
        } else {
            candidate.id == message.id
        };
        belongs && !candidate.is_read.unwrap_or(true)
    })
}

/// Move a selection index by `delta`, clamped to `[0, len)`.
fn step(idx: usize, delta: i32, len: usize) -> usize {
    if len == 0 {
        return 0;
    }
    let max = len - 1;
    if delta < 0 {
        idx.saturating_sub((-delta) as usize)
    } else {
        (idx + delta as usize).min(max)
    }
}

pub const DEFAULT_FOLDER_PANEL_WIDTH: u16 = 26;
const MIN_FOLDER_PANEL_WIDTH: u16 = 11;
pub const MIN_FOLDER_PANEL_HEIGHT: u16 = 5;

pub(crate) fn mail_folder_label(folder: &MailFolder) -> String {
    let name = folder.display_name.clone().unwrap_or_default();
    match folder.unread_item_count.unwrap_or(0) {
        unread if unread > 0 => format!("{name} ({unread})"),
        _ => name,
    }
}

#[cfg(test)]
fn folder_panel_width(folders: &[MailFolder]) -> u16 {
    // Two border columns and one selected-row marker sit beside the label.
    folders
        .iter()
        .map(mail_folder_label)
        .map(|label| u16::try_from(Line::from(label).width()).unwrap_or(u16::MAX))
        .max()
        .unwrap_or(0)
        .saturating_add(3)
        .max(MIN_FOLDER_PANEL_WIDTH)
}

pub fn folder_panel_height(folders: &[MailFolder]) -> u16 {
    (folders.len().clamp(3, 12) as u16).saturating_add(2)
}

fn resize_panel_extent(current: u16, delta: i16, minimum: u16) -> u16 {
    if delta < 0 {
        current.saturating_sub(delta.unsigned_abs())
    } else {
        current.saturating_add(delta as u16)
    }
    .max(minimum)
}

/// Folder-list filter (case-insensitive substring on the rendered label).
pub fn filter_folders(folders: &[MailFolder], query: &str) -> Vec<usize> {
    let q = query.to_ascii_lowercase();
    folders
        .iter()
        .enumerate()
        .filter(|(_, folder)| {
            q.is_empty() || mail_folder_label(folder).to_ascii_lowercase().contains(&q)
        })
        .map(|(i, _)| i)
        .collect()
}

/// Palette fuzzy-ish filter (case-insensitive substring on label or id).
pub fn filter_commands(query: &str) -> Vec<(&'static str, &'static str)> {
    let q = query.to_ascii_lowercase();
    PALETTE_COMMANDS
        .iter()
        .filter(|(id, label)| {
            q.is_empty() || label.to_ascii_lowercase().contains(&q) || id.contains(&q)
        })
        .copied()
        .collect()
}

#[derive(Debug, Clone)]
pub struct MailBodyImage {
    pub message_id: String,
    pub image: content::BodyImage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum ReadingKey {
    Message(String),
    Thread(String),
}

#[cfg(test)]
pub(crate) fn test_app() -> App {
    use m365_core::config::*;
    let config = m365_core::Config {
        client_id: "cid".into(),
        tenant_id: "organizations".into(),
        scopes: vec!["User.Read".into(), "Mail.Send".into()],
        tunnel_base_url: None,
        redis_url: "redis://127.0.0.1:6379".into(),
        token_cache_path: std::env::temp_dir().join(format!(
            "m365-tui-test-no-token-{}.json",
            std::process::id()
        )),
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
        presence_activity_source: PresenceActivitySource::App,
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
    };
    let (tx, _rx) = mpsc::channel(32);
    App::new(Session::new(config), tx)
}

#[cfg(test)]
mod presence_renewal_tests {
    use super::*;

    #[test]
    fn available_session_renews_before_graph_activity_timeout() {
        let graph_available_timeout = std::time::Duration::from_secs(5 * 60);

        assert_eq!(
            presence_renew_after(SessionTarget::Present(AVAILABLE)),
            PRESENCE_AVAILABLE_RENEW_AFTER
        );
        assert!(PRESENCE_AVAILABLE_RENEW_AFTER < graph_available_timeout);
        assert_eq!(
            presence_renew_after(SessionTarget::Present(AWAY)),
            PRESENCE_RENEW_AFTER
        );
        assert_eq!(
            presence_renew_after(SessionTarget::Present(("Busy", "InACall"))),
            PRESENCE_RENEW_AFTER
        );
    }
}

#[cfg(test)]
mod persistent_cache_pruning_tests {
    use super::*;

    #[test]
    fn combined_image_and_conversation_cache_respects_one_limit() {
        let unique = format!(
            "m365-cache-prune-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        );
        let root = std::env::temp_dir().join(unique);
        let conversations = root.join(TEAMS_CONVERSATION_CACHE_DIR);
        std::fs::create_dir_all(&conversations).unwrap();

        let image_key = "11111111111111111111111111111111";
        std::fs::write(root.join(format!("{image_key}.meta")), vec![b'x'; 40]).unwrap();
        std::fs::write(root.join(format!("{image_key}-0.png")), vec![b'x'; 40]).unwrap();
        std::fs::write(
            conversations.join("22222222222222222222222222222222.json"),
            vec![b'y'; 80],
        )
        .unwrap();
        std::fs::write(
            conversations.join("33333333333333333333333333333333.json"),
            vec![b'z'; 80],
        )
        .unwrap();

        std::fs::write(root.join(UI_STATE_FILE), vec![b'u'; 300]).unwrap();
        std::fs::write(root.join(STAY_ARCHIVED_STATE_FILE), vec![b's'; 300]).unwrap();

        prune_persistent_cache(&root, 200).unwrap();

        let image_size = std::fs::read_dir(&root)
            .unwrap()
            .flatten()
            .filter_map(|entry| {
                let name = entry.file_name();
                let name = name.to_string_lossy();
                if !name.starts_with(image_key) {
                    return None;
                }
                entry
                    .metadata()
                    .ok()
                    .filter(|metadata| metadata.is_file())
                    .map(|metadata| metadata.len())
            })
            .sum::<u64>();
        let conversation_size = std::fs::read_dir(&conversations)
            .unwrap()
            .flatten()
            .filter_map(|entry| entry.metadata().ok())
            .filter(|metadata| metadata.is_file())
            .map(|metadata| metadata.len())
            .sum::<u64>();

        assert!(image_size + conversation_size <= 200);
        assert!(root.join(UI_STATE_FILE).exists());
        assert!(root.join(STAY_ARCHIVED_STATE_FILE).exists());

        let _ = std::fs::remove_dir_all(root);
    }
}

#[cfg(test)]
mod teams_slash_command_tests {
    use super::{parse_teams_slash_command, validate_chat_topic, TeamsSlashCommand};

    #[test]
    fn parses_commands_and_arguments() {
        assert_eq!(
            parse_teams_slash_command("/topic Release room"),
            Ok(Some(TeamsSlashCommand::Topic("Release room".into())))
        );
        assert_eq!(
            parse_teams_slash_command("/edit corrected text"),
            Ok(Some(TeamsSlashCommand::Edit("corrected text".into())))
        );
        assert_eq!(
            parse_teams_slash_command("/pin"),
            Ok(Some(TeamsSlashCommand::Pin))
        );
        assert_eq!(
            parse_teams_slash_command("/unpin"),
            Ok(Some(TeamsSlashCommand::Unpin))
        );
        assert_eq!(
            parse_teams_slash_command("/pins"),
            Ok(Some(TeamsSlashCommand::Pins))
        );
        assert_eq!(
            parse_teams_slash_command("/unread"),
            Ok(Some(TeamsSlashCommand::Unread))
        );
        assert_eq!(
            parse_teams_slash_command("/delete"),
            Ok(Some(TeamsSlashCommand::Delete))
        );
        assert_eq!(
            parse_teams_slash_command("/ooo"),
            Ok(Some(TeamsSlashCommand::OutOfOffice))
        );
        assert_eq!(
            parse_teams_slash_command("/ooo set"),
            Ok(Some(TeamsSlashCommand::OutOfOfficeSet))
        );
        assert_eq!(
            parse_teams_slash_command("/ooo CLEAR"),
            Ok(Some(TeamsSlashCommand::OutOfOfficeClear))
        );
        assert_eq!(
            parse_teams_slash_command("/help"),
            Ok(Some(TeamsSlashCommand::Help))
        );
        assert_eq!(
            parse_teams_slash_command("/help all"),
            Ok(Some(TeamsSlashCommand::HelpAll))
        );
        assert_eq!(
            parse_teams_slash_command("/rename Release room"),
            Ok(Some(TeamsSlashCommand::Rename("Release room".into())))
        );
        assert_eq!(parse_teams_slash_command("ordinary message"), Ok(None));
    }

    #[test]
    fn rejects_bad_command_shapes() {
        assert!(parse_teams_slash_command("/topic").is_err());
        assert!(parse_teams_slash_command("/rename").is_err());
        assert!(parse_teams_slash_command("/edit").is_err());
        assert!(parse_teams_slash_command("/pins extra").is_err());
        assert!(parse_teams_slash_command("/ooo forever").is_err());
        assert!(parse_teams_slash_command("/help nope").is_err());
        assert!(parse_teams_slash_command("/unknown").is_err());
    }

    #[test]
    fn validates_graph_topic_constraints() {
        assert!(validate_chat_topic("Release room").is_ok());
        assert!(validate_chat_topic("").is_err());
        assert!(validate_chat_topic("bad:topic").is_err());
        assert!(validate_chat_topic(&"x".repeat(251)).is_err());
    }
}

#[cfg(test)]
mod contact_name_learning_tests {
    use super::peer_display_name_from_messages;
    use m365_core::models::ChatMessage;

    fn message(id: &str, sender_id: &str, display_name: Option<&str>) -> ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": id,
            "from": {
                "user": {
                    "id": sender_id,
                    "displayName": display_name
                }
            }
        }))
        .expect("valid chat-message fixture")
    }

    #[test]
    fn learns_name_from_later_peer_message() {
        let messages = vec![
            message("1", "peer", None),
            message("2", "me", Some("Me")),
            message("3", "peer", Some("Personal Contact")),
        ];

        assert_eq!(
            peer_display_name_from_messages(&messages, Some("me"), None),
            Some("Personal Contact")
        );
    }

    #[test]
    fn ignores_empty_peer_names_and_self_names() {
        let messages = vec![
            message("1", "me", Some("Me")),
            message("2", "peer", Some("   ")),
            message("3", "peer", None),
        ];

        assert_eq!(
            peer_display_name_from_messages(&messages, Some("me"), None),
            None
        );
    }

    #[test]
    fn requires_signed_in_user_id() {
        let messages = vec![message("1", "peer", Some("Personal Contact"))];
        assert_eq!(peer_display_name_from_messages(&messages, None, None), None);
    }

    #[test]
    fn personal_chat_name_ignores_unrelated_aad_sender() {
        let messages = vec![
            message(
                "1",
                "c4e05f66-261e-42a3-b590-df5eb4fb54fd",
                Some("Wrong Entra Identity"),
            ),
            message("2", "47885ff3e9ffcead", Some("Jira Notifier")),
        ];

        assert_eq!(
            peer_display_name_from_messages(&messages, Some("me"), Some("47885ff3e9ffcead"),),
            Some("Jira Notifier")
        );
    }

    #[test]
    fn personal_chat_name_does_not_fall_back_to_aad_sender() {
        let messages = vec![message(
            "1",
            "c4e05f66-261e-42a3-b590-df5eb4fb54fd",
            Some("Wrong Entra Identity"),
        )];

        assert_eq!(
            peer_display_name_from_messages(&messages, Some("me"), Some("47885ff3e9ffcead"),),
            None
        );
    }
}

#[cfg(test)]
mod ntfy_category_tag_tests {
    use super::ntfy_tag_for_category;

    #[test]
    fn maps_notification_categories_to_ntfy_emoji_tags() {
        assert_eq!(ntfy_tag_for_category("Mail"), Some("email"));
        assert_eq!(ntfy_tag_for_category("Teams"), Some("speech_balloon"));
        assert_eq!(ntfy_tag_for_category("Calendar"), Some("calendar"));
        assert_eq!(ntfy_tag_for_category("Unknown"), None);
    }
}

#[cfg(test)]
mod ntfy_snooze_poll_tests {
    use super::{
        next_poll_retry_interval, ntfy_snooze_remaining_at, POLL_BACKOFF_MAX_SECONDS, POLL_SECONDS,
    };

    #[test]
    fn poll_retry_backoff_caps_at_five_minutes() {
        let mut interval = POLL_SECONDS;
        let mut actual = Vec::new();
        for _ in 0..5 {
            interval = next_poll_retry_interval(interval);
            actual.push(interval);
        }

        assert_eq!(actual, vec![40, 80, 160, 300, 300]);
        assert_eq!(actual.last().copied(), Some(POLL_BACKOFF_MAX_SECONDS));
    }

    #[test]
    fn persisted_snooze_remaining_expires_at_deadline() {
        let now = std::time::SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(1_000);
        let until = now + std::time::Duration::from_secs(90);

        assert_eq!(
            ntfy_snooze_remaining_at(Some(until), now),
            Some(std::time::Duration::from_secs(90))
        );
        assert_eq!(ntfy_snooze_remaining_at(Some(until), until), None);
        assert_eq!(ntfy_snooze_remaining_at(None, now), None);
    }

    #[test]
    fn snooze_state_uses_token_cache_directory_and_round_trips_deadline() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::SystemTime::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "m365-tui-ntfy-snooze-{}-{nonce}",
            std::process::id()
        ));
        std::fs::create_dir_all(&dir).unwrap();

        let token_cache = dir.join("token-cache.json");
        let state_path = super::ntfy_snooze_state_path(&token_cache);
        assert_eq!(state_path, dir.join("ntfy-snooze"));

        let until = std::time::SystemTime::now() + std::time::Duration::from_secs(3600);
        super::store_ntfy_snooze_until(&token_cache, until).unwrap();
        assert!(state_path.exists());

        let loaded =
            super::load_ntfy_snooze_file(&state_path, std::time::SystemTime::now()).unwrap();
        assert_eq!(
            loaded
                .duration_since(std::time::SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_secs(),
            until
                .duration_since(std::time::SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_secs()
        );

        std::fs::remove_file(&state_path).unwrap();
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn expired_or_malformed_snooze_state_is_removed() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::SystemTime::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "m365-tui-ntfy-snooze-invalid-{}-{nonce}",
            std::process::id()
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("ntfy-snooze");

        std::fs::write(&path, b"not-an-epoch\n").unwrap();
        assert!(super::load_ntfy_snooze_file(&path, std::time::SystemTime::now()).is_none());
        assert!(!path.exists());

        std::fs::write(&path, b"1\n").unwrap();
        assert!(super::load_ntfy_snooze_file(&path, std::time::SystemTime::now()).is_none());
        assert!(!path.exists());

        std::fs::remove_dir_all(&dir).unwrap();
    }
}

#[cfg(test)]
mod hot_poll_tests {
    use super::{
        chat_poll_tier_from_age, teams_message_id_is_newer, teams_message_is_incoming, App,
        ChatPollTier,
    };
    use m365_core::models::ChatMessage;

    #[test]
    fn message_id_tiebreak_is_monotonic() {
        assert!(teams_message_id_is_newer("101", "100"));
        assert!(!teams_message_id_is_newer("100", "101"));
        assert!(teams_message_id_is_newer("b", "a"));
        assert!(!teams_message_id_is_newer("a", "b"));
    }

    fn message(id: &str, created: &str) -> ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": id,
            "createdDateTime": created,
            "body": {
                "contentType": "text",
                "content": id
            }
        }))
        .expect("valid chat-message fixture")
    }

    #[test]
    fn application_messages_are_incoming_but_own_and_system_messages_are_not() {
        let application: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "app-1",
            "createdDateTime": "2026-10-07T07:00:00Z",
            "from": {
                "application": {
                    "id": "workflow-app",
                    "displayName": "Workflows"
                }
            },
            "body": {
                "contentType": "text",
                "content": "Automated message"
            }
        }))
        .unwrap();
        assert!(teams_message_is_incoming(&application, Some("me")));

        let own: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "own-1",
            "from": {
                "user": {
                    "id": "me",
                    "displayName": "Me"
                }
            }
        }))
        .unwrap();
        assert!(!teams_message_is_incoming(&own, Some("me")));

        let system: ChatMessage = serde_json::from_value(serde_json::json!({
            "id": "system-1",
            "messageType": "systemEventMessage",
            "from": {
                "application": {
                    "id": "workflow-app",
                    "displayName": "Workflows"
                }
            }
        }))
        .unwrap();
        assert!(!teams_message_is_incoming(&system, Some("me")));
    }

    #[test]
    fn hot_poll_notifications_include_every_message_after_baseline() {
        let messages = vec![
            message("300", "2026-09-25T08:00:03Z"),
            message("100", "2026-09-25T08:00:01Z"),
            message("200", "2026-09-25T08:00:02Z"),
        ];

        let ids: Vec<&str> = App::hot_poll_new_messages(&messages, Some("100"), true)
            .into_iter()
            .map(|message| message.id.as_str())
            .collect();

        assert_eq!(ids, vec!["200", "300"]);
    }

    #[test]
    fn hot_poll_notification_missing_baseline_announces_only_newest() {
        let messages = vec![
            message("200", "2026-09-25T08:00:02Z"),
            message("300", "2026-09-25T08:00:03Z"),
            message("100", "2026-09-25T08:00:01Z"),
        ];

        let ids: Vec<&str> = App::hot_poll_new_messages(&messages, Some("missing"), true)
            .into_iter()
            .map(|message| message.id.as_str())
            .collect();

        assert_eq!(ids, vec!["300"]);
    }

    #[test]
    fn hot_poll_notification_does_nothing_without_newest_change() {
        let messages = vec![
            message("100", "2026-09-25T08:00:01Z"),
            message("200", "2026-09-25T08:00:02Z"),
        ];

        assert!(App::hot_poll_new_messages(&messages, Some("100"), false).is_empty());
    }

    #[test]
    fn activity_age_maps_to_expected_poll_tier() {
        use std::time::Duration;

        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(0)),
            ChatPollTier::Hot
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(120)),
            ChatPollTier::Hot
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(121)),
            ChatPollTier::Warm
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(300)),
            ChatPollTier::Warm
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(301)),
            ChatPollTier::Cool
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(900)),
            ChatPollTier::Cool
        );
        assert_eq!(
            chat_poll_tier_from_age(Duration::from_secs(901)),
            ChatPollTier::Normal
        );
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn upstream_and_local_global_shortcuts_coexist() {
        let mut app = super::test_app();
        std::sync::Arc::make_mut(&mut app.session.config).ntfy = m365_core::config::NtfyMode::Away;
        app.on_key(crossterm::event::KeyEvent::new(
            crossterm::event::KeyCode::F(4),
            crossterm::event::KeyModifiers::NONE,
        ));
        assert!(matches!(
            app.overlay,
            Some(super::Overlay::NtfySnooze { .. })
        ));
        app.on_key(crossterm::event::KeyEvent::new(
            crossterm::event::KeyCode::Esc,
            crossterm::event::KeyModifiers::NONE,
        ));
        app.on_key(crossterm::event::KeyEvent::new(
            crossterm::event::KeyCode::Char('s'),
            crossterm::event::KeyModifiers::NONE,
        ));
        assert!(matches!(app.overlay, Some(super::Overlay::Settings { .. })));
    }

    use super::*;
    use super::{
        calendar_reminder_threshold_minutes, merge_newest_first, next_field, ntfy_presence_is_away,
        parse_recipients, reconcile_authoritative_message_window, step,
    };
    use m365_core::models::ChatMessage;

    #[test]
    fn ntfy_away_mode_is_strictly_graph_away() {
        let presence = |availability: &str| -> m365_core::models::Presence {
            serde_json::from_value(serde_json::json!({
                "availability": availability,
                "activity": availability
            }))
            .unwrap()
        };

        let away = presence("Away");
        let available = presence("Available");
        let offline = presence("Offline");

        assert!(ntfy_presence_is_away(Some(&away)));
        assert!(!ntfy_presence_is_away(Some(&available)));
        assert!(!ntfy_presence_is_away(Some(&offline)));
        assert!(!ntfy_presence_is_away(None));
    }

    #[test]
    fn ntfy_workday_modes_require_graph_work_plan() {
        let presence = |availability: &str| -> m365_core::models::Presence {
            serde_json::from_value(serde_json::json!({
                "availability": availability,
                "activity": availability
            }))
            .unwrap()
        };

        let away = presence("Away");
        let available = presence("Available");

        assert!(super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::Always,
            Some(&available),
            None,
        ));
        assert!(super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::Away,
            Some(&away),
            Some(false),
        ));

        assert!(!super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AlwaysWd,
            Some(&available),
            None,
        ));
        assert!(!super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AlwaysWd,
            Some(&available),
            Some(false),
        ));
        assert!(super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AlwaysWd,
            Some(&available),
            Some(true),
        ));

        assert!(!super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AwayWd,
            Some(&available),
            Some(true),
        ));
        assert!(!super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AwayWd,
            Some(&away),
            Some(false),
        ));
        assert!(super::ntfy_mode_should_forward(
            m365_core::config::NtfyMode::AwayWd,
            Some(&away),
            Some(true),
        ));
    }

    #[test]
    fn calendar_reminder_thresholds_are_exact() {
        assert_eq!(calendar_reminder_threshold_minutes(901), None);
        assert_eq!(calendar_reminder_threshold_minutes(900), Some(15));
        assert_eq!(calendar_reminder_threshold_minutes(301), Some(15));
        assert_eq!(calendar_reminder_threshold_minutes(300), Some(5));
        assert_eq!(calendar_reminder_threshold_minutes(1), Some(5));
        assert_eq!(calendar_reminder_threshold_minutes(0), None);
        assert_eq!(calendar_reminder_threshold_minutes(-1), None);
    }

    #[test]
    fn messages_sort_chronologically_regardless_of_arrival_order() {
        use super::sort_key;
        let msg = |id: &str, at: &str| -> m365_core::models::ChatMessage {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "createdDateTime": at,
                "body": { "contentType": "text", "content": "x" }
            }))
            .unwrap()
        };
        // Same minute, different seconds — the case Graph returned out of order.
        let mut list = [
            msg("1785859200000", "2026-08-04T15:59:50Z"),
            msg("1785859141228", "2026-08-04T15:59:01Z"),
            msg("1785859178276", "2026-08-04T15:59:38Z"),
        ];
        list.sort_by_key(sort_key);
        let ids: Vec<&str> = list.iter().map(|m| m.id.as_str()).collect();
        assert_eq!(
            ids,
            ["1785859141228", "1785859178276", "1785859200000"],
            "oldest first"
        );

        // A missing timestamp must not panic or reorder everything else.
        let mut with_gap = [
            msg("2", "2026-08-04T16:00:00Z"),
            serde_json::from_value(serde_json::json!({ "id": "1" })).unwrap(),
        ];
        with_gap.sort_by_key(sort_key);
        assert_eq!(with_gap[0].id, "1");
    }

    #[test]
    fn merge_keeps_older_pages_and_dedups() {
        // Newest-first: the user had scrolled back to "a"; a refresh returns the
        // newest three, one of which ("e") is new.
        let existing = vec!["d", "c", "b", "a"];
        let fresh = vec!["e", "d", "c"];
        let merged = merge_newest_first(fresh, existing, |s: &&str| s.to_string());
        assert_eq!(merged, vec!["e", "d", "c", "b", "a"]);
    }

    fn cached_chat_message(id: &str, created: &str, body: &str, deleted: bool) -> ChatMessage {
        serde_json::from_value(serde_json::json!({
            "id": id,
            "createdDateTime": created,
            "body": {
                "contentType": "text",
                "content": body
            },
            "deletedDateTime": deleted.then_some("2026-10-15T12:00:00Z")
        }))
        .expect("valid cached chat-message fixture")
    }

    #[test]
    fn authoritative_message_window_drops_stale_entries_but_keeps_older_history() {
        let fresh = vec![
            cached_chat_message("3000", "2026-10-15T10:03:00Z", "newest", false),
            cached_chat_message("2000", "2026-10-15T10:02:00Z", "fresh replacement", false),
        ];
        let existing = vec![
            cached_chat_message("1000", "2026-10-15T10:01:00Z", "older page", false),
            cached_chat_message("2000", "2026-10-15T10:02:00Z", "old body", false),
            cached_chat_message("2500", "2026-10-15T10:02:30Z", "stale ghost", false),
        ];

        let merged = reconcile_authoritative_message_window(fresh, existing);
        let ids: Vec<&str> = merged.iter().map(|message| message.id.as_str()).collect();

        assert_eq!(ids, vec!["1000", "2000", "3000"]);
        assert_eq!(merged[1].text(), "fresh replacement");
    }

    #[test]
    fn authoritative_message_window_uses_delete_tombstone_to_purge_cache() {
        let fresh = vec![
            cached_chat_message("3000", "2026-10-15T10:03:00Z", "deleted", true),
            cached_chat_message("2000", "2026-10-15T10:02:00Z", "still here", false),
        ];
        let existing = vec![
            cached_chat_message("1000", "2026-10-15T10:01:00Z", "older page", false),
            cached_chat_message("3000", "2026-10-15T10:03:00Z", "old cached copy", false),
        ];

        let merged = reconcile_authoritative_message_window(fresh, existing);
        let ids: Vec<&str> = merged.iter().map(|message| message.id.as_str()).collect();

        assert_eq!(ids, vec!["1000", "2000"]);
    }

    #[test]
    fn empty_authoritative_message_window_clears_cached_conversation() {
        let existing = vec![cached_chat_message(
            "1000",
            "2026-10-15T10:01:00Z",
            "stale",
            false,
        )];

        assert!(reconcile_authoritative_message_window(Vec::new(), existing).is_empty());
    }

    #[test]
    fn merge_handles_empty_sides() {
        let none: Vec<&str> = vec![];
        assert_eq!(
            merge_newest_first(vec!["a"], none.clone(), |s: &&str| s.to_string()),
            vec!["a"]
        );
        assert_eq!(
            merge_newest_first(none, vec!["a", "b"], |s: &&str| s.to_string()),
            vec!["a", "b"]
        );
    }

    #[test]
    fn merge_is_idempotent_when_nothing_changed() {
        let existing = vec!["c", "b", "a"];
        let merged = merge_newest_first(vec!["c", "b", "a"], existing, |s: &&str| s.to_string());
        assert_eq!(
            merged,
            vec!["c", "b", "a"],
            "no duplicates on an unchanged refresh"
        );
    }

    #[test]
    fn reads_resident_memory_from_proc_status() {
        use super::parse_vmrss;
        let sample = "Name:\tm365\nVmPeak:\t  123456 kB\nVmRSS:\t   24680 kB\nThreads:\t9\n";
        assert_eq!(parse_vmrss(sample), Some(24680));
        assert_eq!(parse_vmrss("no such field"), None);
    }

    #[test]
    fn teams_archive_drawer_rows_map_active_and_hidden_chats() {
        let chat = |id: &str, hidden: bool| -> m365_core::models::Chat {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "chatType": "group",
                "viewpoint": { "isHidden": hidden }
            }))
            .unwrap()
        };

        let chats = vec![
            chat("active-a", false),
            chat("hidden-a", true),
            chat("active-b", false),
            chat("hidden-b", true),
        ];
        let stay_archived = std::collections::HashSet::new();

        assert_eq!(super::teams_active_chat_count(&chats, &stay_archived), 2);
        assert_eq!(super::teams_archived_chat_count(&chats, &stay_archived), 2);
        assert_eq!(
            super::teams_chat_row_count(&chats, &stay_archived, false),
            3
        );
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, false, 0),
            Some(0)
        );
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, false, 1),
            Some(2)
        );
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, false, 2),
            None,
            "drawer row is not a chat"
        );

        assert_eq!(super::teams_chat_row_count(&chats, &stay_archived, true), 5);
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, true, 2),
            None,
            "drawer remains between active and archived rows"
        );
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, true, 3),
            Some(1)
        );
        assert_eq!(
            super::teams_chat_index_for_row(&chats, &stay_archived, true, 4),
            Some(3)
        );

        assert_eq!(
            super::teams_chat_row_for_id(&chats, &stay_archived, true, "hidden-a"),
            Some(3)
        );
        assert_eq!(
            super::teams_chat_row_for_id(&chats, &stay_archived, false, "hidden-a"),
            None,
            "hidden chat has no visible row while drawer is closed"
        );
    }

    #[test]
    fn stay_archived_keeps_server_visible_chat_in_archive_rows() {
        let chat = |id: &str, hidden: bool| -> m365_core::models::Chat {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "chatType": "group",
                "viewpoint": { "isHidden": hidden }
            }))
            .unwrap()
        };

        let chats = vec![
            chat("active", false),
            chat("stay", false),
            chat("hidden", true),
        ];
        let stay_archived = std::collections::HashSet::from(["stay".to_string()]);

        assert_eq!(super::teams_active_chat_count(&chats, &stay_archived), 1);
        assert_eq!(super::teams_archived_chat_count(&chats, &stay_archived), 2);
        assert_eq!(
            super::teams_chat_row_for_id(&chats, &stay_archived, true, "stay"),
            Some(2)
        );

        let in_flight = std::collections::HashSet::new();
        assert_eq!(
            super::teams_stay_archived_rehide_candidates(&chats, &stay_archived, &in_flight),
            vec!["stay".to_string()]
        );

        let in_flight = std::collections::HashSet::from(["stay".to_string()]);
        assert!(
            super::teams_stay_archived_rehide_candidates(&chats, &stay_archived, &in_flight)
                .is_empty()
        );
    }

    #[test]
    fn stay_archived_state_round_trips() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "m365-tui-stay-archived-{}-{nonce}",
            std::process::id()
        ));

        let expected = std::collections::HashSet::from([
            "19:first@thread.v2".to_string(),
            "19:second@thread.v2".to_string(),
        ]);

        super::store_stay_archived(&dir, &expected).unwrap();
        assert_eq!(super::load_stay_archived(Some(&dir)), expected);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn ui_state_round_trips_archive_drawer_state() {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!(
            "m365-tui-ui-state-archive-{}-{nonce}",
            std::process::id()
        ));

        super::store_ui_state(
            &dir,
            super::Screen::Teams,
            super::DEFAULT_CALENDAR_DAYS,
            super::CalendarView::Agenda,
            Some("19:test@thread.v2"),
            true,
        )
        .unwrap();

        let state = super::load_ui_state(Some(&dir));
        assert!(state.teams_archive_expanded);

        super::store_ui_state(
            &dir,
            super::Screen::Teams,
            super::DEFAULT_CALENDAR_DAYS,
            super::CalendarView::Agenda,
            Some("19:test@thread.v2"),
            false,
        )
        .unwrap();

        let state = super::load_ui_state(Some(&dir));
        assert!(!state.teams_archive_expanded);

        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn selection_step_clamps() {
        assert_eq!(step(0, -1, 5), 0);
        assert_eq!(step(4, 1, 5), 4);
        assert_eq!(step(2, 1, 5), 3);
        assert_eq!(step(0, 1, 0), 0, "empty list stays at 0");
    }

    #[test]
    fn folder_panel_fits_rendered_labels() {
        let folder = |name: &str, unread: i64| {
            serde_json::from_value(serde_json::json!({
                "id": name,
                "displayName": name,
                "unreadItemCount": unread
            }))
            .unwrap()
        };
        assert_eq!(folder_panel_width(&[]), 11, "title remains visible");
        assert_eq!(folder_panel_width(&[folder("Inbox", 0)]), 11);
        assert_eq!(folder_panel_width(&[folder("Long folder", 12)]), 19);
    }

    #[test]
    fn folder_panel_height_tracks_items_with_bounds() {
        let folder = |name: &str| -> m365_core::models::MailFolder {
            serde_json::from_value(serde_json::json!({ "id": name, "displayName": name })).unwrap()
        };
        assert_eq!(folder_panel_height(&[]), 5);
        assert_eq!(folder_panel_height(&[folder("Inbox")]), 5);
        assert_eq!(folder_panel_height(&vec![folder("x"); 20]), 14);
    }

    #[test]
    fn panel_resize_respects_minimum() {
        assert_eq!(resize_panel_extent(8, -1, 5), 7);
        assert_eq!(resize_panel_extent(5, -1, 5), 5);
        assert_eq!(resize_panel_extent(8, 1, 5), 9);
    }

    #[test]
    fn compose_field_cycles_within_allowed_fields() {
        assert_eq!(next_field(&[0, 1, 2], 0), 1);
        assert_eq!(next_field(&[0, 1, 2], 2), 0);
        assert_eq!(next_field(&[2], 2), 2, "reply has only a body field");
        assert_eq!(next_field(&[0, 2], 0), 2, "forward skips subject");
        assert_eq!(
            ComposeKind::NewMail.fields(),
            &[
                COMPOSE_TO,
                COMPOSE_CC,
                COMPOSE_BCC,
                COMPOSE_SUBJECT,
                COMPOSE_BODY,
                COMPOSE_ATTACH
            ]
        );
        let mut reply = Compose::new(ComposeKind::ReplyMail { id: "id".into() }, COMPOSE_BODY);
        assert_eq!(reply.fields(), &[COMPOSE_BODY, COMPOSE_ATTACH]);
        reply.edit_recipients = true;
        assert_eq!(
            reply.fields(),
            &[
                COMPOSE_TO,
                COMPOSE_CC,
                COMPOSE_BCC,
                COMPOSE_BODY,
                COMPOSE_ATTACH
            ]
        );
    }

    #[test]
    fn recipients_split_on_common_separators() {
        assert_eq!(
            parse_recipients("a@x.pt, b@x.pt; c@x.pt d@x.pt"),
            vec!["a@x.pt", "b@x.pt", "c@x.pt", "d@x.pt"]
        );
        assert!(parse_recipients("  ,; ").is_empty());
    }

    #[test]
    fn mail_content_ids_are_parsed_and_normalized() {
        assert_eq!(mail_content_id("cid:logo@example"), Some("logo@example"));
        assert_eq!(
            mail_content_id("CID:<logo@example>"),
            Some("<logo@example>")
        );
        assert_eq!(normalize_content_id(" <logo@example> "), "logo@example");
        assert_eq!(mail_content_id("https://example.com/logo.png"), None);
        assert!(!mail_image_is_local("https://example.com/pixel.gif"));
        assert!(mail_image_is_local("cid:shot.png"));
        assert!(mail_image_is_local(
            "https://graph.microsoft.com/v1.0/me/messages/m/attachments/a/$value"
        ));
        assert_eq!(
            graph_attachment_path(
                "https://graph.microsoft.com/v1.0/me/messages/m/attachments/a/$value?foo=1"
            )
            .as_deref(),
            Some("me/messages/m/attachments/a/$value")
        );
        let entries = reading_mail_image_entries(&[
            MailBodyImage {
                message_id: "m".into(),
                image: crate::content::BodyImage {
                    alt: "shot".into(),
                    src: "cid:shot.png".into(),
                    at_line: 0,
                },
            },
            MailBodyImage {
                message_id: "m".into(),
                image: crate::content::BodyImage {
                    alt: "pixel".into(),
                    src: "https://example.com/pixel.gif".into(),
                    at_line: 1,
                },
            },
        ]);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].1, "shot");
        assert!(entries[0].0.contains("cid:shot.png"));
        let att: m365_core::models::Attachment = serde_json::from_value(serde_json::json!({
            "id": "1",
            "name": "image001.png",
            "contentId": "image001.png@01DABC"
        }))
        .unwrap();
        assert!(cid_matches(&att, "image001.png@01DABC"));
        assert!(cid_matches(&att, "image001.png"));
        assert!(cid_matches(&att, "<image001.png@01DABC>"));
    }

    #[test]
    fn recipient_autocomplete_replaces_only_the_active_recipient() {
        let mut compose = Compose::new(ComposeKind::NewMail, COMPOSE_TO);
        compose.to = crate::editor::TextInput::from("first@example.com; dav");
        assert_eq!(recipient_token(&compose.to), (19, 22, "dav".into()));
        accept_recipient_suggestion(
            &mut compose,
            &MailContact {
                name: Some("David Example".into()),
                address: "david@example.com".into(),
            },
        );
        assert_eq!(compose.to.text(), "first@example.com; david@example.com, ");
    }

    #[test]
    fn contact_cache_deduplicates_addresses_and_learns_names() {
        let recipient = |name: Option<&str>| {
            serde_json::from_value(serde_json::json!({
                "emailAddress": { "name": name, "address": "Person@Example.com" }
            }))
            .unwrap()
        };
        let mut contacts = std::collections::BTreeMap::new();
        cache_mail_contact(&mut contacts, &recipient(None));
        cache_mail_contact(&mut contacts, &recipient(Some("A Person")));
        assert_eq!(contacts.len(), 1);
        assert_eq!(
            contacts["person@example.com"].name.as_deref(),
            Some("A Person")
        );
    }

    #[test]
    fn compose_file_round_trips_subject_and_body() {
        let file = format_compose_file("Hello", "line1\nline2");
        assert_eq!(file, "Subject: Hello\n\nline1\nline2");
        assert_eq!(
            parse_compose_file(&file),
            (Some("Hello".into()), "line1\nline2".into())
        );
        assert_eq!(
            parse_compose_file("just the body\n"),
            (None, "just the body\n".into())
        );
    }

    #[test]
    fn folder_search_matches_labels_case_insensitively() {
        let folder = |name: &str| {
            serde_json::from_value(serde_json::json!({
                "id": name,
                "displayName": name
            }))
            .unwrap()
        };
        let folders = [folder("Inbox"), folder("Sent Items"), folder("Archive")];
        assert_eq!(filter_folders(&folders, ""), vec![0, 1, 2]);
        assert_eq!(filter_folders(&folders, "in"), vec![0]);
        assert_eq!(filter_folders(&folders, "item"), vec![1]);
        assert_eq!(filter_folders(&folders, "ARCHIVE"), vec![2]);
        assert!(filter_folders(&folders, "zzz").is_empty());
    }

    #[test]
    fn mail_threads_collapse_to_the_newest_loaded_message() {
        let message = |id: &str, conversation: Option<&str>| {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "conversationId": conversation,
                "subject": id
            }))
            .unwrap()
        };
        // Folder pages are newest-first. Missing conversation ids deliberately
        // remain distinct instead of collapsing into one synthetic thread.
        let messages = [
            message("new-a", Some("a")),
            message("only-b", Some("b")),
            message("old-a", Some("a")),
            message("standalone-1", None),
            message("standalone-2", None),
        ];
        assert_eq!(mail_row_indices(&messages, false), vec![0, 1, 2, 3, 4]);
        assert_eq!(mail_row_indices(&messages, true), vec![0, 1, 3, 4]);
        assert_eq!(mail_thread_size(&messages, &messages[0]), 2);
        assert_eq!(mail_thread_size(&messages, &messages[3]), 1);
    }

    #[test]
    fn single_message_unread_state_ignores_other_rows() {
        let message = |id: &str, read: bool| {
            serde_json::from_value(serde_json::json!({
                "id": id,
                "conversationId": "thread",
                "isRead": read
            }))
            .unwrap()
        };
        let messages = [message("read", true), message("unread", false)];
        assert!(!mail_row_unread(&messages, &messages[0], false));
        assert!(mail_row_unread(&messages, &messages[0], true));
    }
}
