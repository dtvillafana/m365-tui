//! m365 — a unified terminal client for Outlook and Microsoft Teams.
//!
//! Usage:
//!   m365            launch the TUI
//!   m365 whoami     print the signed-in user and exit (auth smoke test)
//!   m365 login      run device-code login and exit
//!   m365 forward    forward a message and exit
//!   m365 --help     usage; also --version
//!
//! Arguments are resolved before any configuration is read or sign-in is
//! attempted, so `--help` and `--version` work on a machine that has never been
//! configured. Anything else would greet a first-time user with a device-code
//! prompt for asking what the flags are.

mod app;
mod clipboard;
mod content;
mod editor;
mod files;
mod images;
mod navigation;
mod notify;
mod opener;
mod termimg;
mod ui;
mod wrap;

use std::io::stdout;
use std::time::Duration;

use anyhow::{Context, Result};
use app::{format_compose_file, parse_compose_file, App, AppMessage, PushState};
use crossterm::event::{
    DisableBracketedPaste, EnableBracketedPaste, Event, EventStream, KeyEventKind,
};
use crossterm::execute;
use crossterm::terminal::{
    disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen,
};
use futures_util::StreamExt;
use m365_core::events::ChangeEvent;
use m365_core::{subscriptions, DeviceCodePrompt, Session};
use ratatui::backend::CrosstermBackend;
use ratatui::Terminal;
use tokio::sync::mpsc;

/// How often to poll the server to refresh the current view.
const POLL_SECONDS: u64 = 20;

/// Lightweight local tick: refreshes memory usage and ages out status text.
/// Costs nothing on the network.
const TICK_SECONDS: u64 = 2;

/// What the command line asked for, decided before any I/O.
enum Command {
    Tui,
    WhoAmI,
    Login,
    Forward {
        id: Option<String>,
        query: Option<String>,
        to: Vec<String>,
        comment: String,
    },
}

const USAGE: &str = "\
m365 — a terminal client for Outlook and Microsoft Teams

USAGE:
    m365 [COMMAND]

COMMANDS:
    (none)      launch the TUI
    login       sign in and cache the token, then exit
    whoami      print the signed-in account, then exit
    forward     forward a message: m365 forward --to ADDR[,ADDR...] [--comment TEXT] (--id ID | --query TEXT | ID)

OPTIONS:
    -h, --help     print this help
    -V, --version  print the version

Configuration is read from the environment or a .env file; M365_CLIENT_ID is
the only required value. See https://github.com/rootHytx/m365-tui for setup.";

const FORWARD_USAGE: &str = "\
m365 forward — send an existing Outlook message to new recipients

USAGE:
    m365 forward --to ADDR[,ADDR...] [--comment TEXT] --id ID
    m365 forward --to ADDR[,ADDR...] [--comment TEXT] --query TEXT
    m365 forward --to ADDR[,ADDR...] [--comment TEXT] ID

--query uses Graph mailbox search. A single match is forwarded; several
matches are listed so you can pass --id.";

fn parse_args() -> Command {
    parse_args_from(std::env::args().skip(1))
}

fn parse_args_from<I, S>(args: I) -> Command
where
    I: IntoIterator<Item = S>,
    S: AsRef<str>,
{
    let mut args = args.into_iter();
    match args.next().as_ref().map(|s| s.as_ref()) {
        None => Command::Tui,
        Some("whoami") => Command::WhoAmI,
        Some("login") => Command::Login,
        Some("forward") => match parse_forward_args(args) {
            Ok(cmd) => cmd,
            Err(e) => {
                eprintln!("m365 forward: {e}\n");
                eprintln!("{FORWARD_USAGE}");
                std::process::exit(2);
            }
        },
        Some("-h") | Some("--help") | Some("help") => {
            println!("{USAGE}");
            std::process::exit(0);
        }
        Some("-V") | Some("--version") | Some("version") => {
            println!("m365 {}", env!("CARGO_PKG_VERSION"));
            std::process::exit(0);
        }
        Some(other) => {
            eprintln!("m365: unrecognised argument '{other}'\n");
            eprintln!("{USAGE}");
            std::process::exit(2);
        }
    }
}

fn parse_forward_args<I, S>(args: I) -> Result<Command, String>
where
    I: IntoIterator<Item = S>,
    S: AsRef<str>,
{
    let args: Vec<String> = args.into_iter().map(|s| s.as_ref().to_string()).collect();
    let mut to = None;
    let mut comment = String::new();
    let mut id = None;
    let mut query = None;
    let mut i = 0;
    while i < args.len() {
        match args[i].as_str() {
            "-h" | "--help" => {
                println!("{FORWARD_USAGE}");
                std::process::exit(0);
            }
            "--to" | "-t" => {
                i += 1;
                let value = args
                    .get(i)
                    .ok_or_else(|| "missing value for --to".to_string())?;
                to = Some(split_addrs(value));
            }
            "--comment" | "-c" => {
                i += 1;
                comment = args
                    .get(i)
                    .ok_or_else(|| "missing value for --comment".to_string())?
                    .clone();
            }
            "--id" => {
                i += 1;
                id = Some(
                    args.get(i)
                        .ok_or_else(|| "missing value for --id".to_string())?
                        .clone(),
                );
            }
            "--query" | "-q" => {
                i += 1;
                query = Some(
                    args.get(i)
                        .ok_or_else(|| "missing value for --query".to_string())?
                        .clone(),
                );
            }
            flag if flag.starts_with('-') => {
                return Err(format!("unrecognised flag '{flag}'"));
            }
            positional => {
                if id.is_some() {
                    return Err("only one message id is allowed".into());
                }
                id = Some(positional.to_string());
            }
        }
        i += 1;
    }

    let to = to.ok_or_else(|| "missing --to".to_string())?;
    if to.is_empty() {
        return Err("add at least one recipient with --to".into());
    }
    if id.is_none() && query.as_ref().is_none_or(|q| q.trim().is_empty()) {
        return Err("pass --id ID or --query TEXT".into());
    }
    Ok(Command::Forward {
        id,
        query,
        to,
        comment,
    })
}

fn split_addrs(s: &str) -> Vec<String> {
    s.split([',', ';'])
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect()
}

#[tokio::main]
async fn main() -> Result<()> {
    // Before anything else: --help and --version must not require configuration,
    // a network, or a signed-in account.
    let command = parse_args();

    init_tracing();

    let session = match Session::from_env() {
        Ok(s) => s,
        Err(e) => {
            eprintln!("configuration error: {e:#}");
            eprintln!("\nSet at least M365_CLIENT_ID (see README / .env.example).");
            std::process::exit(1);
        }
    };

    // Device-code login up front (prints the code to stdout, before the TUI).
    session
        .ensure_logged_in(|p: DeviceCodePrompt| print_device_prompt(&p))
        .await
        .context("sign-in failed")?;

    match command {
        Command::WhoAmI => {
            let me = session.whoami().await?;
            println!(
                "{} <{}>",
                me.display_name.clone().unwrap_or_default(),
                me.best_email().unwrap_or("")
            );
            Ok(())
        }
        Command::Login => {
            println!("signed in — token cached.");
            Ok(())
        }
        Command::Forward {
            id,
            query,
            to,
            comment,
        } => run_forward(&session, id, query, to, comment).await,
        Command::Tui => run_tui(session).await,
    }
}

async fn run_forward(
    session: &Session,
    id: Option<String>,
    query: Option<String>,
    to: Vec<String>,
    comment: String,
) -> Result<()> {
    let id = match id {
        Some(id) => id,
        None => {
            let query = query.expect("parse_forward_args requires --query when --id is absent");
            resolve_forward_id(&session.graph, &query).await?
        }
    };
    m365_core::mail::forward(&session.graph, &id, &to, &comment).await?;
    println!("forwarded to {}", to.join(", "));
    Ok(())
}

async fn resolve_forward_id(graph: &m365_core::GraphClient, query: &str) -> Result<String> {
    let matches = m365_core::mail::search(graph, query, 10).await?;
    match matches.len() {
        0 => anyhow::bail!("no messages matched {query:?}"),
        1 => Ok(matches[0].id.clone()),
        _ => {
            eprintln!("multiple matches; pass --id to choose:");
            for m in &matches {
                eprintln!(
                    "  {}  {}  {}",
                    m.id,
                    m.received_date_time.as_deref().unwrap_or("-"),
                    m.subject.as_deref().unwrap_or("(no subject)")
                );
            }
            anyhow::bail!("{} messages matched {query:?}", matches.len());
        }
    }
}

fn print_device_prompt(p: &DeviceCodePrompt) {
    println!("\n──────────────────────────────────────────────");
    println!(" Sign in to Microsoft 365");
    println!(" 1. Open: {}", p.verification_uri);
    println!(" 2. Enter code: {}", p.user_code);
    println!("──────────────────────────────────────────────");
    println!(" {}", p.message);
    println!(" (waiting for you to finish in the browser…)\n");
}

async fn run_tui(session: Session) -> Result<()> {
    // Channels: background task results, and Graph change events.
    let (tx, mut app_rx) = mpsc::channel::<AppMessage>(256);
    let (change_tx, mut change_rx) = mpsc::channel::<ChangeEvent>(256);

    // Real-time push: only if a tunnel URL is configured.
    if session.config.notification_url().is_some() {
        spawn_realtime(&session, change_tx, tx.clone());
    } else {
        let _ = tx
            .send(AppMessage::Status(format!(
                "poll mode — refreshing every {POLL_SECONDS}s (set M365_TUNNEL_BASE_URL for instant push)"
            )))
            .await;
    }

    // Periodic poll so the current view refreshes from the server regardless of
    // whether push is configured.
    {
        let poll_tx = tx.clone();
        tokio::spawn(async move {
            let mut ticker = tokio::time::interval(Duration::from_secs(POLL_SECONDS));
            ticker.tick().await; // the first tick fires immediately; skip it
            loop {
                ticker.tick().await;
                if poll_tx.send(AppMessage::Poll).await.is_err() {
                    break;
                }
            }
        });
    }

    {
        let tick_tx = tx.clone();
        tokio::spawn(async move {
            let mut ticker = tokio::time::interval(Duration::from_secs(TICK_SECONDS));
            ticker.tick().await;
            loop {
                ticker.tick().await;
                if tick_tx.send(AppMessage::Tick).await.is_err() {
                    break;
                }
            }
        });
    }

    let mut app = App::new(session, tx);
    app.bootstrap();

    // Terminal setup. Query graphics *after* raw mode so DA1 replies are
    // readable, but *before* the alternate screen so the query isn't lost.
    enable_raw_mode()?;
    app.set_graphics(crate::termimg::Graphics::detect());
    let mut out = stdout();
    execute!(out, EnterAlternateScreen, EnableBracketedPaste)?;
    let mut terminal = Terminal::new(CrosstermBackend::new(out))?;

    let res = event_loop(&mut terminal, &mut app, &mut app_rx, &mut change_rx).await;

    // Terminal teardown (best-effort even on error).
    disable_raw_mode().ok();
    execute!(
        terminal.backend_mut(),
        DisableBracketedPaste,
        LeaveAlternateScreen
    )
    .ok();
    terminal.show_cursor().ok();

    // Drop our presence session on the way out, otherwise the user would keep
    // showing the status we published for up to the session lease.
    if app.presence_session.is_some() {
        let client_id = app.session.config.client_id.clone();
        if let Err(e) =
            m365_core::people::clear_session_presence(&app.session.graph, &client_id).await
        {
            tracing::warn!("could not clear presence session on exit: {e:#}");
        }
    }

    res
}

async fn event_loop(
    terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
    app: &mut App,
    app_rx: &mut mpsc::Receiver<AppMessage>,
    change_rx: &mut mpsc::Receiver<ChangeEvent>,
) -> Result<()> {
    loop {
        {
            let mut reader = EventStream::new();
            loop {
                terminal.draw(|f| ui::render(f, app))?;

                tokio::select! {
                    maybe_event = reader.next() => {
                        match maybe_event {
                            Some(Ok(Event::Key(k))) if k.kind == KeyEventKind::Press => app.on_key(k),
                            Some(Ok(Event::Paste(text))) => app.on_paste(text),
                            Some(Ok(_)) => {}
                            Some(Err(e)) => return Err(e.into()),
                            None => return Ok(()),
                        }
                    }
                    Some(msg) = app_rx.recv() => app.apply(msg),
                    Some(change) = change_rx.recv() => app.on_change(change),
                }

                if app.should_quit {
                    return Ok(());
                }
                if app.pending_external_edit {
                    break;
                }
            }
        }
        // EventStream dropped so $EDITOR can own stdin.
        app.pending_external_edit = false;
        edit_compose_in_editor(terminal, app)?;
    }
}

/// Suspend the TUI, open `$VISUAL`/`$EDITOR` on the compose subject and body,
/// then restore the alternate screen.
fn edit_compose_in_editor(
    terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
    app: &mut App,
) -> Result<()> {
    let Some((subject, body)) = app.compose_edit_snapshot() else {
        return Ok(());
    };

    disable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        DisableBracketedPaste,
        LeaveAlternateScreen
    )?;
    terminal.show_cursor()?;

    let result = run_compose_editor(&subject, &body);

    enable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        EnterAlternateScreen,
        EnableBracketedPaste
    )?;
    terminal.clear()?;
    terminal.hide_cursor()?;

    match result {
        Ok((subject, body)) => app.apply_compose_edit(subject, body),
        Err(e) => app.status = format!("editor failed: {e:#}"),
    }
    Ok(())
}

fn run_compose_editor(subject: &str, body: &str) -> Result<(Option<String>, String)> {
    let path = std::env::temp_dir().join(format!("m365-compose-{}.txt", std::process::id()));
    std::fs::write(&path, format_compose_file(subject, body))
        .with_context(|| format!("writing {}", path.display()))?;
    let editor = std::env::var("VISUAL")
        .ok()
        .filter(|s| !s.is_empty())
        .or_else(|| std::env::var("EDITOR").ok().filter(|s| !s.is_empty()))
        .unwrap_or_else(|| "vi".into());
    let status = std::process::Command::new("sh")
        .arg("-c")
        .arg(format!("{editor} \"$1\""))
        .arg("m365-editor")
        .arg(&path)
        .status()
        .with_context(|| format!("running {editor}"))?;
    let text =
        std::fs::read_to_string(&path).with_context(|| format!("reading {}", path.display()));
    let _ = std::fs::remove_file(&path);
    let text = text?;
    if !status.success() {
        anyhow::bail!("{editor} exited {status}");
    }
    Ok(parse_compose_file(&text))
}

/// Subscribe to the webhook's Redis channel and keep Graph subscriptions alive.
fn spawn_realtime(
    session: &Session,
    change_tx: mpsc::Sender<ChangeEvent>,
    app_tx: mpsc::Sender<AppMessage>,
) {
    let redis_url = session.config.redis_url.clone();
    tokio::spawn(m365_core::events::run_subscriber_forever(
        redis_url, change_tx,
    ));

    let session = session.clone();
    tokio::spawn(async move {
        if let Err(e) = manage_subscriptions(session, app_tx.clone()).await {
            tracing::warn!("subscription manager stopped: {e:#}");
            let _ = app_tx
                .send(AppMessage::Push(PushState::Failed(format!("{e:#}"))))
                .await;
        }
    });
}

/// Create the inbox + all-chats subscriptions and renew them before they lapse.
/// Chat subscriptions expire in ~1h, so we renew every 45 minutes.
async fn manage_subscriptions(session: Session, app_tx: mpsc::Sender<AppMessage>) -> Result<()> {
    let _ = app_tx.send(AppMessage::Push(PushState::Connecting)).await;
    let notify = session.config.notification_url().unwrap();
    let lifecycle = session.config.lifecycle_url();
    let state = &session.config.client_state;

    let mut ids: Vec<String> = Vec::new();
    let create = |res: String| {
        let notify = notify.clone();
        let lifecycle = lifecycle.clone();
        let graph = session.graph.clone();
        let state = state.clone();
        async move {
            subscriptions::create(
                &graph,
                &res,
                "created,updated",
                &notify,
                lifecycle.as_deref(),
                &state,
                55,
            )
            .await
        }
    };

    // The chats resource needs the signed-in user's id spelled out.
    let resources: Vec<String> = match m365_core::people::me(&session.graph).await {
        Ok(user) => vec![
            subscriptions::RES_INBOX.to_string(),
            subscriptions::res_all_chats(&user.id),
        ],
        Err(e) => {
            tracing::warn!("could not resolve the signed-in user: {e:#}");
            vec![subscriptions::RES_INBOX.to_string()]
        }
    };

    let mut last_error = None;
    for res in resources {
        match create(res.clone()).await {
            Ok(s) => {
                tracing::info!("subscribed to {res}: {}", s.id);
                ids.push(s.id);
            }
            Err(e) => {
                tracing::warn!("failed to subscribe to {res}: {e:#}");
                last_error = Some(m365_core::util::graph_error_summary(&format!("{e:#}")));
            }
        }
    }

    // Report health so a broken tunnel is visible rather than silently
    // degrading to polling.
    let _ = app_tx
        .send(AppMessage::Push(if ids.is_empty() {
            PushState::Failed(last_error.unwrap_or_else(|| "no subscriptions created".into()))
        } else {
            PushState::Live
        }))
        .await;

    loop {
        tokio::time::sleep(Duration::from_secs(45 * 60)).await;
        for id in &ids {
            if let Err(e) = subscriptions::renew(&session.graph, id, 55).await {
                tracing::warn!("failed to renew subscription {id}: {e:#}");
            }
        }
    }
}

/// Log to a file in the cache dir so we never corrupt the TUI on stdout/stderr.
fn init_tracing() {
    use tracing_subscriber::EnvFilter;
    let log_path = std::env::temp_dir().join("m365-tui.log");
    let _ = tracing_subscriber::fmt()
        .with_ansi(false)
        .with_env_filter(
            EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("warn")),
        )
        .with_writer(move || {
            std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(&log_path)
                .unwrap_or_else(|_| std::fs::File::create("/dev/null").unwrap())
        })
        .try_init();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn forward_accepts_comma_separated_recipients_and_a_positional_id() {
        let cmd = parse_forward_args(["--to", "a@x.pt, b@x.pt", "AAMk123"]).unwrap();
        let Command::Forward {
            id,
            query,
            comment,
            to,
        } = cmd
        else {
            panic!("expected Forward");
        };
        assert_eq!(id.as_deref(), Some("AAMk123"));
        assert!(query.is_none());
        assert!(comment.is_empty());
        assert_eq!(to, ["a@x.pt", "b@x.pt"]);
    }

    #[test]
    fn forward_accepts_query_and_comment() {
        let cmd = parse_forward_args([
            "--to",
            "a@x.pt",
            "--query",
            "loan survey",
            "--comment",
            "FYI",
        ])
        .unwrap();
        let Command::Forward {
            id,
            query,
            comment,
            to,
        } = cmd
        else {
            panic!("expected Forward");
        };
        assert!(id.is_none());
        assert_eq!(query.as_deref(), Some("loan survey"));
        assert_eq!(comment, "FYI");
        assert_eq!(to, ["a@x.pt"]);
    }

    #[test]
    fn forward_requires_recipients_and_a_message() {
        assert!(parse_forward_args(["--to", "a@x.pt"]).is_err());
        assert!(parse_forward_args(["AAMk123"]).is_err());
        assert!(parse_forward_args(["--to", "a@x.pt", "--query", "  "]).is_err());
    }
}
