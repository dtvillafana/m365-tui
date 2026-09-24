//! All rendering. Pure function of `&App` — no state mutation here.

use ratatui::layout::{Constraint, Direction, Layout, Rect};
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, Clear, List, ListItem, ListState, Paragraph, Wrap};
use ratatui::Frame;

use crate::app::{
    filter_commands, filter_folders, folder_panel_height, mail_folder_label, mail_image_is_local,
    mail_row_unread, mail_same_thread, mail_thread_size, App, Compose, OutlookFocus, Overlay,
    PushState, Screen, TeamsFocus, TeamsMode, COMPOSE_ATTACH, COMPOSE_BCC, COMPOSE_BODY,
    COMPOSE_CC, COMPOSE_SUBJECT, COMPOSE_TO, DEFAULT_FOLDER_PANEL_WIDTH,
};

const ACCENT: Color = Color::Cyan;
const DIM: Color = Color::DarkGray;

pub fn render(f: &mut Frame, app: &mut App) {
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Min(0),
            Constraint::Length(1),
        ])
        .split(f.area());

    // Copy mode takes over the whole frame: one hint row plus borderless,
    // full-width text so a terminal drag-select grabs only the message body.
    if app.copy_mode {
        render_copy_mode(f, app);
        return;
    }

    // A full-image overlay must be the only thing drawn this frame. Terminal
    // graphics protocols keep pixels from earlier widgets unless they are not
    // rendered at all.
    if matches!(
        app.overlay,
        Some(Overlay::ViewImage { .. }) | Some(Overlay::MailImages { .. })
    ) {
        render_overlay(f, app);
        return;
    }

    render_tabs(f, chunks[0], app);
    match app.screen {
        Screen::Outlook => render_outlook(f, chunks[1], app),
        Screen::Teams => render_teams(f, chunks[1], app),
    }
    render_status(f, chunks[2], app);

    if app.overlay.is_some() {
        render_overlay(f, app);
    }
}

/// Full-screen, borderless view of the current message/conversation. No side
/// panes and no borders, so mouse selection captures exactly the text.
fn render_copy_mode(f: &mut Frame, app: &App) {
    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(1), Constraint::Min(0)])
        .split(f.area());

    f.render_widget(
        Paragraph::new(Line::from(Span::styled(
            " COPY MODE — drag to select · y yank all · j/k scroll · g/G top/bottom · z/Esc exit ",
            Style::default()
                .fg(Color::Black)
                .bg(ACCENT)
                .add_modifier(Modifier::BOLD),
        ))),
        rows[0],
    );

    let lines = match app.screen {
        Screen::Outlook => email_lines(app).unwrap_or_default(),
        Screen::Teams => conversation_lines(app, false).0,
    };
    let (wrapped, _) = crate::wrap::wrap_all(&lines, rows[1].width as usize);
    let max = (wrapped.len() as u16).saturating_sub(rows[1].height);
    app.copy_max_scroll.set(max);
    f.render_widget(
        Paragraph::new(wrapped).scroll((app.copy_scroll.min(max), 0)),
        rows[1],
    );
}

/// Top row: which app is active on the left, live state on the right.
/// No key hints live here — those belong in the bottom bar.
fn render_tabs(f: &mut Frame, area: Rect, app: &App) {
    let tab = |name: &str, active: bool| {
        if active {
            Span::styled(
                format!(" {name} "),
                Style::default()
                    .fg(Color::Black)
                    .bg(ACCENT)
                    .add_modifier(Modifier::BOLD),
            )
        } else {
            Span::styled(format!(" {name} "), Style::default().fg(DIM))
        }
    };
    let tabs = Line::from(vec![
        tab("Outlook (F2)", app.screen == Screen::Outlook),
        Span::raw("  "),
        tab("Teams (F2)", app.screen == Screen::Teams),
    ]);

    // Right-hand state: presence · push · memory · last sync.
    let (dot, avail) = presence_indicator(app);
    let (push_label, push_colour) = match &app.push {
        PushState::Off => ("push off", DIM),
        PushState::Connecting => ("push …", Color::Yellow),
        PushState::Live => ("push live", Color::Green),
        PushState::Failed(_) => ("push FAILED", Color::Red),
    };
    let ram = match app.rss_kb {
        Some(kb) if kb >= 1024 => format!("{:.0} MB", kb as f64 / 1024.0),
        Some(kb) => format!("{kb} KB"),
        None => "—".to_string(),
    };
    let sync = match &app.last_sync {
        Some(t) => format!("⟳ {t}"),
        None => "⟳ …".to_string(),
    };
    let sep = || Span::styled(" · ", Style::default().fg(DIM));
    let state_spans = vec![
        Span::styled(format!("{dot} {avail}"), presence_style(app)),
        sep(),
        Span::styled(push_label, Style::default().fg(push_colour)),
        sep(),
        Span::styled(format!("rss {ram}"), Style::default().fg(Color::Gray)),
        sep(),
        Span::styled(sync, Style::default().fg(Color::Green)),
        Span::raw(" "),
    ];
    let state = Line::from(state_spans);

    let state_w = line_width(&state).min(area.width);
    let cols = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Min(0), Constraint::Length(state_w)])
        .split(area);
    f.render_widget(Paragraph::new(tabs), cols[0]);
    f.render_widget(Paragraph::new(state), cols[1]);
}

/// Bottom row: the latest transient message on the left, the keys available
/// right now on the right.
fn render_status(f: &mut Frame, area: Rect, app: &App) {
    let bg = Color::Rgb(30, 30, 40);
    let hints = format!(" {} · ? help ", context_hints(app));
    let hints_w = (hints.chars().count() as u16).min(area.width);

    let cols = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Min(0), Constraint::Length(hints_w)])
        .split(area);

    f.render_widget(
        Paragraph::new(Line::from(Span::styled(
            format!(" {}", app.status),
            Style::default().fg(Color::White).bg(bg),
        )))
        .style(Style::default().bg(bg)),
        cols[0],
    );
    f.render_widget(
        Paragraph::new(Line::from(Span::styled(
            hints,
            Style::default().fg(DIM).bg(bg),
        ))),
        cols[1],
    );
}

fn line_width(line: &Line) -> u16 {
    line.spans
        .iter()
        .map(|s| s.content.chars().count())
        .sum::<usize>() as u16
}

/// Key hints for whatever currently has focus.
fn context_hints(app: &App) -> &'static str {
    if let Some(overlay) = &app.overlay {
        return match overlay {
            Overlay::Compose(_) => "Ctrl+S send · Ctrl+X e $EDITOR · Esc cancel",
            Overlay::Links => "1-9 open · y copy · Esc close",
            Overlay::MailImages {
                fullscreen: Some(_),
                ..
            } => "1-9 jump · j/k next · Esc back",
            Overlay::MailImages { .. } => "1-9 fullscreen · j/k scroll · Esc close",
            Overlay::Attachments => "1-9 save · Esc close",
            Overlay::React => "1-7 react · Esc close",
            Overlay::ViewImage { keys, .. } if keys.len() > 1 => "j/k next image · Esc close",
            Overlay::ViewImage { .. } => "Esc close",
            Overlay::Presence => "1-6 set · c clear · Esc close",
            Overlay::Settings { .. } => "Space/Enter toggle · Esc close",
            Overlay::Search { .. } => "Enter search · Esc cancel",
            Overlay::FolderSearch { .. } => "↑↓ choose · Enter open · Esc cancel",
            Overlay::NewChat { .. } => "Type name/username · ↑↓ choose · Enter chat · Esc cancel",
            Overlay::Palette { .. } => "↑↓ choose · Enter run · Esc close",
            Overlay::MoveMail { .. } => "j/k/g/G choose · Enter move · Esc cancel",
            Overlay::Calendar | Overlay::Help => "Esc close",
        };
    }
    match app.screen {
        Screen::Outlook => match app.outlook_focus {
            OutlookFocus::Folders => "j/k move · g/G · l open · / find · e calendar · H/L resize",
            OutlookFocus::Messages => {
                "j/k move · g/G · l read · t threads · h back · c compose · r reply · u read · m move · d trash · / search"
            }
            OutlookFocus::Reading => {
                "j/k scroll · g/G · h back · u read · m move · d trash · o links · i images · A attach · y copy"
            }
        },
        Screen::Teams => match app.teams.focus {
            TeamsFocus::List => "j/k move · g/G · l open · n new chat · t chats/channels",
            TeamsFocus::Messages => {
                "j/k select · g/G · h back · r reply · E edit · e react · v image"
            }
            TeamsFocus::Composer if app.teams.editing.is_some() => "Enter save · Esc cancel",
            TeamsFocus::Composer => "Enter send · Ctrl+V image · @path Tab · Esc leave",
        },
    }
}

fn pane_direction(app: &App) -> Direction {
    if app.panes_vertical {
        Direction::Vertical
    } else {
        Direction::Horizontal
    }
}

fn outlook_constraints(app: &App) -> [Constraint; 3] {
    if app.panes_vertical {
        let folder_h = app
            .outlook
            .folder_height
            .unwrap_or_else(|| folder_panel_height(&app.outlook.folders));
        [
            Constraint::Length(folder_h),
            Constraint::Percentage(35),
            Constraint::Min(8),
        ]
    } else {
        [
            Constraint::Length(
                app.outlook
                    .folder_width
                    .unwrap_or(DEFAULT_FOLDER_PANEL_WIDTH),
            ),
            Constraint::Percentage(40),
            Constraint::Min(20),
        ]
    }
}

fn teams_list_constraint(app: &App) -> Constraint {
    if app.panes_vertical {
        let n = match app.teams.mode {
            TeamsMode::Chats => app.teams.chats.len(),
            TeamsMode::Channels if app.teams.channels.is_empty() => app.teams.teams.len(),
            TeamsMode::Channels => app.teams.channels.len(),
        };
        Constraint::Length((n.clamp(4, 12) as u16).saturating_add(2))
    } else {
        Constraint::Length(32)
    }
}

// ---------------------------------------------------------------------------
// Outlook
// ---------------------------------------------------------------------------

fn render_outlook(f: &mut Frame, area: Rect, app: &mut App) {
    let cols = Layout::default()
        .direction(pane_direction(app))
        .constraints(outlook_constraints(app))
        .split(area);

    // Folders
    let items: Vec<ListItem> = app
        .outlook
        .folders
        .iter()
        .map(|folder| ListItem::new(mail_folder_label(folder)))
        .collect();
    let mut fstate = ListState::default();
    fstate.select(Some(app.outlook.folder_sel));
    f.render_stateful_widget(
        selectable_list(items, "Folders", app.outlook_focus == OutlookFocus::Folders),
        cols[0],
        &mut fstate,
    );

    // Messages
    let msgs: Vec<ListItem> = app
        .outlook
        .message_rows
        .iter()
        .filter_map(|&i| app.outlook.messages.get(i))
        .map(|m| {
            let members = app.outlook.messages.iter().filter(|candidate| {
                if app.outlook.threaded {
                    mail_same_thread(candidate, m)
                } else {
                    candidate.id == m.id
                }
            });
            let unread = mail_row_unread(&app.outlook.messages, m, app.outlook.threaded);
            let marker = if unread { "●" } else { " " };
            let clip = if members
                .clone()
                .any(|message| message.has_attachments.unwrap_or(false))
            {
                "📎"
            } else {
                ""
            };
            let mut subject = m.subject.clone().unwrap_or_else(|| "(no subject)".into());
            if app.outlook.threaded {
                let count = mail_thread_size(&app.outlook.messages, m);
                if count > 1 {
                    subject.push_str(&format!(" [{count}]"));
                }
            }
            let line = Line::from(vec![
                Span::styled(format!("{marker} "), Style::default().fg(ACCENT)),
                Span::styled(
                    truncate(&m.sender_name(), 18),
                    Style::default().fg(Color::LightGreen),
                ),
                Span::raw("  "),
                Span::styled(clip.to_string(), Style::default().fg(DIM)),
                Span::styled(
                    subject,
                    if unread {
                        Style::default().add_modifier(Modifier::BOLD)
                    } else {
                        Style::default()
                    },
                ),
            ]);
            ListItem::new(line)
        })
        .collect();
    let mut mstate = ListState::default();
    mstate.select(Some(app.outlook.msg_sel));
    let kind = if app.outlook.threaded {
        "Threads"
    } else {
        "Messages"
    };
    let count = app.outlook.message_rows.len();
    let msg_title = if app.outlook.messages_next.is_some() {
        format!("{kind} ({count} · ↓ for more)")
    } else {
        format!("{kind} ({count})")
    };
    f.render_stateful_widget(
        selectable_list(
            msgs,
            &msg_title,
            app.outlook_focus == OutlookFocus::Messages,
        ),
        cols[1],
        &mut mstate,
    );

    // Reading pane — scrollable when focused, like the Teams conversation.
    let focused = app.outlook_focus == OutlookFocus::Reading;
    let title = if focused {
        "Reading (j/k scroll · g/G · Esc back)"
    } else {
        "Reading"
    };
    let block = panel_block(title, focused);
    let inner = block.inner(cols[2]);
    f.render_widget(block, cols[2]);

    match email_flow(app) {
        Some(flow) => {
            let (rows, _) = wrap_flow(app, &flow, inner.width.max(1) as usize);
            // Tell the key handler how far it can usefully scroll.
            app.reading_max_scroll
                .set((display_row_count(&rows) as u16).saturating_sub(inner.height));
            let scroll = app.outlook.reading_scroll.min(app.reading_max_scroll.get());
            render_display_rows(f, inner, &rows, scroll, app);
        }
        None => {
            app.reading_max_scroll.set(0);
            f.render_widget(
                // Keys live in the bottom bar; keep the pane itself uncluttered.
                Paragraph::new("Select a message and press Enter to read.")
                    .wrap(Wrap { trim: false })
                    .style(Style::default().fg(DIM)),
                inner,
            );
        }
    }
}

fn email_flow(app: &App) -> Option<Vec<FlowItem>> {
    let message = app.outlook.reading.as_ref()?;
    if app.outlook.reading_thread.is_empty() {
        let mut flow = email_header_lines(app, message, None, 0, 1)
            .into_iter()
            .map(FlowItem::Line)
            .collect();
        if let Some(body) = &app.outlook.reading_body {
            push_mail_body(&mut flow, app, &message.id, body.lines.as_ref());
        }
        return Some(flow);
    }

    let total = app.outlook.reading_thread.len();
    let mut flow = Vec::new();
    for (index, (message, body)) in app
        .outlook
        .reading_thread
        .iter()
        .zip(&app.outlook.reading_thread_bodies)
        .enumerate()
    {
        if index > 0 {
            flow.push(FlowItem::Line(Line::raw("")));
        }
        flow.extend(
            email_header_lines(app, message, Some(index), total, index)
                .into_iter()
                .map(FlowItem::Line),
        );
        push_mail_body(&mut flow, app, &message.id, body.lines.as_ref());
    }
    Some(flow)
}

fn email_header_lines(
    app: &App,
    message: &m365_core::models::MailMessage,
    index: Option<usize>,
    total: usize,
    attach_on: usize,
) -> Vec<Line<'static>> {
    let mut lines = Vec::new();
    if let Some(i) = index {
        lines.push(Line::from(Span::styled(
            format!("── Message {} of {total} ──", i + 1),
            Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
        )));
    }
    lines.push(kv("Subject", &message.subject.clone().unwrap_or_default()));
    lines.push(kv("From", &message.sender_name()));
    lines.push(kv("Date", message.mail_time().unwrap_or_default()));
    if attach_on == 0 && !app.outlook.reading_attachments.is_empty() {
        lines.push(attachment_line(app));
    }
    lines.push(Line::raw(""));
    lines
}

fn push_mail_body(flow: &mut Vec<FlowItem>, app: &App, message_id: &str, lines: &[Line<'static>]) {
    let images: Vec<crate::content::BodyImage> = app
        .outlook
        .reading_images
        .iter()
        .filter(|image| image.message_id == message_id)
        .map(|image| image.image.clone())
        .collect();
    for piece in crate::content::body_pieces(lines, &images) {
        match piece {
            crate::content::BodyPiece::Line(line) => flow.push(FlowItem::Line(line.clone())),
            crate::content::BodyPiece::Image(image) if mail_image_is_local(&image.src) => {
                flow.push(FlowItem::Image {
                    key: crate::termimg::mail_cache_key(message_id, &image.src),
                });
            }
            crate::content::BodyPiece::Image(_) => {}
        }
    }
}

/// Headers + rendered body of the open email, or `None` if nothing is open.
/// Shared by the reading pane and copy mode.
pub fn email_lines(app: &App) -> Option<Vec<Line<'static>>> {
    let m = app.outlook.reading.as_ref()?;
    if !app.outlook.reading_thread.is_empty() {
        let total = app.outlook.reading_thread.len();
        let mut lines = Vec::new();
        for (i, (message, body)) in app
            .outlook
            .reading_thread
            .iter()
            .zip(&app.outlook.reading_thread_bodies)
            .enumerate()
        {
            if i > 0 {
                lines.push(Line::raw(""));
            }
            lines.push(Line::from(Span::styled(
                format!("── Message {} of {total} ──", i + 1),
                Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
            )));
            lines.push(kv("Subject", &message.subject.clone().unwrap_or_default()));
            lines.push(kv("From", &message.sender_name()));
            lines.push(kv("Date", message.mail_time().unwrap_or_default()));
            if i == 0 && !app.outlook.reading_attachments.is_empty() {
                lines.push(attachment_line(app));
            }
            lines.push(Line::raw(""));
            lines.extend(body.lines.iter().cloned());
        }
        return Some(lines);
    }
    let mut lines = vec![
        kv("Subject", &m.subject.clone().unwrap_or_default()),
        kv("From", &m.sender_name()),
        kv("Date", m.mail_time().unwrap_or_default()),
        Line::raw(""),
    ];
    if !app.outlook.reading_attachments.is_empty() {
        lines.insert(3, attachment_line(app));
    }
    if let Some(body) = &app.outlook.reading_body {
        lines.extend(body.lines.iter().cloned());
    }
    Some(lines)
}

fn attachment_line(app: &App) -> Line<'static> {
    let names: Vec<String> = app
        .outlook
        .reading_attachments
        .iter()
        .map(|a| format!("{} ({})", a.display_name(), a.human_size()))
        .collect();
    kv(
        "Attach",
        &format!("📎 {}  — press A to save", names.join(", ")),
    )
}

enum FlowItem {
    Line(Line<'static>),
    Image {
        key: String,
    },
    LineWithPhoto {
        line: Line<'static>,
        photo_key: String,
    },
}

enum DisplayRow {
    Line(Line<'static>),
    Image {
        key: String,
        height: u16,
    },
    LineWithPhoto {
        lines: Vec<Line<'static>>,
        photo_key: String,
        photo_width: u16,
        height: u16,
    },
}

impl DisplayRow {
    fn height(&self) -> u16 {
        match self {
            DisplayRow::Line(_) => 1,
            DisplayRow::Image { height, .. } | DisplayRow::LineWithPhoto { height, .. } => {
                (*height).max(1)
            }
        }
    }
}

/// Lines of the open Teams conversation, plus the starting line index of each
/// message. `selectable` adds the `▶` cursor and selection highlight (off in
/// copy mode so the text copies cleanly).
pub fn conversation_lines(app: &App, selectable: bool) -> (Vec<Line<'static>>, Vec<usize>) {
    let (flow, starts) = conversation_flow(app, selectable);
    let lines = flow
        .into_iter()
        .map(|item| match item {
            FlowItem::Line(l) | FlowItem::LineWithPhoto { line: l, .. } => l,
            FlowItem::Image { .. } => Line::from("[image]"),
        })
        .collect();
    (lines, starts)
}

fn conversation_flow(app: &App, selectable: bool) -> (Vec<FlowItem>, Vec<usize>) {
    let mut flow: Vec<FlowItem> = Vec::new();
    let mut starts: Vec<usize> = Vec::with_capacity(app.teams.messages.len());
    // Emit a "Today"/"Yesterday"/date separator whenever the day changes.
    let mut last_day: Option<chrono::NaiveDate> = None;
    // Track the previous message so consecutive ones from the same person can
    // share a single author header.
    let mut prev: Option<(String, Option<chrono::DateTime<chrono::Local>>)> = None;

    for (i, m) in app.teams.messages.iter().enumerate() {
        let when = local_time(m.created_date_time.as_deref());
        let mut day_changed = false;
        if let Some(when) = when {
            let day = when.date_naive();
            if last_day != Some(day) {
                if last_day.is_some() {
                    flow.push(FlowItem::Line(Line::from("")));
                }
                flow.push(FlowItem::Line(day_separator(&day_label(day))));
                last_day = Some(day);
                day_changed = true;
            }
        }
        // Record the start *after* any separator, so scrolling to a message
        // puts the message itself at the top — the pinned header carries the
        // date, and we avoid showing the same date twice.
        starts.push(flow.len());
        let selected = selectable && i == app.teams.msg_sel;
        let marker = if !selectable {
            ""
        } else if selected {
            "▶ "
        } else {
            "  "
        };

        // Every message opens with its own local time, so a run sharing one
        // author header still shows when each line was sent. Wrapped body lines
        // line up past that gutter.
        let ts = when
            .map(|w| w.format("%H:%M").to_string())
            .unwrap_or_else(|| " ".repeat(TIME_WIDTH));
        let gutter = " ".repeat(marker.chars().count() + TIME_WIDTH + 1);
        let lead = |extra: Vec<Span<'static>>| {
            let mut spans = vec![
                Span::styled(marker.to_string(), Style::default().fg(ACCENT)),
                Span::styled(format!("{ts} "), Style::default().fg(DIM)),
            ];
            spans.extend(extra);
            Line::from(spans)
        };

        if m.deleted_date_time.is_some() {
            flow.push(FlowItem::Line(lead(vec![Span::styled(
                "(message deleted)",
                Style::default().fg(DIM),
            )])));
            prev = None; // a deletion breaks the run
            continue;
        }

        let author = m.author();
        let grouped = !day_changed
            && prev
                .as_ref()
                .is_some_and(|(a, t)| continues_run(a, *t, &author, when));

        let body: Vec<Line<'static>> = app
            .teams
            .messages_rendered
            .get(i)
            .map(|b| b.lines.clone())
            .unwrap_or_default();
        let imgs = app
            .teams
            .messages_images
            .get(i)
            .cloned()
            .unwrap_or_default();
        let mut pieces: Vec<crate::content::BodyPiece<'_>> =
            crate::content::body_pieces(&body, &imgs);

        // A reply carries the message it answers as a `messageReference`
        // attachment, not as HTML, so it has to be drawn explicitly — and it
        // has to come *before* the reply text to read correctly.
        let quote = m.quoted().map(|q| {
            vec![
                Span::styled("┃ ", Style::default().fg(ACCENT)),
                Span::styled(
                    format!("{}: ", q.author),
                    Style::default().fg(Color::LightGreen),
                ),
                Span::styled(truncate(&q.preview, 70), Style::default().fg(DIM)),
            ]
        });

        match (grouped, quote) {
            // Grouped reply: the quote takes the lead line, the text follows.
            (true, Some(quote)) => flow.push(FlowItem::Line(lead(quote))),
            // Grouped message: the text starts right after the time.
            (true, None) => {
                let first = match pieces.first() {
                    Some(crate::content::BodyPiece::Line(line)) => {
                        let spans = line.spans.clone();
                        pieces.remove(0);
                        spans
                    }
                    _ => Vec::new(),
                };
                flow.push(FlowItem::Line(lead(first)));
            }
            // New author: name on the lead line, then the quote if there is one.
            (false, quote) => {
                let name = lead(vec![Span::styled(
                    author.clone(),
                    Style::default()
                        .fg(if selected {
                            Color::Cyan
                        } else {
                            Color::LightGreen
                        })
                        .add_modifier(Modifier::BOLD),
                )]);
                if let Some(id) = m.author_id().filter(|id| !id.is_empty()) {
                    flow.push(FlowItem::LineWithPhoto {
                        line: name,
                        photo_key: crate::termimg::photo_cache_key(id),
                    });
                } else {
                    flow.push(FlowItem::Line(name));
                }
                if let Some(quote) = quote {
                    let mut spans = vec![Span::raw(gutter.clone())];
                    spans.extend(quote);
                    flow.push(FlowItem::Line(Line::from(spans)));
                }
            }
        }

        for piece in pieces {
            match piece {
                crate::content::BodyPiece::Line(line) => {
                    let mut spans = vec![Span::raw(gutter.clone())];
                    spans.extend(line.spans.clone());
                    flow.push(FlowItem::Line(Line::from(spans)));
                }
                crate::content::BodyPiece::Image(img) => {
                    let hosted = crate::termimg::hosted_content_id(&img.src);
                    let key = crate::termimg::cache_key(&img.src, hosted);
                    flow.push(FlowItem::Image { key });
                }
            }
        }
        for att in &m.attachments {
            if let Some(name) = &att.name {
                flow.push(FlowItem::Line(Line::from(vec![
                    Span::raw(gutter.clone()),
                    Span::styled(format!("📎 {name}"), Style::default().fg(Color::LightBlue)),
                ])));
            }
        }
        if let Some(reactions) = m.reactions_summary() {
            flow.push(FlowItem::Line(Line::from(vec![
                Span::raw(gutter.clone()),
                Span::styled(reactions, Style::default().fg(DIM)),
            ])));
        }
        if m.was_edited() {
            flow.push(FlowItem::Line(Line::from(vec![
                Span::raw(gutter.clone()),
                Span::styled(
                    "edited",
                    Style::default().fg(DIM).add_modifier(Modifier::ITALIC),
                ),
            ])));
        }
        prev = Some((author, when));
    }
    (flow, starts)
}

fn wrap_flow(app: &App, items: &[FlowItem], width: usize) -> (Vec<DisplayRow>, Vec<usize>) {
    let mut rows = Vec::new();
    let mut starts = Vec::with_capacity(items.len());
    for item in items {
        starts.push(display_row_count(&rows));
        match item {
            FlowItem::Line(l) => {
                for row in crate::wrap::wrap_line(l, width) {
                    rows.push(DisplayRow::Line(row));
                }
            }
            FlowItem::Image { key } => rows.push(DisplayRow::Image {
                key: key.clone(),
                height: app.image_display_rows(key, width as u16),
            }),
            FlowItem::LineWithPhoto { line, photo_key } => match app.avatar_cell_size(photo_key) {
                Some((pw, ph)) if (pw as usize) + 1 < width => {
                    let text_w = width - pw as usize - 1;
                    let wrapped = crate::wrap::wrap_line(line, text_w.max(1));
                    let height = (wrapped.len() as u16).max(ph).max(1);
                    rows.push(DisplayRow::LineWithPhoto {
                        lines: wrapped,
                        photo_key: photo_key.clone(),
                        photo_width: pw,
                        height,
                    });
                }
                _ => {
                    for row in crate::wrap::wrap_line(line, width) {
                        rows.push(DisplayRow::Line(row));
                    }
                }
            },
        }
    }
    (rows, starts)
}

fn display_row_count(rows: &[DisplayRow]) -> usize {
    rows.iter().map(|r| r.height() as usize).sum()
}

fn render_display_rows(
    f: &mut Frame,
    area: ratatui::layout::Rect,
    rows: &[DisplayRow],
    scroll: u16,
    app: &mut App,
) {
    let mut y_off = 0u16;
    for row in rows {
        let h = row.height();
        let start = y_off;
        y_off = y_off.saturating_add(h);
        if y_off <= scroll || start >= scroll.saturating_add(area.height) {
            continue;
        }
        if start < scroll {
            continue;
        }
        let dest_y = area.y + (start - scroll);
        if dest_y >= area.bottom() {
            break;
        }
        let draw_h = h.min(area.bottom().saturating_sub(dest_y));
        let dest = ratatui::layout::Rect {
            x: area.x,
            y: dest_y,
            width: area.width,
            height: draw_h,
        };
        match row {
            DisplayRow::Line(l) => {
                f.render_widget(Paragraph::new(l.clone()), dest);
            }
            DisplayRow::Image { key, .. } => {
                if let Some(img) = app.image_cache.get_mut(key) {
                    crate::termimg::render(f, dest, img);
                } else {
                    let label = if app.image_is_pending(key) {
                        "  [image…]"
                    } else {
                        "  [image unavailable]"
                    };
                    f.render_widget(
                        Paragraph::new(Span::styled(label, Style::default().fg(DIM))),
                        dest,
                    );
                }
            }
            DisplayRow::LineWithPhoto {
                lines,
                photo_key,
                photo_width,
                ..
            } => {
                let text_w = dest.width.saturating_sub(*photo_width);
                for (i, line) in lines.iter().enumerate() {
                    let y = dest.y.saturating_add(i as u16);
                    if y >= dest.bottom() {
                        break;
                    }
                    f.render_widget(
                        Paragraph::new(line.clone()),
                        Rect {
                            x: dest.x,
                            y,
                            width: text_w.max(1),
                            height: 1,
                        },
                    );
                }
                let name_w = lines.first().map(line_width).unwrap_or(0);
                let img_x = dest.x.saturating_add(name_w.saturating_add(1));
                if *photo_width > 0 && img_x < dest.right() {
                    let img_area = Rect {
                        x: img_x,
                        y: dest.y,
                        width: (*photo_width).min(dest.right().saturating_sub(img_x)),
                        height: dest.height,
                    };
                    if let Some(img) = app.image_cache.get_mut(photo_key) {
                        crate::termimg::render(f, img_area, img);
                    }
                }
            }
        }
    }
}

/// Width of the `HH:MM` timestamp column.
const TIME_WIDTH: usize = 5;

/// Whether a message continues the previous one's run: same author, and close
/// enough in time that repeating the name would just be noise.
fn continues_run(
    prev_author: &str,
    prev_at: Option<chrono::DateTime<chrono::Local>>,
    author: &str,
    at: Option<chrono::DateTime<chrono::Local>>,
) -> bool {
    if prev_author != author {
        return false;
    }
    match (prev_at, at) {
        // Messages are newest-first, so the gap can run either way.
        (Some(a), Some(b)) => (a - b).num_minutes().abs() <= RUN_GAP_MINUTES,
        _ => true,
    }
}

/// A pause this long starts a fresh header even for the same person.
const RUN_GAP_MINUTES: i64 = 15;

// ---------------------------------------------------------------------------
// Teams
// ---------------------------------------------------------------------------

fn render_teams(f: &mut Frame, area: Rect, app: &mut App) {
    let cols = Layout::default()
        .direction(pane_direction(app))
        .constraints([teams_list_constraint(app), Constraint::Min(20)])
        .split(area);

    let me_id = app.me.as_ref().map(|m| m.id.as_str());

    // Left list: chats or channels
    let (title, items, sel): (&str, Vec<ListItem>, usize) = match app.teams.mode {
        TeamsMode::Chats => {
            let items = app
                .teams
                .chats
                .iter()
                .map(|c| ListItem::new(truncate(&c.label(me_id), 30)))
                .collect();
            ("Chats (t→channels)", items, app.teams.chat_sel)
        }
        TeamsMode::Channels => {
            if app.teams.channels.is_empty() {
                let items = app
                    .teams
                    .teams
                    .iter()
                    .map(|t| ListItem::new(truncate(t.display_name.as_deref().unwrap_or(""), 30)))
                    .collect();
                ("Teams (Enter→channels)", items, app.teams.team_sel)
            } else {
                let items = app
                    .teams
                    .channels
                    .iter()
                    .map(|c| ListItem::new(truncate(c.display_name.as_deref().unwrap_or(""), 30)))
                    .collect();
                ("Channels (t→chats)", items, app.teams.channel_sel)
            }
        }
    };
    let mut lstate = ListState::default();
    lstate.select(Some(sel));
    f.render_stateful_widget(
        selectable_list(items, title, app.teams.focus == TeamsFocus::List),
        cols[0],
        &mut lstate,
    );

    // Right: messages + composer. The composer grows with its content (handy for
    // multi-line pastes) up to a cap; Min(5) leaves room for the border, the
    // pinned date header, and a couple of message rows on small terminals.
    let composer_width = cols[1].width.saturating_sub(2).max(1) as usize;
    let composer_rows = app.teams.composer.wrap(composer_width).len().clamp(1, 6) as u16;
    // One extra row while a reply or edit is in progress, for the banner.
    let reply_row = u16::from(app.teams.replying_to.is_some() || app.teams.editing.is_some());
    let preview_h = composer_preview_height(app, composer_width as u16);
    let right = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Min(5),
            Constraint::Length(composer_rows + reply_row + preview_h + 2),
        ])
        .split(cols[1]);

    let focused = app.teams.focus == TeamsFocus::Messages;
    let (mut flow, msg_starts) = conversation_flow(app, focused);
    if flow.is_empty() {
        flow.push(FlowItem::Line(Line::styled(
            "Select a conversation and press Enter.",
            Style::default().fg(DIM),
        )));
    }
    // Wrap here rather than letting `Paragraph` do it: scrolling needs the exact
    // row count, and `Paragraph` won't report one. Estimating it undercounts —
    // words don't fill a row — which left the newest messages below the edge.
    let inner_w = right[0].width.saturating_sub(2).max(1) as usize;
    let pane_h = right[0].height.saturating_sub(3).max(1) as usize; // borders + date header
    let (rows, row_of_item) = wrap_flow(app, &flow, inner_w);
    let total_h = display_row_count(&rows);
    // Where each message begins, in rendered rows.
    let msg_rows: Vec<usize> = msg_starts
        .iter()
        .map(|&l| row_of_item.get(l).copied().unwrap_or(total_h))
        .collect();
    let sel_end = msg_rows
        .get(app.teams.msg_sel + 1)
        .copied()
        .unwrap_or(total_h);
    let scroll = sel_end
        .saturating_sub(pane_h)
        .min(total_h.saturating_sub(pane_h)) as u16;
    // Flag messages that arrived while the user was reading further back.
    let title = if app.teams.unseen > 0 {
        format!("Conversation — ▼ {} new (G to jump)", app.teams.unseen)
    } else if focused {
        "Conversation (j/k select · g/G · e react · v image · z copy-mode)".to_string()
    } else {
        "Conversation".to_string()
    };

    // The pane is split inside its border: a pinned date header on the first
    // row, then the scrolling message flow. The header tracks the day of the
    // topmost visible message, so it updates as you scroll.
    let block = panel_block(&title, focused);
    let inner = block.inner(right[0]);
    f.render_widget(block, right[0]);
    let pane = Layout::default()
        .direction(Direction::Vertical)
        .constraints([Constraint::Length(1), Constraint::Min(0)])
        .split(inner);

    if let Some(label) = sticky_day_label(app, &msg_rows, scroll) {
        f.render_widget(Paragraph::new(day_separator(&label)), pane[0]);
    }
    render_display_rows(f, pane[1], &rows, scroll, app);

    let composing = app.teams.focus == TeamsFocus::Composer;
    let title = if composing && app.teams.editing.is_some() {
        "Message (Enter save · Esc cancel)"
    } else if composing {
        "Message (Enter send · Ctrl+V image · @path Tab)"
    } else {
        "Message"
    };
    let composer_block = panel_block(title, composing);
    let mut composer_inner = composer_block.inner(right[1]);
    f.render_widget(composer_block, right[1]);

    // Show what's being replied to or edited, so Enter isn't a surprise.
    if app.teams.editing.is_some() {
        let banner = Rect {
            height: 1,
            ..composer_inner
        };
        composer_inner = Rect {
            y: composer_inner.y + 1,
            height: composer_inner.height.saturating_sub(1),
            ..composer_inner
        };
        f.render_widget(
            Paragraph::new(Line::from(vec![Span::styled(
                "┃ editing message",
                Style::default().fg(ACCENT),
            )])),
            banner,
        );
    } else if let Some(idx) = app.teams.replying_to {
        let banner = Rect {
            height: 1,
            ..composer_inner
        };
        composer_inner = Rect {
            y: composer_inner.y + 1,
            height: composer_inner.height.saturating_sub(1),
            ..composer_inner
        };
        let who = app
            .teams
            .messages
            .get(idx)
            .map(|m| m.author())
            .unwrap_or_default();
        let excerpt = app
            .teams
            .messages_rendered
            .get(idx)
            .map(|t| truncate(&crate::content::plain(t).replace('\n', " "), 60))
            .unwrap_or_default();
        f.render_widget(
            Paragraph::new(Line::from(vec![
                Span::styled("┃ replying to ", Style::default().fg(ACCENT)),
                Span::styled(who, Style::default().fg(Color::LightGreen)),
                Span::styled(format!(": {excerpt}"), Style::default().fg(DIM)),
            ])),
            banner,
        );
    }

    render_composer_previews(f, &mut composer_inner, app);

    app.text_width_hint
        .set(composer_inner.width.max(1) as usize);
    if app.teams.composer.is_empty() && app.teams.images.is_empty() && !composing {
        f.render_widget(
            Paragraph::new(Span::styled(
                "press i to type · Enter to send",
                Style::default().fg(DIM),
            )),
            composer_inner,
        );
    } else if let Some((x, y)) = render_text_area(f, composer_inner, &app.teams.composer, composing)
    {
        f.set_cursor_position((x, y));
    }
}

// ---------------------------------------------------------------------------
// Overlays
// ---------------------------------------------------------------------------

fn composer_preview_height(app: &App, width: u16) -> u16 {
    if app.teams.images.is_empty() {
        return 0;
    }
    if app.teams.composer_previews.is_empty() {
        return app.teams.images.len() as u16;
    }
    app.teams
        .composer_previews
        .iter()
        .map(|preview| {
            preview
                .as_ref()
                .map(|image| image.rows_for_width(width))
                .unwrap_or(1)
        })
        .sum()
}

fn render_composer_previews(f: &mut Frame, composer_inner: &mut Rect, app: &mut App) {
    if app.teams.images.is_empty() {
        return;
    }
    if app.teams.composer_previews.is_empty() {
        let n = app.teams.images.len() as u16;
        let banner = Rect {
            height: n.min(composer_inner.height),
            ..*composer_inner
        };
        *composer_inner = Rect {
            y: composer_inner.y + banner.height,
            height: composer_inner.height.saturating_sub(banner.height),
            ..*composer_inner
        };
        let labels: Vec<Line> = app
            .teams
            .images
            .iter()
            .map(|img| {
                Line::from(Span::styled(
                    format!("[image unavailable] {}", img.name),
                    Style::default().fg(DIM),
                ))
            })
            .collect();
        f.render_widget(Paragraph::new(labels), banner);
        return;
    }
    let n = app.teams.composer_previews.len();
    for i in 0..n {
        let rows = app.teams.composer_previews[i]
            .as_ref()
            .map(|image| image.rows_for_width(composer_inner.width))
            .unwrap_or(1)
            .min(composer_inner.height);
        if rows == 0 {
            break;
        }
        let dest = Rect {
            height: rows,
            ..*composer_inner
        };
        *composer_inner = Rect {
            y: composer_inner.y + rows,
            height: composer_inner.height.saturating_sub(rows),
            ..*composer_inner
        };
        if let Some(img) = app.teams.composer_previews[i].as_mut() {
            crate::termimg::render(f, dest, img);
        } else {
            f.render_widget(
                Paragraph::new(Span::styled(
                    "[image unavailable]",
                    Style::default().fg(DIM),
                )),
                dest,
            );
        }
    }
}

fn render_view_image(f: &mut Frame, app: &mut App, keys: &[String], sel: usize, back: bool) {
    let area = f.area();
    f.render_widget(Clear, area);
    let esc = if back { "Esc back" } else { "Esc close" };
    let hint = if keys.len() > 1 {
        format!(
            " {} / {}  · j/k next · {esc} ",
            sel.saturating_add(1).min(keys.len()),
            keys.len()
        )
    } else {
        format!(" {esc} ")
    };
    let hint_h = u16::from(area.height > 0);
    let img_area = Rect {
        height: area.height.saturating_sub(hint_h),
        ..area
    };
    if hint_h == 1 {
        f.render_widget(
            Paragraph::new(Span::styled(hint, Style::default().fg(DIM))),
            Rect {
                x: area.x,
                y: area.bottom().saturating_sub(1),
                width: area.width,
                height: 1,
            },
        );
    }
    let Some(key) = keys.get(sel) else {
        return;
    };
    if let Some(img) = app.image_cache.get_mut(key) {
        let (w, h) = img.fit_contain(img_area.width, img_area.height);
        let dest = Rect {
            x: img_area.x + img_area.width.saturating_sub(w) / 2,
            y: img_area.y + img_area.height.saturating_sub(h) / 2,
            width: w,
            height: h,
        };
        crate::termimg::render(f, dest, img);
    } else {
        let label = if app.image_is_pending(key) {
            "[image…]"
        } else {
            "[image unavailable]"
        };
        f.render_widget(
            Paragraph::new(Span::styled(label, Style::default().fg(DIM))),
            img_area,
        );
    }
}

fn render_mail_images(f: &mut Frame, app: &mut App, scroll: u16, fullscreen: Option<usize>) {
    let entries = app.reading_image_entries();
    if let Some(sel) = fullscreen {
        let keys: Vec<String> = entries.into_iter().map(|(key, _)| key).collect();
        render_view_image(f, app, &keys, sel, true);
        return;
    }

    let area = f.area();
    f.render_widget(Clear, area);
    let block = popup_block("Images — press 1-9 to fullscreen · j/k scroll · Esc close");
    let inner = block.inner(area);
    f.render_widget(block, area);
    if inner.width == 0 || inner.height == 0 {
        app.images_max_scroll.set(0);
        return;
    }

    let preview_cap = crate::termimg::MAIL_PREVIEW_ROWS;
    let mut heights: Vec<u16> = Vec::with_capacity(entries.len());
    for (i, (key, _)) in entries.iter().enumerate() {
        let img_h = app.image_display_rows(key, inner.width).min(preview_cap);
        let gap = u16::from(i + 1 < entries.len());
        heights.push(1 + img_h + gap);
    }
    let total: u16 = heights.iter().copied().sum();
    let max = total.saturating_sub(inner.height);
    app.images_max_scroll.set(max);
    let scroll = scroll.min(max);

    let mut y_off = 0u16;
    for (i, ((key, alt), h)) in entries.iter().zip(&heights).enumerate() {
        let start = y_off;
        y_off = y_off.saturating_add(*h);
        if y_off <= scroll || start >= scroll.saturating_add(inner.height) {
            continue;
        }
        if start < scroll {
            continue;
        }
        let dest_y = inner.y + (start - scroll);
        if dest_y >= inner.bottom() {
            break;
        }
        let n = i + 1;
        let caption = if alt.is_empty() {
            "image"
        } else {
            alt.as_str()
        };
        f.render_widget(
            Paragraph::new(Line::from(vec![
                Span::styled(
                    format!("{n} "),
                    Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
                ),
                Span::raw(caption.to_string()),
            ])),
            Rect {
                x: inner.x,
                y: dest_y,
                width: inner.width,
                height: 1,
            },
        );
        let img_y = dest_y.saturating_add(1);
        if img_y >= inner.bottom() {
            continue;
        }
        let img_h = app.image_display_rows(key, inner.width).min(preview_cap);
        let draw_h = img_h.min(inner.bottom().saturating_sub(img_y));
        if draw_h == 0 {
            continue;
        }
        let img_area = Rect {
            x: inner.x,
            y: img_y,
            width: inner.width,
            height: draw_h,
        };
        if let Some(img) = app.image_cache.get_mut(key) {
            crate::termimg::render(f, img_area, img);
        } else {
            let label = if app.image_is_pending(key) {
                "[image…]"
            } else {
                "[image unavailable]"
            };
            f.render_widget(
                Paragraph::new(Span::styled(label, Style::default().fg(DIM))),
                img_area,
            );
        }
    }
}

fn render_overlay(f: &mut Frame, app: &mut App) {
    if let Some(Overlay::ViewImage { keys, sel }) = &app.overlay {
        let keys = keys.clone();
        let sel = *sel;
        render_view_image(f, app, &keys, sel, false);
        return;
    }
    if let Some(Overlay::MailImages { scroll, fullscreen }) = &app.overlay {
        let scroll = *scroll;
        let fullscreen = *fullscreen;
        render_mail_images(f, app, scroll, fullscreen);
        return;
    }
    let Some(overlay) = &app.overlay else {
        return;
    };
    match overlay {
        Overlay::ViewImage { .. } | Overlay::MailImages { .. } => {}
        Overlay::Help => {
            let area = centered(60, 60, f.area());
            f.render_widget(Clear, area);
            let text = "\
 M365 TUI — keys\n\
 \n\
 Global:  F2 switch app · Ctrl+P palette · p presence · s settings · | panes · ? help · q quit\n\
 \n\
 Links:   o list links in the message · 1-9 open in browser\n\
 Images:  i list images in the open mail · 1-9 fullscreen\n\
 Attach:  A list attachments · 1-9 save to your Downloads folder\n\
          when writing: Tab to Attach, type a path, Tab complete, Enter attach\n\
 \n\
 Copying: y yank focused message · Y yank whole view\n\
          z copy mode (full-width, borderless — drag-select cleanly)\n\
 \n\
 Moving:  h/← out a pane · l/→ into it (opens what's selected)\n\
          j/k or ↑/↓ move · g/G top/bottom · arrows work everywhere hjkl does\n\
          | or \\ toggle horizontal/vertical panes\n\
          Outlook: Shift+H/L resize the Folders panel\n\
 \n\
 Outlook: Enter open · c compose · r reply · a reply-all · f forward\n\
           u read/unread · m move folder · d trash · / search · e calendar\n\
           t toggles threads/individual messages\n\
          folders pane / finds a folder · reading pane j/k scroll · g/G\n\
 \n\
  Teams:   n new chat (chat list) · t chats/channels · j/k select message · g oldest · G newest · e react · v full image\n\
          i type · r reply · E edit (Up in empty composer = last of yours) · Enter send\n\
          Ctrl+V paste image · @path Tab complete image · Ctrl+X remove last image\n\
 \n\
 Compose: To/Cc/Bcc autocomplete from seen mail · ↑/↓ + Tab/Enter choose\n\
          Tab/Shift+Tab field · Ctrl+S send · Esc cancel\n\
          replies: Ctrl+R reply/reply-all · Ctrl+E edit recipients\n\
          body: @path Tab completes and embeds an inline image\n\
          Ctrl+X e $EDITOR (body + subject) · Ctrl+X x unstage last file\n\
          ←→↑↓ move · Ctrl+←→ by word · Home/End line · Ctrl+Home/End all\n\
          Backspace/Delete · Ctrl+W word · Ctrl+U to line start · Ctrl+K to end\n\
          Enter newline in body · paste works (bracketed paste)\n\
 \n\
 Press Esc to close.";
            f.render_widget(
                Paragraph::new(text)
                    .block(popup_block("Help"))
                    .wrap(Wrap { trim: false }),
                area,
            );
        }
        Overlay::Calendar => {
            let area = centered(70, 70, f.area());
            f.render_widget(Clear, area);
            let items: Vec<ListItem> = app
                .outlook
                .calendar
                .iter()
                .map(|e| {
                    let start = e
                        .start
                        .as_ref()
                        .map(|s| s.date_time.replace('T', " "))
                        .unwrap_or_default();
                    let subj = e.subject.clone().unwrap_or_default();
                    let online = if e.is_online_meeting.unwrap_or(false) {
                        " 🔗"
                    } else {
                        ""
                    };
                    ListItem::new(format!("{start}  {subj}{online}"))
                })
                .collect();
            let list = if items.is_empty() {
                List::new(vec![ListItem::new(
                    "No events in the next 7 days (or still loading).",
                )])
            } else {
                List::new(items)
            };
            f.render_widget(
                list.block(popup_block("Calendar — next 7 days (Esc to close)")),
                area,
            );
        }
        Overlay::Search { query } => {
            let area = centered(60, 20, f.area());
            f.render_widget(Clear, area);
            f.render_widget(
                Paragraph::new(format!(
                    "Search mail:\n\n> {query}▏\n\nEnter to search · Esc to cancel"
                ))
                .block(popup_block("Search")),
                area,
            );
        }
        Overlay::FolderSearch { query, sel } => {
            let area = centered(50, 60, f.area());
            f.render_widget(Clear, area);
            let matches = filter_folders(&app.outlook.folders, query);
            let block = popup_block("Find folder");
            let inner = block.inner(area);
            f.render_widget(block, area);
            let rows = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Length(2), Constraint::Min(0)])
                .split(inner);
            f.render_widget(Paragraph::new(format!("> {query}▏")), rows[0]);
            let items: Vec<ListItem> = matches
                .iter()
                .filter_map(|&i| app.outlook.folders.get(i))
                .map(|folder| ListItem::new(mail_folder_label(folder)))
                .collect();
            let mut st = ListState::default();
            st.select(Some(*sel));
            f.render_stateful_widget(
                List::new(items).highlight_style(
                    Style::default()
                        .fg(Color::Black)
                        .bg(ACCENT)
                        .add_modifier(Modifier::BOLD),
                ),
                rows[1],
                &mut st,
            );
        }
        Overlay::NewChat {
            query,
            results,
            sel,
            loading,
        } => {
            let area = centered(60, 60, f.area());
            f.render_widget(Clear, area);
            let block = popup_block("New Teams chat — find a person");
            let inner = block.inner(area);
            f.render_widget(block, area);
            let rows = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Length(2), Constraint::Min(0)])
                .split(inner);
            f.render_widget(Paragraph::new(format!("> {query}▏")), rows[0]);
            let items: Vec<ListItem> = if results.is_empty() {
                vec![ListItem::new(if query.is_empty() {
                    "Type a name or username to search the directory"
                } else if *loading {
                    "Searching…"
                } else {
                    "No matching users"
                })]
            } else {
                results
                    .iter()
                    .map(|user| {
                        let name = user.display_name.as_deref().unwrap_or("(no name)");
                        let username = user
                            .user_principal_name
                            .as_deref()
                            .or(user.mail.as_deref())
                            .unwrap_or("");
                        ListItem::new(format!("{name}  <{username}>"))
                    })
                    .collect()
            };
            let mut state = ListState::default();
            if !results.is_empty() {
                state.select(Some(*sel));
            }
            f.render_stateful_widget(
                List::new(items).highlight_style(
                    Style::default()
                        .fg(Color::Black)
                        .bg(ACCENT)
                        .add_modifier(Modifier::BOLD),
                ),
                rows[1],
                &mut state,
            );
        }
        Overlay::Palette { query, sel } => {
            let area = centered(50, 60, f.area());
            f.render_widget(Clear, area);
            let matches = filter_commands(query);
            let block = popup_block("Command palette");
            let inner = block.inner(area);
            f.render_widget(block, area);
            let rows = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Length(2), Constraint::Min(0)])
                .split(inner);
            f.render_widget(Paragraph::new(format!("> {query}▏")), rows[0]);
            let items: Vec<ListItem> = matches
                .iter()
                .map(|(_, label)| ListItem::new(*label))
                .collect();
            let mut st = ListState::default();
            st.select(Some(*sel));
            f.render_stateful_widget(
                List::new(items).highlight_style(
                    Style::default()
                        .fg(Color::Black)
                        .bg(ACCENT)
                        .add_modifier(Modifier::BOLD),
                ),
                rows[1],
                &mut st,
            );
        }
        Overlay::Compose(c) => render_compose(f, c, app),
        Overlay::React => {
            let area = centered(50, 24, f.area());
            f.render_widget(Clear, area);
            let picks: String = crate::app::REACTIONS
                .iter()
                .enumerate()
                .map(|(i, e)| format!("{}  {e}   ", i + 1))
                .collect();
            f.render_widget(
                Paragraph::new(format!(
                    "React to the selected message:\n\n{picks}\n\nPress 1-7 · Esc cancel"
                ))
                .wrap(Wrap { trim: false })
                .block(popup_block("Add reaction")),
                area,
            );
        }
        Overlay::Attachments => {
            let area = centered(70, 50, f.area());
            f.render_widget(Clear, area);
            let block = popup_block("Attachments — press 1-9 to save · Esc close");
            let inner = block.inner(area);
            f.render_widget(block, area);
            let items: Vec<ListItem> = app
                .outlook
                .reading_attachments
                .iter()
                .take(9)
                .enumerate()
                .map(|(i, a)| {
                    ListItem::new(Line::from(vec![
                        Span::styled(
                            format!("{} ", i + 1),
                            Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
                        ),
                        Span::raw(a.display_name()),
                        Span::styled(
                            format!(
                                "  {}  {}",
                                a.human_size(),
                                a.content_type.clone().unwrap_or_default()
                            ),
                            Style::default().fg(DIM),
                        ),
                    ]))
                })
                .collect();
            f.render_widget(List::new(items), inner);
            let hint = format!("saves to {}", crate::files::download_dir().display());
            let hint_area = Rect {
                y: inner.y + inner.height.saturating_sub(1),
                height: 1,
                ..inner
            };
            f.render_widget(
                Paragraph::new(Span::styled(hint, Style::default().fg(DIM))),
                hint_area,
            );
        }
        Overlay::Links => {
            let links = app.focused_links();
            let area = centered(80, 60, f.area());
            f.render_widget(Clear, area);
            let block = popup_block("Links — press 1-9 to open · y copy first · Esc close");
            let inner = block.inner(area);
            f.render_widget(block, area);
            let width = inner.width.saturating_sub(4).max(10) as usize;
            let items: Vec<ListItem> = links
                .iter()
                .take(9)
                .enumerate()
                .map(|(i, url)| {
                    // Wrap long URLs across lines so the whole target is visible.
                    let mut lines = vec![Line::from(vec![
                        Span::styled(
                            format!("{} ", i + 1),
                            Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
                        ),
                        Span::styled(host_of(url), Style::default().fg(Color::LightGreen)),
                    ])];
                    for chunk in chunks_of(url, width) {
                        lines.push(Line::styled(format!("  {chunk}"), Style::default().fg(DIM)));
                    }
                    ListItem::new(lines)
                })
                .collect();
            f.render_widget(List::new(items), inner);
        }
        Overlay::MoveMail { sel } => {
            let area = centered(50, 70, f.area());
            f.render_widget(Clear, area);
            let items: Vec<ListItem> = app
                .outlook
                .folders
                .iter()
                .map(|folder| ListItem::new(mail_folder_label(folder)))
                .collect();
            let mut st = ListState::default();
            st.select(Some(*sel));
            f.render_stateful_widget(
                List::new(items)
                    .block(popup_block("Move to folder — Enter to move · Esc cancel"))
                    .highlight_style(
                        Style::default()
                            .fg(Color::Black)
                            .bg(ACCENT)
                            .add_modifier(Modifier::BOLD),
                    ),
                area,
                &mut st,
            );
        }
        Overlay::Presence => {
            let area = centered(46, 55, f.area());
            f.render_widget(Clear, area);
            let mut body = String::new();
            if let Some(a) = app
                .my_presence
                .as_ref()
                .and_then(|p| p.availability.as_deref())
            {
                body.push_str(&format!("Current: {a}\n\n"));
            }
            for (i, opt) in crate::app::PRESENCE_OPTIONS.iter().enumerate() {
                body.push_str(&format!("{}  {}\n", i + 1, opt.label));
            }
            body.push_str("\nc  Clear (revert to automatic)\nEsc cancel");
            body.push_str(
                "\n\nThis app publishes its own presence session, so the status\nshows even with no Teams client running. Quitting clears it.",
            );
            if !app.session.config.can_write_presence() {
                body.push_str("\n\nread-only: set M365_PRESENCE_WRITE=1 and grant\nPresence.ReadWrite to enable changing status");
            }
            f.render_widget(
                Paragraph::new(body).block(popup_block("Set presence")),
                area,
            );
        }
        Overlay::Settings { sel } => {
            let area = centered(58, 30, f.area());
            f.render_widget(Clear, area);
            let marker = if app.settings.preview_mail_on_hover {
                "[x]"
            } else {
                "[ ]"
            };
            let line = Line::styled(
                format!("{marker} Preview selected email automatically"),
                if *sel == 0 {
                    Style::default().fg(Color::Black).bg(ACCENT)
                } else {
                    Style::default()
                },
            );
            f.render_widget(
                Paragraph::new(vec![
                    line,
                    Line::raw(""),
                    Line::styled(
                        "Space/Enter toggle · settings persist across restarts",
                        Style::default().fg(DIM),
                    ),
                ])
                .block(popup_block("Settings")),
                area,
            );
        }
    }
}

/// Render a wrapped, vertically-scrolling text area. Returns the on-screen
/// cursor position when focused. Shared by the compose body and the Teams
/// composer so both wrap and scroll identically.
fn render_text_area(
    f: &mut Frame,
    area: Rect,
    input: &crate::editor::TextInput,
    focused: bool,
) -> Option<(u16, u16)> {
    let width = area.width.max(1) as usize;
    let height = area.height.max(1) as usize;
    let wrapped = input.wrap(width);
    let (crow, ccol) = input.cursor_position(width);
    // Keep the cursor row on screen.
    let scroll = crow.saturating_sub(height.saturating_sub(1));
    let visible: Vec<Line> = wrapped
        .iter()
        .skip(scroll)
        .take(height)
        .map(|r| Line::raw(r.text.clone()))
        .collect();
    f.render_widget(Paragraph::new(visible), area);
    focused.then(|| {
        (
            area.x + ccol.min(width.saturating_sub(1)) as u16,
            area.y + (crow - scroll) as u16,
        )
    })
}

/// Render one single-line field, scrolling horizontally to keep the cursor in
/// view. Returns the on-screen cursor column when this field is focused.
fn render_line_field(
    f: &mut Frame,
    area: Rect,
    label: &str,
    input: &crate::editor::TextInput,
    focused: bool,
) -> Option<(u16, u16)> {
    let style = if focused {
        Style::default().fg(ACCENT).add_modifier(Modifier::BOLD)
    } else {
        Style::default()
    };
    let label_w = label.chars().count() as u16;
    let avail = area.width.saturating_sub(label_w).max(1) as usize;
    let text = input.text();
    let chars: Vec<char> = text.chars().collect();
    // Scroll so the cursor stays visible in a long recipient list.
    let offset = input.cursor().saturating_sub(avail.saturating_sub(1));
    let shown: String = chars.iter().skip(offset).take(avail).collect();

    f.render_widget(Paragraph::new(format!("{label}{shown}")).style(style), area);

    focused.then(|| {
        let col = area.x + label_w + (input.cursor() - offset) as u16;
        (col.min(area.x + area.width.saturating_sub(1)), area.y)
    })
}

fn render_compose(f: &mut Frame, c: &Compose, app: &App) {
    let area = centered(70, 70, f.area());
    f.render_widget(Clear, area);
    let block = popup_block(c.kind.title());
    let inner = block.inner(area);
    f.render_widget(block, area);

    let fields = c.fields();
    let show_to = fields.contains(&COMPOSE_TO);
    let show_cc = fields.contains(&COMPOSE_CC);
    let show_bcc = fields.contains(&COMPOSE_BCC);
    let show_subject = fields.contains(&COMPOSE_SUBJECT);
    let suggestions = app.compose_suggestions(c);

    // Recipient/subject rows + autocomplete + body + attachments + hint.
    let staged = c.attachments.len() as u16;
    let mut constraints = Vec::new();
    if show_to {
        constraints.push(Constraint::Length(1));
    }
    if show_cc {
        constraints.push(Constraint::Length(1));
    }
    if show_bcc {
        constraints.push(Constraint::Length(1));
    }
    if !suggestions.is_empty() {
        constraints.push(Constraint::Length(suggestions.len() as u16));
    }
    if show_subject {
        constraints.push(Constraint::Length(1));
    }
    constraints.push(Constraint::Length(1)); // Body: label
    constraints.push(Constraint::Min(0)); // body
    constraints.push(Constraint::Length(1)); // Attach: input
    constraints.push(Constraint::Length(staged.min(4))); // staged files
    constraints.push(Constraint::Length(1)); // hint
    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints(constraints)
        .split(inner);

    let mut cursor: Option<(u16, u16)> = None;
    let mut i = 0;
    if show_to {
        cursor =
            render_line_field(f, rows[i], "To:      ", &c.to, c.field == COMPOSE_TO).or(cursor);
        i += 1;
    }
    if show_cc {
        cursor =
            render_line_field(f, rows[i], "Cc:      ", &c.cc, c.field == COMPOSE_CC).or(cursor);
        i += 1;
    }
    if show_bcc {
        cursor =
            render_line_field(f, rows[i], "Bcc:     ", &c.bcc, c.field == COMPOSE_BCC).or(cursor);
        i += 1;
    }
    if !suggestions.is_empty() {
        let selected = c.suggestion_sel.min(suggestions.len() - 1);
        let lines = suggestions.iter().enumerate().map(|(index, contact)| {
            let marker = if index == selected { "▶" } else { " " };
            let label = match contact.name.as_deref() {
                Some(name) => format!("  {marker} {name} <{}>", contact.address),
                None => format!("  {marker} {}", contact.address),
            };
            Line::styled(
                label,
                if index == selected {
                    Style::default().fg(Color::Black).bg(ACCENT)
                } else {
                    Style::default().fg(DIM)
                },
            )
        });
        f.render_widget(Paragraph::new(lines.collect::<Vec<_>>()), rows[i]);
        i += 1;
    }
    if show_subject {
        cursor = render_line_field(
            f,
            rows[i],
            "Subject: ",
            &c.subject,
            c.field == COMPOSE_SUBJECT,
        )
        .or(cursor);
        i += 1;
    }

    let body_style = if c.field == COMPOSE_BODY {
        Style::default().fg(ACCENT).add_modifier(Modifier::BOLD)
    } else {
        Style::default()
    };
    f.render_widget(Paragraph::new("Body:").style(body_style), rows[i]);
    i += 1;

    let body_area = rows[i];
    // Tell the key handler what width Up/Down should move by.
    app.text_width_hint.set(body_area.width.max(1) as usize);
    cursor = render_text_area(f, body_area, &c.body, c.field == COMPOSE_BODY).or(cursor);
    i += 1;

    // Attach: type a path, Enter stages it.
    cursor = render_line_field(
        f,
        rows[i],
        "Attach:  ",
        &c.attach,
        c.field == COMPOSE_ATTACH,
    )
    .or(cursor);
    i += 1;

    // Staged files (most recent last), capped to the rows we reserved.
    let staged_area = rows[i];
    if staged_area.height > 0 {
        let shown = staged_area.height as usize;
        let skip = c.attachments.len().saturating_sub(shown);
        let lines: Vec<Line> = c
            .attachments
            .iter()
            .skip(skip)
            .map(|(path, size)| {
                Line::from(vec![
                    Span::styled("  📎 ", Style::default().fg(Color::LightBlue)),
                    Span::raw(
                        path.file_name()
                            .map(|n| n.to_string_lossy().to_string())
                            .unwrap_or_default(),
                    ),
                    Span::styled(format!("  {}", human_size(*size)), Style::default().fg(DIM)),
                ])
            })
            .collect();
        f.render_widget(Paragraph::new(lines), staged_area);
    }
    i += 1;

    let hint = if c.ctrl_x {
        "Ctrl+X — e $EDITOR · x unstage last attachment"
    } else if c.field == COMPOSE_ATTACH {
        "Tab complete path · Enter attach · Ctrl+X x unstage · Ctrl+X e $EDITOR · Ctrl+S send"
    } else if !suggestions.is_empty() {
        "↑/↓ choose · Tab/Enter complete · keep typing to filter · Ctrl+S send"
    } else if matches!(
        c.kind,
        crate::app::ComposeKind::ReplyMail { .. } | crate::app::ComposeKind::ReplyAllMail { .. }
    ) {
        "Tab field/complete @image · Ctrl+R reply mode · Ctrl+E recipients · Ctrl+S send"
    } else {
        "Tab field/complete @image · Ctrl+X e $EDITOR · Ctrl+S send · Esc cancel"
    };
    f.render_widget(
        Paragraph::new(Span::styled(hint, Style::default().fg(DIM))),
        rows[i],
    );

    if let Some((x, y)) = cursor {
        f.set_cursor_position((x, y));
    }
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

fn human_size(bytes: u64) -> String {
    if bytes >= 1024 * 1024 {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    } else if bytes >= 1024 {
        format!("{:.0} KB", bytes as f64 / 1024.0)
    } else {
        format!("{bytes} B")
    }
}

/// The host part of a URL, for a readable link label.
fn host_of(url: &str) -> String {
    url.split("://")
        .nth(1)
        .and_then(|rest| rest.split('/').next())
        .unwrap_or(url)
        .to_string()
}

/// Split a long string into fixed-width chunks so it can be shown in full.
fn chunks_of(s: &str, width: usize) -> Vec<String> {
    let chars: Vec<char> = s.chars().collect();
    chars
        .chunks(width.max(1))
        .map(|c| c.iter().collect())
        .collect()
}

/// Parse a Graph UTC timestamp into the local timezone.
fn local_time(ts: Option<&str>) -> Option<chrono::DateTime<chrono::Local>> {
    chrono::DateTime::parse_from_rfc3339(ts?)
        .ok()
        .map(|t| t.with_timezone(&chrono::Local))
}

/// `Today` / `Yesterday` / `Mon 21 Jul` (with the year for other years).
fn day_label(day: chrono::NaiveDate) -> String {
    use chrono::Datelike;
    let today = chrono::Local::now().date_naive();
    if day == today {
        "Today".to_string()
    } else if Some(day) == today.pred_opt() {
        "Yesterday".to_string()
    } else if day.year() == today.year() {
        day.format("%a %-d %b").to_string()
    } else {
        day.format("%a %-d %b %Y").to_string()
    }
}

/// Day label for the message currently at the top of the visible area — the
/// content of the pinned header. `starts` is ascending, so the topmost visible
/// message is the last one starting at or above the scroll offset.
fn sticky_day_label(app: &App, starts: &[usize], scroll: u16) -> Option<String> {
    let idx = topmost_message_index(starts, scroll);
    let when = local_time(app.teams.messages.get(idx)?.created_date_time.as_deref())?;
    Some(day_label(when.date_naive()))
}

/// Index of the message occupying the top of the visible area.
fn topmost_message_index(starts: &[usize], scroll: u16) -> usize {
    starts
        .iter()
        .rposition(|&s| s <= scroll as usize)
        .unwrap_or(0)
}

fn day_separator(label: &str) -> Line<'static> {
    Line::from(vec![
        Span::styled("── ", Style::default().fg(DIM)),
        Span::styled(
            label.to_string(),
            Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
        ),
        Span::styled(" ".to_string() + &"─".repeat(40), Style::default().fg(DIM)),
    ])
}

/// Tab-bar presence dot symbol + availability label.
fn presence_indicator(app: &App) -> (&'static str, String) {
    let avail = app
        .my_presence
        .as_ref()
        .and_then(|p| p.availability.clone())
        .unwrap_or_else(|| "…".into());
    ("●", avail)
}

fn presence_style(app: &App) -> Style {
    let color = match app
        .my_presence
        .as_ref()
        .and_then(|p| p.availability.as_deref())
        .unwrap_or("")
    {
        "Available" | "AvailableIdle" => Color::Green,
        "Busy" | "BusyIdle" | "DoNotDisturb" => Color::Red,
        "Away" | "BeRightBack" => Color::Yellow,
        _ => DIM,
    };
    Style::default().fg(color)
}

fn selectable_list<'a>(items: Vec<ListItem<'a>>, title: &'a str, focused: bool) -> List<'a> {
    List::new(items)
        .block(panel_block(title, focused))
        .highlight_style(
            Style::default()
                .fg(Color::Black)
                .bg(if focused { ACCENT } else { Color::Gray })
                .add_modifier(Modifier::BOLD),
        )
        .highlight_symbol("▏")
}

fn panel_block(title: &str, focused: bool) -> Block<'_> {
    let color = if focused { ACCENT } else { DIM };
    Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(color))
        .title(Span::styled(
            format!(" {title} "),
            Style::default().fg(color).add_modifier(Modifier::BOLD),
        ))
}

fn popup_block(title: &str) -> Block<'_> {
    Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(ACCENT))
        .title(Span::styled(
            format!(" {title} "),
            Style::default().fg(ACCENT).add_modifier(Modifier::BOLD),
        ))
}

fn kv<'a>(k: &'a str, v: &str) -> Line<'a> {
    Line::from(vec![
        Span::styled(format!("{k}: "), Style::default().fg(DIM)),
        Span::raw(v.to_string()),
    ])
}

fn truncate(s: &str, max: usize) -> String {
    if s.chars().count() <= max {
        s.to_string()
    } else {
        let mut out: String = s.chars().take(max.saturating_sub(1)).collect();
        out.push('…');
        out
    }
}

fn centered(pct_x: u16, pct_y: u16, area: Rect) -> Rect {
    let v = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Percentage((100 - pct_y) / 2),
            Constraint::Percentage(pct_y),
            Constraint::Percentage((100 - pct_y) / 2),
        ])
        .split(area);
    Layout::default()
        .direction(Direction::Horizontal)
        .constraints([
            Constraint::Percentage((100 - pct_x) / 2),
            Constraint::Percentage(pct_x),
            Constraint::Percentage((100 - pct_x) / 2),
        ])
        .split(v[1])[1]
}

#[cfg(test)]
mod tests {
    use super::{day_label, local_time};

    #[test]
    fn labels_relative_days() {
        let today = chrono::Local::now().date_naive();
        assert_eq!(day_label(today), "Today");
        assert_eq!(day_label(today.pred_opt().unwrap()), "Yesterday");
        // An older date renders as a weekday/day/month, not "Today".
        let old = chrono::NaiveDate::from_ymd_opt(2024, 3, 5).unwrap();
        let label = day_label(old);
        assert!(label.contains("Mar"), "unexpected label: {label}");
        assert!(label.contains("2024"), "past years show the year: {label}");
        assert!(!label.contains('-'), "no literal padding modifier: {label}");
    }

    #[test]
    fn sticky_header_tracks_topmost_message() {
        use super::topmost_message_index;
        // Three messages beginning at lines 0, 5 and 12.
        let starts = [0usize, 5, 12];
        assert_eq!(topmost_message_index(&starts, 0), 0);
        assert_eq!(topmost_message_index(&starts, 4), 0); // still inside msg 0
        assert_eq!(topmost_message_index(&starts, 5), 1); // exactly at msg 1
        assert_eq!(topmost_message_index(&starts, 11), 1);
        assert_eq!(topmost_message_index(&starts, 12), 2);
        assert_eq!(topmost_message_index(&starts, 99), 2); // clamped past the end
                                                           // A separator above the first message must not select a negative index.
        assert_eq!(topmost_message_index(&[3, 9], 0), 0);
        assert_eq!(topmost_message_index(&[], 7), 0);
    }

    #[test]
    fn groups_consecutive_messages_from_one_sender() {
        use super::continues_run;
        let at = |h, m| {
            Some(
                chrono::NaiveDate::from_ymd_opt(2026, 8, 3)
                    .unwrap()
                    .and_hms_opt(h, m, 0)
                    .unwrap()
                    .and_local_timezone(chrono::Local)
                    .unwrap(),
            )
        };

        // Same person, a minute apart: one header covers both.
        assert!(continues_run("Jaime", at(16, 30), "Jaime", at(16, 29)));
        // Different people never group.
        assert!(!continues_run("Jaime", at(16, 30), "António", at(16, 29)));
        // A long pause earns a fresh header even for the same person.
        assert!(!continues_run("Jaime", at(16, 30), "Jaime", at(15, 00)));
        // Gap is symmetric — the list runs newest-first.
        assert!(!continues_run("Jaime", at(15, 00), "Jaime", at(16, 30)));
        // Missing timestamps fall back to the author check alone.
        assert!(continues_run("Jaime", None, "Jaime", at(16, 30)));
    }

    #[test]
    fn body_lines_align_under_the_timestamp_gutter() {
        use super::TIME_WIDTH;
        // Every message opens with `marker + HH:MM + space`; wrapped body lines
        // are indented by exactly that, so text stays in one column whether or
        // not the message is grouped.
        for marker in ["▶ ", "  ", ""] {
            let lead = marker.chars().count() + TIME_WIDTH + 1;
            let gutter = " ".repeat(lead);
            assert_eq!(gutter.chars().count(), lead, "marker {marker:?}");
        }
    }

    #[test]
    fn parses_graph_timestamps_to_local() {
        assert!(local_time(Some("2026-07-27T14:30:00Z")).is_some());
        assert!(local_time(Some("2026-07-27T14:30:00.123Z")).is_some());
        assert!(local_time(Some("not a date")).is_none());
        assert!(local_time(None).is_none());
    }
}
