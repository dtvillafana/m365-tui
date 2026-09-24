//! Calendar agenda and month rendering, ported from the Pitriss m365-tui fork.

use ratatui::layout::Rect;
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, List, ListItem, ListState, Paragraph, Wrap};
use ratatui::Frame;

use crate::app::{App, CalendarView};

const ACCENT: Color = Color::Cyan;
const DIM: Color = Color::DarkGray;

fn truncate(s: &str, max: usize) -> String {
    let count = s.chars().count();
    if count <= max {
        return s.to_string();
    }
    if max <= 1 {
        return "…".to_string();
    }
    format!("{}…", s.chars().take(max.saturating_sub(1)).collect::<String>())
}

// Calendar
// ---------------------------------------------------------------------------

fn calendar_local_datetime(
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

fn calendar_response_marker(event: &m365_core::models::Event) -> (&'static str, Color) {
    if event.is_cancelled.unwrap_or(false) {
        return ("!", Color::Red);
    }
    if event.is_organizer.unwrap_or(false) {
        return ("O", ACCENT);
    }

    let response = event
        .response_status
        .as_ref()
        .and_then(|status| status.response.as_deref())
        .unwrap_or("")
        .to_ascii_lowercase();

    match response.as_str() {
        "accepted" => ("A", Color::Green),
        "tentativelyaccepted" => ("T", Color::Yellow),
        "declined" => ("D", Color::Red),
        "notresponded" | "none" => ("?", Color::Yellow),
        "organizer" => ("O", ACCENT),
        _ => (" ", DIM),
    }
}

fn calendar_response_label(event: &m365_core::models::Event) -> &'static str {
    if event.is_cancelled.unwrap_or(false) {
        return "cancelled";
    }
    if event.is_organizer.unwrap_or(false) {
        return "organizer";
    }

    let response = event
        .response_status
        .as_ref()
        .and_then(|status| status.response.as_deref())
        .unwrap_or("");

    if response.eq_ignore_ascii_case("accepted") {
        "accepted"
    } else if response.eq_ignore_ascii_case("tentativelyAccepted") {
        "tentative"
    } else if response.eq_ignore_ascii_case("declined") {
        "declined"
    } else if response.eq_ignore_ascii_case("notResponded") || response.eq_ignore_ascii_case("none")
    {
        "waiting"
    } else if response.eq_ignore_ascii_case("organizer") {
        "organizer"
    } else {
        "unknown"
    }
}

// Join availability is independent of RSVP and the isOnlineMeeting flag.
fn calendar_has_join_url(event: &m365_core::models::Event) -> bool {
    event
        .online_meeting
        .as_ref()
        .and_then(|meeting| meeting.join_url.as_deref())
        .is_some_and(|url| !url.trim().is_empty())
}

fn calendar_needs_response(event: &m365_core::models::Event) -> bool {
    !event.is_cancelled.unwrap_or(false)
        && !event.is_organizer.unwrap_or(false)
        && event
            .response_status
            .as_ref()
            .and_then(|status| status.response.as_deref())
            .is_some_and(|response| {
                response.eq_ignore_ascii_case("notResponded")
                    || response.eq_ignore_ascii_case("none")
            })
}

fn calendar_time_label(event: &m365_core::models::Event) -> String {
    if event.is_all_day.unwrap_or(false) {
        return "all day    ".to_string();
    }

    let start = event
        .start
        .as_ref()
        .and_then(calendar_local_datetime)
        .map(|value| value.format("%H:%M").to_string())
        .unwrap_or_else(|| "--:--".into());
    let end = event
        .end
        .as_ref()
        .and_then(calendar_local_datetime)
        .map(|value| value.format("%H:%M").to_string())
        .unwrap_or_else(|| "--:--".into());
    format!("{start}-{end}")
}

fn calendar_day_label(event: &m365_core::models::Event) -> String {
    event
        .start
        .as_ref()
        .and_then(calendar_local_datetime)
        .map(|value| value.format("%d.%m.").to_string())
        .unwrap_or_else(|| "--.--.".into())
}

#[cfg_attr(not(test), allow(dead_code))]
fn calendar_plain_line(event: &m365_core::models::Event, show_day: bool) -> String {
    let day = if show_day {
        calendar_day_label(event)
    } else {
        "      ".into()
    };
    let time = calendar_time_label(event);
    let marker = calendar_response_marker(event).0;
    let meeting = if calendar_has_join_url(event) {
        "M"
    } else {
        " "
    };
    let subject = event.subject.as_deref().unwrap_or("(no subject)");
    format!("{day}  {time:<11}  [{marker}] [{meeting}] {subject}")
}

fn calendar_month_first_ui(offset: i32) -> chrono::NaiveDate {
    let today = chrono::Local::now().date_naive();
    let month_index =
        chrono::Datelike::year(&today) * 12 + chrono::Datelike::month0(&today) as i32 + offset;
    let year = month_index.div_euclid(12);
    let month = month_index.rem_euclid(12) as u32 + 1;
    chrono::NaiveDate::from_ymd_opt(year, month, 1).expect("valid calendar month")
}

fn calendar_month_columns(width: u16) -> usize {
    match width {
        272.. => 4,
        204..=271 => 3,
        136..=203 => 2,
        _ => 1,
    }
}

fn calendar_month_day_widths(inner_width: usize) -> [usize; 7] {
    let usable = inner_width.saturating_sub(6);
    let base = usable / 7;
    let remainder = usable % 7;
    let mut widths = [base; 7];
    for width in widths.iter_mut().take(remainder) {
        *width += 1;
    }
    widths
}

fn calendar_event_date_range(
    event: &m365_core::models::Event,
) -> Option<(
    chrono::NaiveDate,
    chrono::NaiveDate,
    chrono::DateTime<chrono::Local>,
)> {
    let start = event.start.as_ref().and_then(calendar_local_datetime)?;
    let end = event.end.as_ref().and_then(calendar_local_datetime)?;
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

    Some((start_date, end_date, start))
}

fn calendar_month_event_style(event: &m365_core::models::Event, selected: bool) -> Style {
    let color = if event.is_cancelled.unwrap_or(false) {
        Color::DarkGray
    } else if event.is_organizer.unwrap_or(false) {
        ACCENT
    } else {
        let response = event
            .response_status
            .as_ref()
            .and_then(|status| status.response.as_deref())
            .unwrap_or("");

        if response.eq_ignore_ascii_case("accepted") {
            Color::Green
        } else if response.eq_ignore_ascii_case("tentativelyAccepted") {
            Color::Yellow
        } else if response.eq_ignore_ascii_case("declined") {
            Color::Red
        } else if response.eq_ignore_ascii_case("notResponded")
            || response.eq_ignore_ascii_case("none")
        {
            Color::Yellow
        } else {
            Color::Blue
        }
    };

    let foreground = if matches!(color, Color::Green | Color::Yellow | Color::Cyan) {
        Color::Black
    } else {
        Color::White
    };

    let mut style = Style::default().fg(foreground).bg(color);
    if selected {
        style = style.add_modifier(Modifier::BOLD | Modifier::UNDERLINED | Modifier::REVERSED);
    } else if calendar_needs_response(event) {
        style = style.add_modifier(Modifier::BOLD);
    }
    style
}

fn calendar_month_separator(widths: &[usize; 7]) -> Line<'static> {
    let mut spans = Vec::new();
    for (day, width) in widths.iter().enumerate() {
        if day > 0 {
            spans.push(Span::styled("┼", Style::default().fg(DIM)));
        }
        spans.push(Span::styled("─".repeat(*width), Style::default().fg(DIM)));
    }
    Line::from(spans)
}

fn calendar_month_center(value: &str, width: usize) -> String {
    let value = truncate(value, width);
    let used = value.chars().count();
    let padding = width.saturating_sub(used);
    let left = padding / 2;
    let right = padding - left;
    format!("{}{}{}", " ".repeat(left), value, " ".repeat(right))
}

fn render_calendar_month_panel(
    f: &mut Frame,
    area: Rect,
    app: &App,
    month_offset: i32,
    active: bool,
) {
    let first = calendar_month_first_ui(month_offset);
    let leading = chrono::Datelike::weekday(&first).num_days_from_monday() as i64;
    let grid_start = first - chrono::Duration::days(leading);
    let today = chrono::Local::now().date_naive();

    let inner_width = area.width.saturating_sub(2) as usize;
    let inner_height = area.height.saturating_sub(2) as usize;
    let widths = calendar_month_day_widths(inner_width);

    if widths.iter().any(|width| *width < 5) || inner_height < 19 {
        f.render_widget(
            Paragraph::new("Window too small for month view.").block(
                Block::default()
                    .borders(Borders::ALL)
                    .title(first.format("%B %Y").to_string()),
            ),
            area,
        );
        return;
    }

    let event_rows = ((inner_height.saturating_sub(13)) / 6).clamp(1, 4);
    let weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    let mut lines: Vec<Line<'static>> = Vec::new();

    let mut header = Vec::new();
    for (day, width) in widths.iter().enumerate() {
        if day > 0 {
            header.push(Span::styled("│", Style::default().fg(DIM)));
        }
        header.push(Span::styled(
            calendar_month_center(weekdays[day], *width),
            Style::default()
                .fg(Color::Gray)
                .add_modifier(Modifier::BOLD),
        ));
    }
    lines.push(Line::from(header));
    lines.push(calendar_month_separator(&widths));

    for week in 0..6usize {
        let week_start = grid_start + chrono::Duration::days((week * 7) as i64);
        let week_end = week_start + chrono::Duration::days(6);

        let mut candidates: Vec<(usize, chrono::NaiveDate, chrono::NaiveDate)> = app
            .calendar
            .events
            .iter()
            .enumerate()
            .filter_map(|(index, event)| {
                let (start, end, _) = calendar_event_date_range(event)?;
                (start <= week_end && end >= week_start).then_some((index, start, end))
            })
            .collect();
        candidates.sort_by_key(|(_, start, end)| {
            (
                *start,
                std::cmp::Reverse(end.signed_duration_since(*start).num_days()),
            )
        });

        let mut lane_masks = vec![0u8; event_rows];
        let mut segments: Vec<(usize, usize, usize, usize)> = Vec::new();
        let mut overflow = [0usize; 7];

        for (event_index, event_start, event_end) in candidates {
            let visible_start = if event_start > week_start {
                event_start
            } else {
                week_start
            };
            let visible_end = if event_end < week_end {
                event_end
            } else {
                week_end
            };
            let start_day = visible_start.signed_duration_since(week_start).num_days() as usize;
            let end_day = visible_end.signed_duration_since(week_start).num_days() as usize;

            let mut mask = 0u8;
            for day in start_day..=end_day {
                mask |= 1u8 << day;
            }

            if let Some((lane, occupied)) = lane_masks
                .iter_mut()
                .enumerate()
                .find(|(_, occupied)| (**occupied & mask) == 0)
            {
                *occupied |= mask;
                segments.push((event_index, lane, start_day, end_day));
            } else {
                for count in overflow.iter_mut().take(end_day + 1).skip(start_day) {
                    *count += 1;
                }
            }
        }

        let mut dates = Vec::new();
        for (day, width) in widths.iter().enumerate() {
            if day > 0 {
                dates.push(Span::styled("│", Style::default().fg(DIM)));
            }
            let date = week_start + chrono::Duration::days(day as i64);
            let in_month = chrono::Datelike::month(&date) == chrono::Datelike::month(&first);
            let label = if overflow[day] > 0 {
                format!("{} +{}", chrono::Datelike::day(&date), overflow[day])
            } else {
                chrono::Datelike::day(&date).to_string()
            };
            let label = truncate(&label, *width);
            let text = format!("{:<width$}", label, width = *width);

            let style = if date == today {
                Style::default()
                    .fg(Color::Black)
                    .bg(ACCENT)
                    .add_modifier(Modifier::BOLD)
            } else if in_month {
                Style::default()
                    .fg(Color::Gray)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(DIM)
            };
            dates.push(Span::styled(text, style));
        }
        lines.push(Line::from(dates));

        for lane in 0..event_rows {
            let mut row = Vec::new();

            for (day, width) in widths.iter().enumerate() {
                if day > 0 {
                    let bridge = segments.iter().find(|(_, segment_lane, start, end)| {
                        *segment_lane == lane && day > *start && day <= *end
                    });
                    if let Some((event_index, _, _, _)) = bridge {
                        let event = &app.calendar.events[*event_index];
                        row.push(Span::styled(
                            " ",
                            calendar_month_event_style(
                                event,
                                active && *event_index == app.calendar.selected,
                            ),
                        ));
                    } else {
                        row.push(Span::styled("│", Style::default().fg(DIM)));
                    }
                }

                let segment = segments.iter().find(|(_, segment_lane, start, end)| {
                    *segment_lane == lane && day >= *start && day <= *end
                });

                if let Some((event_index, _, start_day, end_day)) = segment {
                    let event = &app.calendar.events[*event_index];
                    let (event_start, event_end, start_time) =
                        calendar_event_date_range(event).expect("validated event range");
                    let date = week_start + chrono::Duration::days(day as i64);
                    let first_visible_day = day == *start_day;
                    let last_visible_day = day == *end_day;

                    let mut label = if first_visible_day {
                        let mut prefix = String::new();
                        if event_start < week_start {
                            prefix.push_str("◀ ");
                        } else if !event.is_all_day.unwrap_or(false) && event_start == date {
                            prefix.push_str(&start_time.format("%H:%M ").to_string());
                        }
                        if calendar_has_join_url(event) {
                            prefix.push_str("M ");
                        }
                        if active && *event_index == app.calendar.selected {
                            prefix.push_str("▶ ");
                        }
                        prefix + event.subject.as_deref().unwrap_or("(no subject)")
                    } else if active && *event_index == app.calendar.selected {
                        "═".repeat(*width)
                    } else {
                        "━".repeat(*width)
                    };

                    if last_visible_day && event_end > week_end && *width >= 2 {
                        label.push_str(" ▶");
                    }

                    let label = truncate(&label, *width);
                    row.push(Span::styled(
                        format!("{:<width$}", label, width = *width),
                        calendar_month_event_style(
                            event,
                            active && *event_index == app.calendar.selected,
                        ),
                    ));
                } else {
                    row.push(Span::raw(" ".repeat(*width)));
                }
            }

            lines.push(Line::from(row));
        }

        if week < 5 {
            lines.push(calendar_month_separator(&widths));
        }
    }

    let title = if active {
        format!("{} · active", first.format("%B %Y"))
    } else {
        first.format("%B %Y").to_string()
    };

    f.render_widget(
        Paragraph::new(lines)
            .block(Block::default().borders(Borders::ALL).title(title))
            .wrap(Wrap { trim: false }),
        area,
    );
}

fn render_calendar_month(f: &mut Frame, area: Rect, app: &App) {
    let columns = calendar_month_columns(area.width);
    let gap = 1u16;
    let total_gap = gap * columns.saturating_sub(1) as u16;
    let usable = area.width.saturating_sub(total_gap);
    let base_width = usable / columns as u16;
    let remainder = usable % columns as u16;

    let mut x = area.x;
    for column in 0..columns {
        let width = base_width + if column < remainder as usize { 1 } else { 0 };
        let panel = Rect {
            x,
            y: area.y,
            width,
            height: area.height,
        };
        render_calendar_month_panel(
            f,
            panel,
            app,
            app.calendar.month_offset + column as i32,
            column == 0,
        );
        x = x.saturating_add(width).saturating_add(gap);
    }
}

pub fn render_calendar(f: &mut Frame, area: Rect, app: &App) {
    if app.calendar.view == CalendarView::Month {
        render_calendar_month(f, area, app);
        return;
    }

    let mut previous_day = String::new();
    let items: Vec<ListItem> = app
        .calendar
        .events
        .iter()
        .map(|event| {
            let day = calendar_day_label(event);
            let show_day = day != previous_day;
            previous_day = day;
            let (marker, marker_color) = calendar_response_marker(event);
            let day = if show_day {
                calendar_day_label(event)
            } else {
                "      ".into()
            };
            let meeting = if calendar_has_join_url(event) {
                "M"
            } else {
                " "
            };
            ListItem::new(Line::from(vec![
                Span::styled(
                    format!("{day}  "),
                    Style::default()
                        .fg(Color::Gray)
                        .add_modifier(Modifier::BOLD),
                ),
                Span::raw(format!("{:<11}  ", calendar_time_label(event))),
                Span::styled(
                    format!("[{marker}] "),
                    Style::default()
                        .fg(marker_color)
                        .add_modifier(Modifier::BOLD),
                ),
                Span::styled(
                    format!("[{meeting}] "),
                    if meeting == "M" {
                        Style::default().fg(ACCENT).add_modifier(Modifier::BOLD)
                    } else {
                        Style::default().fg(DIM)
                    },
                ),
                Span::styled(
                    event
                        .subject
                        .as_deref()
                        .unwrap_or("(no subject)")
                        .to_string(),
                    if calendar_needs_response(event) {
                        Style::default()
                            .fg(Color::Yellow)
                            .add_modifier(Modifier::BOLD)
                    } else {
                        Style::default()
                    },
                ),
            ]))
        })
        .collect();

    let items = if items.is_empty() {
        vec![ListItem::new(
            "No events in the next 7 days (or still loading).",
        )]
    } else {
        items
    };

    let mut state = ListState::default();
    if !app.calendar.events.is_empty() {
        state.select(Some(
            app.calendar
                .selected
                .min(app.calendar.events.len().saturating_sub(1)),
        ));
    }

    f.render_stateful_widget(
        List::new(items)
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .title(format!("Calendar — next {} days", app.calendar.days)),
            )
            .highlight_style(
                Style::default()
                    .fg(Color::Black)
                    .bg(ACCENT)
                    .add_modifier(Modifier::BOLD),
            ),
        area,
        &mut state,
    );
}
pub fn calendar_event_detail(event: &m365_core::models::Event) -> Vec<Line<'static>> {
    let subject = event
        .subject
        .as_deref()
        .unwrap_or("(no subject)")
        .to_string();
    let date = event
        .start
        .as_ref()
        .and_then(calendar_local_datetime)
        .map(|value| value.format("%A %d.%m.%Y").to_string())
        .unwrap_or_else(|| "unknown".into());
    let time = if event.is_all_day.unwrap_or(false) {
        "all day".to_string()
    } else {
        calendar_time_label(event)
    };

    let organizer = event
        .organizer
        .as_ref()
        .and_then(|recipient| recipient.email_address.as_ref())
        .map(|address| {
            address
                .name
                .clone()
                .or_else(|| address.address.clone())
                .unwrap_or_else(|| "unknown".into())
        })
        .unwrap_or_else(|| "unknown".into());

    let location = event
        .location
        .as_ref()
        .and_then(|location| location.display_name.clone())
        .filter(|value| !value.trim().is_empty())
        .unwrap_or_else(|| "-".into());

    let (_, response_color) = calendar_response_marker(event);
    let response = calendar_response_label(event);
    let online = event
        .online_meeting
        .as_ref()
        .and_then(|meeting| meeting.join_url.clone())
        .filter(|value| !value.trim().is_empty())
        .unwrap_or_else(|| {
            if event.is_online_meeting.unwrap_or(false) {
                "online meeting".into()
            } else {
                "-".into()
            }
        });

    let preview = event
        .body_preview
        .as_deref()
        .unwrap_or("")
        .trim()
        .to_string();

    let mut lines = vec![
        Line::from(Span::styled(
            subject,
            Style::default().add_modifier(Modifier::BOLD),
        )),
        Line::from(""),
        Line::from(vec![
            Span::styled("Date:      ", Style::default().fg(Color::Gray)),
            Span::raw(date),
        ]),
        Line::from(vec![
            Span::styled("Time:      ", Style::default().fg(Color::Gray)),
            Span::raw(time),
        ]),
        Line::from(vec![
            Span::styled("Status:    ", Style::default().fg(Color::Gray)),
            Span::styled(
                response.to_string(),
                Style::default()
                    .fg(response_color)
                    .add_modifier(Modifier::BOLD),
            ),
        ]),
        Line::from(vec![
            Span::styled("Organizer: ", Style::default().fg(Color::Gray)),
            Span::raw(organizer),
        ]),
        Line::from(vec![
            Span::styled("Location:  ", Style::default().fg(Color::Gray)),
            Span::raw(location),
        ]),
        Line::from(vec![
            Span::styled("Meeting:   ", Style::default().fg(Color::Gray)),
            Span::raw(online),
        ]),
    ];

    if !preview.is_empty() {
        lines.push(Line::from(""));
        lines.push(Line::from(Span::styled(
            "Preview",
            Style::default()
                .fg(Color::Gray)
                .add_modifier(Modifier::BOLD),
        )));
        lines.push(Line::from(preview));
    }

    lines
}

#[cfg(test)]
mod calendar_meeting_indicator_tests {
    use super::{calendar_has_join_url, calendar_plain_line, calendar_response_marker};
    use m365_core::models::Event;
    use serde_json::json;

    #[test]
    fn rsvp_and_join_indicators_are_independent() {
        for (response, marker) in [
            ("accepted", "A"),
            ("tentativelyAccepted", "T"),
            ("declined", "D"),
            ("notResponded", "?"),
            ("organizer", "O"),
        ] {
            for online in [false, true] {
                for (url, has_join) in [
                    (None, false),
                    (Some(""), false),
                    (Some("   "), false),
                    (Some("https://example.com/meeting"), true),
                ] {
                    let event: Event = serde_json::from_value(json!({
                        "id": "test",
                        "subject": "Example",
                        "responseStatus": {"response": response},
                        "isOnlineMeeting": online,
                        "onlineMeeting": {"joinUrl": url}
                    }))
                    .unwrap();
                    assert_eq!(calendar_response_marker(&event).0, marker);
                    assert_eq!(calendar_has_join_url(&event), has_join);
                    let join = if has_join { "M" } else { " " };
                    let line = calendar_plain_line(&event, true);
                    assert!(line.ends_with(&format!("[{marker}] [{join}] Example")));
                    assert_eq!(line.chars().count(), 36);
                }
            }
        }
    }

    #[test]
    fn organizer_and_cancelled_flags_do_not_hide_join_links() {
        for (organizer, cancelled, marker) in
            [(true, false, "O"), (false, true, "!"), (true, true, "!")]
        {
            let event: Event = serde_json::from_value(json!({
                "id": "test",
                "isOrganizer": organizer,
                "isCancelled": cancelled,
                "onlineMeeting": {"joinUrl": "https://example.com/meeting"}
            }))
            .unwrap();
            assert_eq!(calendar_response_marker(&event).0, marker);
            assert!(calendar_has_join_url(&event));
        }
    }
}
