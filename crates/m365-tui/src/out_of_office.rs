//! Out of Office editor helpers.
//!
//! Scheduled OOO values are edited as segmented local wall-clock fields. Each
//! segment is validated when focus leaves it; the complete local timestamp is
//! validated when leaving the minute field so DST gaps/overlaps are rejected.
//! Graph is sent UTC instants, avoiding Windows/IANA time-zone ambiguity.

use chrono::{
    DateTime, Duration, Local, LocalResult, NaiveDate, NaiveDateTime, NaiveTime, TimeZone, Utc,
};
use m365_core::models::{AutomaticRepliesSetting, DateTimeTimeZone, MailboxSettingsSnapshot};

use crate::content;
use crate::editor::TextInput;

pub const OUTLOOK_OOO_URL: &str =
    "https://outlook.office.com/mail/options/accounts-category/automaticReply";

pub const START_YEAR: usize = 0;
pub const START_MONTH: usize = 1;
pub const START_DAY: usize = 2;
pub const START_HOUR: usize = 3;
pub const START_MINUTE: usize = 4;
pub const END_YEAR: usize = 5;
pub const END_MONTH: usize = 6;
pub const END_DAY: usize = 7;
pub const END_HOUR: usize = 8;
pub const END_MINUTE: usize = 9;
pub const AUDIENCE_FIELD: usize = 10;
pub const MESSAGE_FIELD: usize = 11;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Audience {
    Internal,
    All,
}

impl Audience {
    pub fn label(self) -> &'static str {
        match self {
            Self::Internal => "Internal",
            Self::All => "All",
        }
    }

    pub fn toggle(&mut self) {
        *self = match self {
            Self::Internal => Self::All,
            Self::All => Self::Internal,
        };
    }
}

#[derive(Debug, Clone)]
pub struct View {
    pub settings: MailboxSettingsSnapshot,
    pub status: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub start_detail: Option<String>,
    pub end_detail: Option<String>,
    pub audience: Audience,
    pub message: String,
    pub local_zone: String,
    pub mailbox_zone: Option<String>,
    pub stored_zone: Option<String>,
    pub external_differs: bool,
    pub dst_transition: bool,
}

#[derive(Debug, Clone)]
pub struct Editor {
    pub start_year: TextInput,
    pub start_month: TextInput,
    pub start_day: TextInput,
    pub start_hour: TextInput,
    pub start_minute: TextInput,
    pub end_year: TextInput,
    pub end_month: TextInput,
    pub end_day: TextInput,
    pub end_hour: TextInput,
    pub end_minute: TextInput,
    pub audience: Audience,
    pub message: TextInput,
    pub field: usize,
    pub scroll: u16,
    pub validation_error: Option<String>,
    pub local_zone: String,
    pub mailbox_zone: Option<String>,
    pub stored_zone: Option<String>,
    replace_on_type: bool,
    original_plain: String,
    original_internal_html: Option<String>,
}

impl Editor {
    pub const FIELD_COUNT: usize = 12;

    pub fn from_settings(settings: &MailboxSettingsSnapshot) -> Self {
        let setting = &settings.automatic_replies_setting;
        let (default_start, default_end) = default_range();

        let start = setting
            .scheduled_start_date_time
            .as_ref()
            .and_then(graph_datetime_for_editor)
            .unwrap_or(default_start);
        let end = setting
            .scheduled_end_date_time
            .as_ref()
            .and_then(graph_datetime_for_editor)
            .unwrap_or(default_end);

        let original_internal_html = setting.internal_reply_message.clone();
        let original_plain = original_internal_html
            .as_deref()
            .map(html_to_plain)
            .unwrap_or_default();

        Self {
            start_year: input(start.format("%Y")),
            start_month: input(start.format("%m")),
            start_day: input(start.format("%d")),
            start_hour: input(start.format("%H")),
            start_minute: input(start.format("%M")),
            end_year: input(end.format("%Y")),
            end_month: input(end.format("%m")),
            end_day: input(end.format("%d")),
            end_hour: input(end.format("%H")),
            end_minute: input(end.format("%M")),
            audience: audience_from_setting(setting),
            message: TextInput::from(original_plain.as_str()),
            field: START_YEAR,
            scroll: 0,
            validation_error: None,
            local_zone: local_timezone_name(),
            mailbox_zone: settings.time_zone.clone(),
            stored_zone: setting
                .scheduled_start_date_time
                .as_ref()
                .and_then(|value| value.time_zone.clone()),
            replace_on_type: true,
            original_plain,
            original_internal_html,
        }
    }

    pub fn is_message_field(&self) -> bool {
        self.field == MESSAGE_FIELD
    }

    pub fn is_audience_field(&self) -> bool {
        self.field == AUDIENCE_FIELD
    }

    pub fn is_numeric_field(&self) -> bool {
        self.field <= END_MINUTE
    }

    pub fn active_mut(&mut self) -> Option<&mut TextInput> {
        self.input_mut(self.field)
    }

    pub fn mark_editing(&mut self) {
        self.replace_on_type = false;
        self.validation_error = None;
    }

    pub fn next_field(&mut self) {
        self.leave_field(1);
    }

    pub fn prev_field(&mut self) {
        self.leave_field(-1);
    }

    pub fn insert_digit(&mut self, ch: char) {
        if !ch.is_ascii_digit() || !self.is_numeric_field() {
            return;
        }

        let field = self.field;
        let max_len = numeric_width(field);
        if self.replace_on_type {
            if let Some(value) = self.input_mut(field) {
                value.clear();
            }
            self.replace_on_type = false;
        }

        let current_len = self
            .input(field)
            .map(|value| value.text().chars().count())
            .unwrap_or(0);

        if current_len >= max_len {
            self.validation_error = Some(format!(
                "{} accepts at most {max_len} digits",
                field_label(field)
            ));
            return;
        }

        if let Some(value) = self.input_mut(field) {
            value.insert(ch);
        }
        self.validation_error = None;

        let new_len = self
            .input(field)
            .map(|value| value.text().chars().count())
            .unwrap_or(0);

        // Full-width input advances automatically only when it validates.
        if new_len == max_len {
            self.leave_field(1);
        }
    }

    pub fn reject_non_digit(&mut self) {
        if self.is_numeric_field() {
            self.validation_error =
                Some(format!("{} accepts digits only", field_label(self.field)));
        }
    }

    pub fn paste(&mut self, text: &str) {
        if self.is_message_field() {
            self.message.insert_str(text);
            self.validation_error = None;
            return;
        }
        if !self.is_numeric_field() {
            return;
        }

        let digits: String = text.chars().filter(|ch| ch.is_ascii_digit()).collect();
        if digits.is_empty() {
            self.reject_non_digit();
            return;
        }

        let max_len = numeric_width(self.field);
        let digits: String = digits.chars().take(max_len).collect();
        let field = self.field;
        if let Some(value) = self.input_mut(field) {
            *value = TextInput::from(digits.as_str());
        }
        self.replace_on_type = false;
        self.validation_error = None;

        if digits.chars().count() == max_len {
            self.leave_field(1);
        }
    }

    pub fn page_up(&mut self) {
        self.scroll = self.scroll.saturating_sub(3);
    }

    pub fn page_down(&mut self) {
        self.scroll = self.scroll.saturating_add(3).min(16);
    }

    pub fn focus_row(&self) -> u16 {
        match self.field {
            START_YEAR | START_MONTH | START_DAY => 2,
            START_HOUR | START_MINUTE => 3,
            END_YEAR | END_MONTH | END_DAY => 4,
            END_HOUR | END_MINUTE => 5,
            AUDIENCE_FIELD => 6,
            MESSAGE_FIELD => 10,
            _ => 0,
        }
    }

    pub fn start_detail(&self) -> String {
        self.start_local()
            .map(|value| local_time_detail_value(&value))
            .unwrap_or_else(|error| format!("! {error}"))
    }

    pub fn end_detail(&self) -> String {
        self.end_local()
            .map(|value| local_time_detail_value(&value))
            .unwrap_or_else(|error| format!("! {error}"))
    }

    pub fn dst_transition(&self) -> bool {
        let Ok(start) = self.start_local() else {
            return false;
        };
        let Ok(end) = self.end_local() else {
            return false;
        };
        start.offset().local_minus_utc() != end.offset().local_minus_utc()
    }

    pub fn setting(&self, markdown: bool) -> Result<AutomaticRepliesSetting, String> {
        let start = self.start_local()?;
        let end = self.end_local()?;
        if end <= start {
            return Err("OOO end must be after start".into());
        }

        let message = self.message.text();
        if message.trim().is_empty() {
            return Err("OOO message cannot be empty".into());
        }

        let html = if message == self.original_plain {
            self.original_internal_html
                .clone()
                .filter(|value| !value.trim().is_empty())
                .unwrap_or_else(|| {
                    m365_core::teams_format::outgoing_body(&message, markdown).html_fragment()
                })
        } else {
            m365_core::teams_format::outgoing_body(&message, markdown).html_fragment()
        };

        let external = match self.audience {
            Audience::Internal => String::new(),
            Audience::All => html.clone(),
        };

        Ok(AutomaticRepliesSetting {
            status: "scheduled".into(),
            external_audience: Some(
                match self.audience {
                    Audience::Internal => "none",
                    Audience::All => "all",
                }
                .into(),
            ),
            scheduled_start_date_time: Some(DateTimeTimeZone {
                date_time: start
                    .with_timezone(&Utc)
                    .format("%Y-%m-%dT%H:%M:%S")
                    .to_string(),
                time_zone: Some("UTC".into()),
            }),
            scheduled_end_date_time: Some(DateTimeTimeZone {
                date_time: end
                    .with_timezone(&Utc)
                    .format("%Y-%m-%dT%H:%M:%S")
                    .to_string(),
                time_zone: Some("UTC".into()),
            }),
            internal_reply_message: Some(html),
            external_reply_message: Some(external),
        })
    }

    fn leave_field(&mut self, delta: isize) {
        let current = self.field;
        if let Err(error) = self.validate_field(current) {
            self.validation_error = Some(error);
            return;
        }

        self.validation_error = None;
        let count = Self::FIELD_COUNT as isize;
        self.field = ((self.field as isize + delta).rem_euclid(count)) as usize;
        self.replace_on_type = self.is_numeric_field();
    }

    fn validate_field(&mut self, field: usize) -> Result<(), String> {
        match field {
            START_YEAR | END_YEAR => self.validate_year(field)?,
            START_MONTH | END_MONTH => self.validate_two_digit(field, 1, 12)?,
            START_DAY => {
                self.validate_two_digit(field, 1, 31)?;
                self.validate_date(true)?;
            }
            END_DAY => {
                self.validate_two_digit(field, 1, 31)?;
                self.validate_date(false)?;
            }
            START_HOUR | END_HOUR => self.validate_two_digit(field, 0, 23)?,
            START_MINUTE => {
                self.validate_two_digit(field, 0, 59)?;
                self.start_local()?;
            }
            END_MINUTE => {
                self.validate_two_digit(field, 0, 59)?;
                self.end_local()?;
            }
            AUDIENCE_FIELD | MESSAGE_FIELD => {}
            _ => return Err("unknown OOO editor field".into()),
        }
        Ok(())
    }

    fn validate_year(&mut self, field: usize) -> Result<(), String> {
        let raw = self.input(field).map(TextInput::text).unwrap_or_default();
        if raw.chars().count() != 4 || !raw.chars().all(|ch| ch.is_ascii_digit()) {
            return Err(format!("{} must contain four digits", field_label(field)));
        }
        let year = raw
            .parse::<i32>()
            .map_err(|_| format!("{} is not a valid year", field_label(field)))?;
        if !(1900..=9999).contains(&year) {
            return Err(format!(
                "{} must be between 1900 and 9999",
                field_label(field)
            ));
        }
        Ok(())
    }

    fn validate_two_digit(&mut self, field: usize, min: u32, max: u32) -> Result<(), String> {
        let raw = self.input(field).map(TextInput::text).unwrap_or_default();
        if raw.is_empty() || raw.chars().count() > 2 || !raw.chars().all(|ch| ch.is_ascii_digit()) {
            return Err(format!(
                "{} must contain one or two digits",
                field_label(field)
            ));
        }
        let value = raw
            .parse::<u32>()
            .map_err(|_| format!("{} is not a number", field_label(field)))?;
        if !(min..=max).contains(&value) {
            return Err(format!(
                "{} must be in range {min}..{max}",
                field_label(field)
            ));
        }
        let normalized = format!("{value:02}");
        if let Some(input) = self.input_mut(field) {
            *input = TextInput::from(normalized.as_str());
        }
        Ok(())
    }

    fn validate_date(&self, start: bool) -> Result<(), String> {
        let (year, month, day) = if start {
            (
                self.number(START_YEAR)? as i32,
                self.number(START_MONTH)?,
                self.number(START_DAY)?,
            )
        } else {
            (
                self.number(END_YEAR)? as i32,
                self.number(END_MONTH)?,
                self.number(END_DAY)?,
            )
        };

        NaiveDate::from_ymd_opt(year, month, day)
            .map(|_| ())
            .ok_or_else(|| {
                format!(
                    "{} date {year:04}-{month:02}-{day:02} does not exist",
                    if start { "Start" } else { "End" }
                )
            })
    }

    fn start_local(&self) -> Result<DateTime<Local>, String> {
        parse_local_parts(
            self.number(START_YEAR)? as i32,
            self.number(START_MONTH)?,
            self.number(START_DAY)?,
            self.number(START_HOUR)?,
            self.number(START_MINUTE)?,
            "Start",
        )
    }

    fn end_local(&self) -> Result<DateTime<Local>, String> {
        parse_local_parts(
            self.number(END_YEAR)? as i32,
            self.number(END_MONTH)?,
            self.number(END_DAY)?,
            self.number(END_HOUR)?,
            self.number(END_MINUTE)?,
            "End",
        )
    }

    fn number(&self, field: usize) -> Result<u32, String> {
        self.input(field)
            .ok_or_else(|| "OOO editor field is unavailable".to_string())?
            .text()
            .parse::<u32>()
            .map_err(|_| format!("{} is invalid", field_label(field)))
    }

    fn input(&self, field: usize) -> Option<&TextInput> {
        match field {
            START_YEAR => Some(&self.start_year),
            START_MONTH => Some(&self.start_month),
            START_DAY => Some(&self.start_day),
            START_HOUR => Some(&self.start_hour),
            START_MINUTE => Some(&self.start_minute),
            END_YEAR => Some(&self.end_year),
            END_MONTH => Some(&self.end_month),
            END_DAY => Some(&self.end_day),
            END_HOUR => Some(&self.end_hour),
            END_MINUTE => Some(&self.end_minute),
            MESSAGE_FIELD => Some(&self.message),
            _ => None,
        }
    }

    fn input_mut(&mut self, field: usize) -> Option<&mut TextInput> {
        match field {
            START_YEAR => Some(&mut self.start_year),
            START_MONTH => Some(&mut self.start_month),
            START_DAY => Some(&mut self.start_day),
            START_HOUR => Some(&mut self.start_hour),
            START_MINUTE => Some(&mut self.start_minute),
            END_YEAR => Some(&mut self.end_year),
            END_MONTH => Some(&mut self.end_month),
            END_DAY => Some(&mut self.end_day),
            END_HOUR => Some(&mut self.end_hour),
            END_MINUTE => Some(&mut self.end_minute),
            MESSAGE_FIELD => Some(&mut self.message),
            _ => None,
        }
    }
}

impl View {
    pub fn from_settings(settings: MailboxSettingsSnapshot) -> Self {
        let setting = &settings.automatic_replies_setting;
        let local_zone = local_timezone_name();
        let start_local = setting
            .scheduled_start_date_time
            .as_ref()
            .and_then(graph_datetime_to_local);
        let end_local = setting
            .scheduled_end_date_time
            .as_ref()
            .and_then(graph_datetime_to_local);

        let (start, start_detail) = display_schedule_value(
            setting.scheduled_start_date_time.as_ref(),
            start_local.as_ref(),
        );
        let (end, end_detail) =
            display_schedule_value(setting.scheduled_end_date_time.as_ref(), end_local.as_ref());

        let dst_transition = match (start_local.as_ref(), end_local.as_ref()) {
            (Some(start), Some(end)) => {
                start.offset().local_minus_utc() != end.offset().local_minus_utc()
            }
            _ => false,
        };

        let message = setting
            .internal_reply_message
            .as_deref()
            .map(html_to_plain)
            .unwrap_or_default();

        let external_differs = setting
            .external_reply_message
            .as_deref()
            .filter(|value| !value.trim().is_empty())
            .is_some_and(|value| {
                value.trim()
                    != setting
                        .internal_reply_message
                        .as_deref()
                        .unwrap_or_default()
                        .trim()
            });

        Self {
            status: if setting.status.trim().is_empty() {
                "unknown".into()
            } else {
                setting.status.clone()
            },
            start,
            end,
            start_detail,
            end_detail,
            audience: audience_from_setting(setting),
            message,
            local_zone,
            mailbox_zone: settings.time_zone.clone(),
            stored_zone: setting
                .scheduled_start_date_time
                .as_ref()
                .and_then(|value| value.time_zone.clone()),
            external_differs,
            dst_transition,
            settings,
        }
    }

    pub fn text(&self) -> String {
        let mut out = String::new();
        out.push_str(&format!("Status: {}\n", self.status));
        out.push_str(&format!("Time zone: {} · DST-aware\n", self.local_zone));

        if let Some(mailbox_zone) = self.mailbox_zone.as_deref() {
            if mailbox_zone != self.local_zone {
                out.push_str(&format!("Mailbox zone: {mailbox_zone}\n"));
            }
        }

        if !self.status.eq_ignore_ascii_case("disabled") {
            if let Some(start) = self.start.as_deref() {
                out.push_str(&format!(
                    "Start: {start}{}\n",
                    self.start_detail
                        .as_deref()
                        .map(|detail| format!("  {detail}"))
                        .unwrap_or_default()
                ));
            }
            if let Some(end) = self.end.as_deref() {
                out.push_str(&format!(
                    "End:   {end}{}\n",
                    self.end_detail
                        .as_deref()
                        .map(|detail| format!("  {detail}"))
                        .unwrap_or_default()
                ));
            }
            if self.dst_transition {
                out.push_str("DST: transition occurs inside this interval\n");
            }
            if let Some(stored_zone) = self.stored_zone.as_deref() {
                if stored_zone != "UTC" && stored_zone != self.local_zone {
                    out.push_str(&format!(
                        "Stored schedule zone: {stored_zone} (displayed wall time may differ)\n"
                    ));
                }
            }
            out.push_str(&format!("Audience: {}\n\n", self.audience.label()));
            out.push_str("Message:\n");
            if self.message.trim().is_empty() {
                out.push_str("(empty)");
            } else {
                out.push_str(&self.message);
            }
            if self.external_differs && self.audience == Audience::All {
                out.push_str(
          "\n\n! Existing external reply differs; editing with All will unify both messages.",
        );
            }
        }

        out.push_str("\n\ns set/edit · c clear · w Outlook Web · Esc close");
        out
    }
}

pub fn local_timezone_name() -> String {
    if let Ok(value) = std::env::var("TZ") {
        let value = value.trim().trim_start_matches(':');
        if !value.is_empty() {
            if let Some((_, zone)) = value.split_once("/zoneinfo/") {
                return zone.to_string();
            }
            return value.to_string();
        }
    }

    if let Ok(target) = std::fs::read_link("/etc/localtime") {
        let value = target.to_string_lossy();
        if let Some((_, zone)) = value.split_once("/zoneinfo/") {
            if !zone.trim().is_empty() {
                return zone.to_string();
            }
        }
    }

    Local::now().offset().to_string()
}

fn input(value: impl std::fmt::Display) -> TextInput {
    TextInput::from(value.to_string().as_str())
}

fn numeric_width(field: usize) -> usize {
    match field {
        START_YEAR | END_YEAR => 4,
        START_MONTH | START_DAY | START_HOUR | START_MINUTE | END_MONTH | END_DAY | END_HOUR
        | END_MINUTE => 2,
        _ => 0,
    }
}

fn field_label(field: usize) -> &'static str {
    match field {
        START_YEAR => "Start year",
        START_MONTH => "Start month",
        START_DAY => "Start day",
        START_HOUR => "Start hour",
        START_MINUTE => "Start minute",
        END_YEAR => "End year",
        END_MONTH => "End month",
        END_DAY => "End day",
        END_HOUR => "End hour",
        END_MINUTE => "End minute",
        AUDIENCE_FIELD => "Audience",
        MESSAGE_FIELD => "Message",
        _ => "Field",
    }
}

fn default_range() -> (DateTime<Local>, DateTime<Local>) {
    let start = Local::now() + Duration::minutes(15);
    let end = start + Duration::days(1);
    (start, end)
}

fn parse_local_parts(
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    label: &str,
) -> Result<DateTime<Local>, String> {
    let date = NaiveDate::from_ymd_opt(year, month, day)
        .ok_or_else(|| format!("{label} date {year:04}-{month:02}-{day:02} does not exist"))?;
    let time = NaiveTime::from_hms_opt(hour, minute, 0)
        .ok_or_else(|| format!("{label} time {hour:02}:{minute:02} is invalid"))?;
    let naive = date.and_time(time);

    match Local.from_local_datetime(&naive) {
        LocalResult::Single(value) => Ok(value),
        LocalResult::None => Err(format!(
            "{label} local time does not exist because of a DST change"
        )),
        LocalResult::Ambiguous(_, _) => Err(format!(
            "{label} local time is ambiguous because of a DST change"
        )),
    }
}

fn local_time_detail_value(value: &DateTime<Local>) -> String {
    let abbreviation = value.format("%Z").to_string();
    let offset = value.format("%:z").to_string();
    if abbreviation.trim().is_empty() {
        format!("UTC{offset}")
    } else {
        format!("{abbreviation} · UTC{offset}")
    }
}

fn parse_graph_naive(value: &str) -> Option<NaiveDateTime> {
    let normalized = value.trim().replace(" T ", "T");
    for format in [
        "%Y-%m-%dT%H:%M:%S%.f",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M",
    ] {
        if let Ok(value) = NaiveDateTime::parse_from_str(&normalized, format) {
            return Some(value);
        }
    }
    None
}

fn graph_datetime_to_local(value: &DateTimeTimeZone) -> Option<DateTime<Local>> {
    let naive = parse_graph_naive(&value.date_time)?;
    let zone = value.time_zone.as_deref().unwrap_or("").trim();

    if zone.eq_ignore_ascii_case("UTC") || zone.eq_ignore_ascii_case("Etc/UTC") {
        return Some(Utc.from_utc_datetime(&naive).with_timezone(&Local));
    }

    if zone.is_empty() || zone == local_timezone_name() {
        return match Local.from_local_datetime(&naive) {
            LocalResult::Single(value) => Some(value),
            _ => None,
        };
    }

    None
}

fn graph_datetime_for_editor(value: &DateTimeTimeZone) -> Option<DateTime<Local>> {
    if let Some(value) = graph_datetime_to_local(value) {
        return Some(value);
    }

    let naive = parse_graph_naive(&value.date_time)?;
    match Local.from_local_datetime(&naive) {
        LocalResult::Single(value) => Some(value),
        _ => None,
    }
}

fn display_schedule_value(
    raw: Option<&DateTimeTimeZone>,
    local: Option<&DateTime<Local>>,
) -> (Option<String>, Option<String>) {
    if let Some(local) = local {
        return (
            Some(local.format("%Y-%m-%d %H:%M").to_string()),
            Some(local_time_detail_value(local)),
        );
    }

    let Some(raw) = raw else {
        return (None, None);
    };
    let Some(naive) = parse_graph_naive(&raw.date_time) else {
        return (
            Some(raw.date_time.clone()),
            raw.time_zone
                .as_deref()
                .map(|zone| format!("stored zone: {zone}")),
        );
    };
    (
        Some(naive.format("%Y-%m-%d %H:%M").to_string()),
        raw.time_zone
            .as_deref()
            .map(|zone| format!("stored zone: {zone}")),
    )
}

fn audience_from_setting(setting: &AutomaticRepliesSetting) -> Audience {
    if setting
        .external_audience
        .as_deref()
        .is_some_and(|value| value.eq_ignore_ascii_case("all"))
    {
        Audience::All
    } else {
        Audience::Internal
    }
}

fn html_to_plain(value: &str) -> String {
    let rendered = content::render_body(Some("html"), value);
    content::plain(&rendered.text).trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn settings() -> MailboxSettingsSnapshot {
        MailboxSettingsSnapshot {
            time_zone: Some("test-zone".into()),
            automatic_replies_setting: AutomaticRepliesSetting::default(),
        }
    }

    #[test]
    fn invalid_month_is_rejected_on_field_leave() {
        let mut editor = Editor::from_settings(&settings());
        editor.field = START_MONTH;
        editor.start_month = TextInput::from("13");
        editor.next_field();
        assert_eq!(editor.field, START_MONTH);
        assert!(editor.validation_error.is_some());
    }

    #[test]
    fn valid_single_digit_month_is_normalized_on_leave() {
        let mut editor = Editor::from_settings(&settings());
        editor.field = START_MONTH;
        editor.start_month = TextInput::from("7");
        editor.next_field();
        assert_eq!(editor.field, START_DAY);
        assert_eq!(editor.start_month.text(), "07");
        assert!(editor.validation_error.is_none());
    }

    #[test]
    fn impossible_date_is_rejected_when_leaving_day() {
        let mut editor = Editor::from_settings(&settings());
        editor.start_year = TextInput::from("2026");
        editor.start_month = TextInput::from("02");
        editor.start_day = TextInput::from("30");
        editor.field = START_DAY;
        editor.next_field();
        assert_eq!(editor.field, START_DAY);
        assert!(editor
            .validation_error
            .as_deref()
            .is_some_and(|error| error.contains("does not exist")));
    }

    #[test]
    fn audience_maps_only_internal_or_all() {
        let mut setting = AutomaticRepliesSetting::default();
        assert_eq!(audience_from_setting(&setting), Audience::Internal);
        setting.external_audience = Some("all".into());
        assert_eq!(audience_from_setting(&setting), Audience::All);
    }

    #[test]
    fn graph_datetime_parser_accepts_fractional_seconds() {
        let value = parse_graph_naive("2026-10-21T12:34:56.0000000").unwrap();
        assert_eq!(
            value.format("%Y-%m-%d %H:%M").to_string(),
            "2026-10-21 12:34"
        );
    }
}
