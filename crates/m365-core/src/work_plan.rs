//! Microsoft 365 work-plan helpers used by ntfy work-day forwarding and diagnostics.

use anyhow::Result;
use chrono::{DateTime, SecondsFormat, Utc};
use serde::Deserialize;

use crate::graph::GraphClient;

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorkPlanOccurrence {
    #[serde(default)]
    work_location_type: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorkPlanDateTime {
    date_time: String,
    #[serde(default)]
    time_zone: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct RecurrencePattern {
    #[serde(default)]
    days_of_week: Vec<String>,
}

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct RecurrenceRange {
    #[serde(default)]
    recurrence_time_zone: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
struct PatternedRecurrence {
    #[serde(default)]
    pattern: RecurrencePattern,
    #[serde(default)]
    range: RecurrenceRange,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorkPlanRecurrence {
    #[serde(default)]
    work_location_type: Option<String>,
    #[serde(default)]
    start: Option<WorkPlanDateTime>,
    #[serde(default)]
    end: Option<WorkPlanDateTime>,
    #[serde(default)]
    recurrence: PatternedRecurrence,
}

/// Compact, non-sensitive representation used by the F6 diagnostics view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkPlanRecurrenceSummary {
    pub days: Vec<String>,
    pub start_time: String,
    pub end_time: String,
    pub time_zone: Option<String>,
    pub location: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkPlanDiagnostics {
    pub working_now: bool,
    pub recurrences: Vec<WorkPlanRecurrenceSummary>,
}

fn occurrence_kind_is_working(kind: Option<&str>) -> bool {
    kind.is_some_and(|kind| {
        kind.eq_ignore_ascii_case("office")
            || kind.eq_ignore_ascii_case("remote")
            || kind.eq_ignore_ascii_case("unspecified")
    })
}

fn occurrences_mean_working_now(occurrences: &[WorkPlanOccurrence]) -> bool {
    if occurrences.iter().any(|occurrence| {
        occurrence
            .work_location_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("timeOff"))
    }) {
        return false;
    }

    occurrences
        .iter()
        .any(|occurrence| occurrence_kind_is_working(occurrence.work_location_type.as_deref()))
}

fn occurrences_view_path(now: DateTime<Utc>) -> String {
    let end = now + chrono::Duration::seconds(1);
    let start = now.to_rfc3339_opts(SecondsFormat::Secs, true);
    let end = end.to_rfc3339_opts(SecondsFormat::Secs, true);

    format!(
        "me/settings/workHoursAndLocations/occurrencesView(startDateTime='{start}',endDateTime='{end}')?$select=workLocationType&$top=20"
    )
}

fn short_time(value: &str) -> String {
    let normalized = value.trim().replace(" T ", "T");
    let time = normalized
        .split_once('T')
        .map(|(_, time)| time.trim())
        .unwrap_or(normalized.as_str());
    time.chars().take(5).collect()
}

fn day_label(value: &str) -> String {
    match value.to_ascii_lowercase().as_str() {
        "monday" => "Mon",
        "tuesday" => "Tue",
        "wednesday" => "Wed",
        "thursday" => "Thu",
        "friday" => "Fri",
        "saturday" => "Sat",
        "sunday" => "Sun",
        _ => value,
    }
    .to_string()
}

fn summarize_recurrence(value: WorkPlanRecurrence) -> WorkPlanRecurrenceSummary {
    let start_time = value
        .start
        .as_ref()
        .map(|value| short_time(&value.date_time))
        .unwrap_or_else(|| "?".into());
    let end_time = value
        .end
        .as_ref()
        .map(|value| short_time(&value.date_time))
        .unwrap_or_else(|| "?".into());
    let time_zone = value
        .start
        .as_ref()
        .and_then(|value| value.time_zone.clone())
        .or_else(|| value.recurrence.range.recurrence_time_zone.clone())
        .or_else(|| value.end.as_ref().and_then(|value| value.time_zone.clone()));
    let days = value
        .recurrence
        .pattern
        .days_of_week
        .iter()
        .map(|day| day_label(day))
        .collect();

    WorkPlanRecurrenceSummary {
        days,
        start_time,
        end_time,
        time_zone,
        location: value.work_location_type,
    }
}

/// Return whether Microsoft 365 says the current instant is inside the user's
/// work plan. A time-off occurrence takes precedence over a working occurrence.
pub async fn working_now(graph: &GraphClient, now: DateTime<Utc>) -> Result<bool> {
    let occurrences: Vec<WorkPlanOccurrence> = graph.get_page(&occurrences_view_path(now)).await?;
    Ok(occurrences_mean_working_now(&occurrences))
}

/// Read the current state plus recurring hours/location rules for diagnostics.
///
/// This uses the same Microsoft 365 work-hours-and-locations resource as the
/// ntfy work-day gate. The project already requests Calendars.ReadWrite, which
/// is sufficient for these delegated calls.
pub async fn diagnostics(graph: &GraphClient, now: DateTime<Utc>) -> Result<WorkPlanDiagnostics> {
    let working_now = working_now(graph, now).await?;
    let recurrences: Vec<WorkPlanRecurrence> = graph
        .get_collection("me/settings/workHoursAndLocations/recurrences?$top=50")
        .await?;

    Ok(WorkPlanDiagnostics {
        working_now,
        recurrences: recurrences
            .into_iter()
            .map(summarize_recurrence)
            .collect(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    fn occurrence(kind: &str) -> WorkPlanOccurrence {
        WorkPlanOccurrence {
            work_location_type: Some(kind.to_string()),
        }
    }

    #[test]
    fn working_types_are_recognized() {
        assert!(occurrences_mean_working_now(&[occurrence("office")]));
        assert!(occurrences_mean_working_now(&[occurrence("remote")]));
        assert!(occurrences_mean_working_now(&[occurrence("unspecified")]));
        assert!(!occurrences_mean_working_now(&[]));
        assert!(!occurrences_mean_working_now(&[occurrence(
            "unknownFutureValue"
        )]));
    }

    #[test]
    fn time_off_overrides_working_occurrences() {
        assert!(!occurrences_mean_working_now(&[
            occurrence("remote"),
            occurrence("timeOff"),
        ]));
    }

    #[test]
    fn path_uses_a_one_second_utc_window() {
        let now = Utc.with_ymd_and_hms(2026, 9, 22, 8, 30, 45).unwrap();
        assert_eq!(
            occurrences_view_path(now),
            "me/settings/workHoursAndLocations/occurrencesView(startDateTime='2026-09-22T08:30:45Z',endDateTime='2026-09-22T08:30:46Z')?$select=workLocationType&$top=20"
        );
    }

    #[test]
    fn recurrence_summary_is_safe_and_compact() {
        let value: WorkPlanRecurrence = serde_json::from_value(serde_json::json!({
            "workLocationType": "remote",
            "start": {
                "dateTime": "2026-09-21T08:00:00.0000000",
                "timeZone": "Central Europe Standard Time"
            },
            "end": {
                "dateTime": "2026-09-21T16:30:00.0000000",
                "timeZone": "Central Europe Standard Time"
            },
            "recurrence": {
                "pattern": {
                    "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"]
                },
                "range": {
                    "recurrenceTimeZone": "Central Europe Standard Time"
                }
            }
        }))
        .unwrap();

        let summary = summarize_recurrence(value);
        assert_eq!(summary.days, ["Mon", "Tue", "Wed", "Thu", "Fri"]);
        assert_eq!(summary.start_time, "08:00");
        assert_eq!(summary.end_time, "16:30");
        assert_eq!(
            summary.time_zone.as_deref(),
            Some("Central Europe Standard Time")
        );
        assert_eq!(summary.location.as_deref(), Some("remote"));
    }
}
