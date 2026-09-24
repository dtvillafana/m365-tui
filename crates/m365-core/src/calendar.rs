//! Outlook calendar endpoints.

use anyhow::Result;
use serde_json::json;

use crate::graph::GraphClient;
use crate::models::Event;

/// Events overlapping the given ISO-8601 UTC window, ordered by start time.
/// `start`/`end` look like `2026-07-23T00:00:00Z`.
pub async fn calendar_view(graph: &GraphClient, start: &str, end: &str) -> Result<Vec<Event>> {
    let path = format!(
        "me/calendarView?startDateTime={start}&endDateTime={end}\
         &$orderby=start/dateTime&$top=100\
         &$select=id,subject,start,end,organizer,attendees,location,isOnlineMeeting,onlineMeeting,bodyPreview,responseStatus,isOrganizer,isCancelled,isAllDay,responseRequested"
    );
    graph.get_collection(&path).await
}

pub async fn get_event(graph: &GraphClient, id: &str) -> Result<Event> {
    graph.get_json(&format!("me/events/{id}")).await
}

/// RSVP response kind.
#[derive(Debug, Clone, Copy)]
pub enum Rsvp {
    Accept,
    Decline,
    Tentative,
}

impl Rsvp {
    fn action(self) -> &'static str {
        match self {
            Rsvp::Accept => "accept",
            Rsvp::Decline => "decline",
            Rsvp::Tentative => "tentativelyAccept",
        }
    }
}

/// Respond to a meeting invitation.
pub async fn respond(graph: &GraphClient, event_id: &str, rsvp: Rsvp, comment: &str) -> Result<()> {
    let path = format!("me/events/{event_id}/{}", rsvp.action());
    graph
        .post_action(&path, &json!({ "sendResponse": true, "comment": comment }))
        .await
}

/// Create a simple event.
pub async fn create_event(
    graph: &GraphClient,
    subject: &str,
    start: &str,
    end: &str,
    time_zone: &str,
    attendees: &[String],
) -> Result<Event> {
    let attendees: Vec<_> = attendees
        .iter()
        .map(|a| json!({ "emailAddress": { "address": a }, "type": "required" }))
        .collect();
    let payload = json!({
        "subject": subject,
        "start": { "dateTime": start, "timeZone": time_zone },
        "end": { "dateTime": end, "timeZone": time_zone },
        "attendees": attendees,
    });
    graph.post_json("me/events", &payload).await
}
