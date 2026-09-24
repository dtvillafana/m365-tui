//! People lookup and presence — used for cross-navigation (email <-> chat) and
//! status dots.

use anyhow::Result;
use serde_json::json;

use crate::graph::GraphClient;
use crate::models::{Person, Presence, User};

pub async fn me(graph: &GraphClient) -> Result<User> {
    graph
        .get_json("me?$select=id,displayName,mail,userPrincipalName,jobTitle")
        .await
}

/// The signed-in user's current presence.
pub async fn my_presence(graph: &GraphClient) -> Result<Presence> {
    graph.get_json("me/presence").await
}

/// Set the signed-in user's preferred presence (the sticky "set status" in
/// Teams). Valid pairs: Available/Available, Busy/Busy,
/// DoNotDisturb/DoNotDisturb, BeRightBack/BeRightBack, Away/Away, Offline/OffWork.
pub async fn set_preferred_presence(
    graph: &GraphClient,
    availability: &str,
    activity: &str,
) -> Result<()> {
    graph
        .post_action(
            "me/presence/setUserPreferredPresence",
            &json!({ "availability": availability, "activity": activity }),
        )
        .await
}

/// Register this app as a *presence session* for the user.
///
/// `setUserPreferredPresence` only records a preference; the status a colleague
/// sees comes from an active session, which is normally the Teams client. An app
/// may hold its own session, which is what makes a status visible with no Teams
/// client running. Sessions expire (5 min – 4 h), so this must be re-asserted.
///
/// `session_id` must be the application (client) ID.
pub async fn set_session_presence(
    graph: &GraphClient,
    session_id: &str,
    availability: &str,
    activity: &str,
    expiration: &str,
) -> Result<()> {
    graph
        .post_action(
            "me/presence/setPresence",
            &json!({
                "sessionId": session_id,
                "availability": availability,
                "activity": activity,
                "expirationDuration": expiration,
            }),
        )
        .await
}

/// Drop this app's presence session, so the user stops appearing online because
/// of us. Called when a status is cleared and on exit.
pub async fn clear_session_presence(graph: &GraphClient, session_id: &str) -> Result<()> {
    graph
        .post_action(
            "me/presence/clearPresence",
            &json!({ "sessionId": session_id }),
        )
        .await
}

/// Clear the preferred presence, reverting to automatically-calculated status.
pub async fn clear_preferred_presence(graph: &GraphClient) -> Result<()> {
    graph
        .post_action("me/presence/clearUserPreferredPresence", &json!({}))
        .await
}

/// Relevant people for the signed-in user, optionally filtered by a search term
/// (matches name or email).
pub async fn relevant_people(graph: &GraphClient, search: Option<&str>) -> Result<Vec<Person>> {
    let path = match search {
        Some(q) => format!("me/people?$search=\"{}\"&$top=25", q.replace('"', "")),
        None => "me/people?$top=25".to_string(),
    };
    graph.get_collection(&path).await
}

/// Search directory users by display name or the beginning of their username
/// (UPN). Requires delegated User.ReadBasic.All.
pub async fn search_users(graph: &GraphClient, query: &str) -> Result<Vec<User>> {
    let Some(path) = user_search_path(query) else {
        return Ok(Vec::new());
    };
    graph.get_page(&path).await
}

fn user_search_path(query: &str) -> Option<String> {
    let escaped = query.trim().replace('\'', "''");
    if escaped.is_empty() {
        return None;
    }
    let mut url = reqwest::Url::parse("https://graph.microsoft.com/v1.0/users").ok()?;
    url.query_pairs_mut()
        .append_pair("$select", "id,displayName,mail,userPrincipalName")
        .append_pair("$top", "25")
        .append_pair(
            "$filter",
            &format!(
                "startswith(displayName,'{escaped}') or startswith(userPrincipalName,'{escaped}')"
            ),
        );
    // Pass only the relative path so M365_GRAPH_BASE still works for mocks.
    Some(format!("users?{}", url.query().unwrap_or_default()))
}

/// Resolve a user id from an email address (used to open a chat with an email
/// sender). Returns `None` if the address is not a known directory user.
pub async fn user_id_for_email(graph: &GraphClient, email: &str) -> Result<Option<String>> {
    // /users/{email} accepts the UPN/mail directly for directory members.
    match graph
        .get_json::<User>(&format!("users/{email}?$select=id"))
        .await
    {
        Ok(u) => Ok(Some(u.id)),
        Err(_) => Ok(None),
    }
}

/// Profile photo bytes (`GET /users/{id}/photo/$value`).
///
/// `User.Read` covers the signed-in user. Other people's photos need
/// `User.Read.All` or `ProfilePhoto.Read.All`; without those Graph returns 404
/// and the caller should treat it as "no photo".
pub async fn photo(graph: &GraphClient, user_id: &str) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!("users/{user_id}/photo/$value"))
        .await
}

/// Presence for a set of user ids (Teams status dots).
pub async fn presences(graph: &GraphClient, user_ids: &[String]) -> Result<Vec<Presence>> {
    if user_ids.is_empty() {
        return Ok(Vec::new());
    }
    let payload = json!({ "ids": user_ids });
    #[derive(serde::Deserialize)]
    struct Wrapper {
        value: Vec<Presence>,
    }
    let w: Wrapper = graph
        .post_json("communications/getPresencesByUserId", &payload)
        .await?;
    Ok(w.value)
}

#[cfg(test)]
mod tests {
    use super::user_search_path;

    #[test]
    fn directory_search_encodes_and_escapes_username() {
        assert!(user_search_path("  ").is_none());
        let path = user_search_path(" O'Brien+test ").unwrap();
        let url = reqwest::Url::parse(&format!("https://graph.microsoft.com/v1.0/{path}")).unwrap();
        let filter = url
            .query_pairs()
            .find(|(key, _)| key == "$filter")
            .unwrap()
            .1;
        assert_eq!(filter, "startswith(displayName,'O''Brien+test') or startswith(userPrincipalName,'O''Brien+test')");
    }
}
