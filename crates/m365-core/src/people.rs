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

/// Set the signed-in user's Teams presence status message.
///
/// The Graph API accepts only text content. Omitting expiryDateTime means the
/// message remains until it is replaced or explicitly cleared.
pub async fn set_status_message(graph: &GraphClient, user_id: &str, message: &str) -> Result<()> {
    graph
        .post_action(
            &format!("users/{user_id}/presence/setStatusMessage"),
            &json!({
                "statusMessage": {
                    "message": {
                        "content": message,
                        "contentType": "text"
                    }
                }
            }),
        )
        .await
}

/// Clear the Teams presence status message by publishing empty text.
pub async fn clear_status_message(graph: &GraphClient, user_id: &str) -> Result<()> {
    set_status_message(graph, user_id, "").await
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

/// Resolve a user id from an email address (used to open a chat with an email
fn exact_profile_match_index(
    people: &[Person],
    user_id: Option<&str>,
    email: Option<&str>,
    display_name: &str,
) -> Option<usize> {
    if let Some(index) = user_id.and_then(|id| {
        people
            .iter()
            .position(|person| person.id.as_deref() == Some(id))
    }) {
        return Some(index);
    }

    if let Some(email) = email.map(str::trim).filter(|value| !value.is_empty()) {
        if let Some(index) = people.iter().position(|person| {
            person
                .user_principal_name
                .as_deref()
                .is_some_and(|upn| upn.eq_ignore_ascii_case(email))
                || person.scored_email_addresses.iter().any(|candidate| {
                    candidate
                        .address
                        .as_deref()
                        .is_some_and(|address| address.eq_ignore_ascii_case(email))
                })
        }) {
            return Some(index);
        }
    }

    let name = display_name.trim();
    if name.is_empty() {
        return None;
    }

    let mut matches = people
        .iter()
        .enumerate()
        .filter(|(_, person)| {
            person
                .display_name
                .as_deref()
                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(name))
        })
        .map(|(index, _)| index);

    let first = matches.next()?;
    if matches.next().is_some() {
        None
    } else {
        Some(first)
    }
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct DirectoryUserProfile {
    id: String,
    #[serde(default)]
    display_name: Option<String>,
    #[serde(default)]
    given_name: Option<String>,
    #[serde(default)]
    surname: Option<String>,
    #[serde(default)]
    job_title: Option<String>,
    #[serde(default)]
    department: Option<String>,
    #[serde(default)]
    office_location: Option<String>,
    #[serde(default)]
    company_name: Option<String>,
    #[serde(default)]
    mail: Option<String>,
    #[serde(default)]
    user_principal_name: Option<String>,
    #[serde(default)]
    business_phones: Vec<String>,
    #[serde(default)]
    mobile_phone: Option<String>,
}

fn directory_user_into_person(user: DirectoryUserProfile) -> Result<Person> {
    let mut phones = Vec::new();

    for number in user.business_phones {
        let number = number.trim();
        if !number.is_empty() {
            phones.push(serde_json::json!({
                "type": "business",
                "number": number,
            }));
        }
    }

    if let Some(number) = user
        .mobile_phone
        .as_deref()
        .map(str::trim)
        .filter(|number| !number.is_empty())
    {
        phones.push(serde_json::json!({
            "type": "mobile",
            "number": number,
        }));
    }

    let email = user
        .mail
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .or_else(|| {
            user.user_principal_name
                .as_deref()
                .map(str::trim)
                .filter(|value| !value.is_empty())
        });

    let scored_email_addresses = email
        .map(|address| vec![serde_json::json!({ "address": address })])
        .unwrap_or_default();

    Ok(serde_json::from_value(serde_json::json!({
        "id": user.id,
        "displayName": user.display_name,
        "givenName": user.given_name,
        "surname": user.surname,
        "jobTitle": user.job_title,
        "companyName": user.company_name,
        "department": user.department,
        "officeLocation": user.office_location,
        "userPrincipalName": user.user_principal_name,
        "scoredEmailAddresses": scored_email_addresses,
        "phones": phones,
    }))?)
}

/// Resolve an internal Teams contact directly from the Entra directory.
///
/// Unlike `/me/people`, `/users/{id}` is deterministic: the Teams member GUID
/// identifies the directory user exactly.
pub async fn directory_profile_for_contact(graph: &GraphClient, user_id: &str) -> Result<Person> {
    let path = format!(
        "users/{user_id}?$select=id,displayName,givenName,surname,jobTitle,department,officeLocation,companyName,mail,userPrincipalName,businessPhones,mobilePhone"
    );
    let user: DirectoryUserProfile = graph.get_json(&path).await?;
    let person = directory_user_into_person(user)?;

    Ok(person)
}

/// Resolve the richest People.Read profile available for a Teams contact.
///
/// `/me/people` can include organization users, local contacts and recent
/// communication peers without adding a broad directory-read permission.
pub async fn profile_for_contact(
    graph: &GraphClient,
    user_id: Option<&str>,
    email: Option<&str>,
    display_name: &str,
) -> Result<Option<Person>> {
    let query = email
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| display_name.trim());

    if query.is_empty() {
        return Ok(None);
    }

    let people = relevant_people(graph, Some(query)).await?;

    if people.is_empty() {
        return Ok(None);
    }

    let Some(index) = exact_profile_match_index(&people, user_id, email, display_name) else {
        return Ok(None);
    };

    Ok(people.into_iter().nth(index))
}

/// Best-effort profile photo fetch. 403/404 are handled by the caller as
/// "avatar unavailable" so profile display never depends on photo permission.
pub async fn profile_photo(graph: &GraphClient, user_id: &str) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!("users/{user_id}/photo/$value"))
        .await
}

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

/// Presence for one user id.
///
/// Used as a fallback for federated contacts when the batch endpoint omits
/// a cross-tenant user even though direct presence lookup is allowed.
pub async fn presence(graph: &GraphClient, user_id: &str) -> Result<Presence> {
    graph.get_json(&format!("users/{user_id}/presence")).await
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

pub async fn search_users(graph: &GraphClient, query: &str) -> Result<Vec<User>> {
    let Some(path) = user_search_path(query) else {
        return Ok(Vec::new());
    };
    graph.get_page(&path).await
}

pub async fn photo(graph: &GraphClient, user_id: &str) -> Result<Vec<u8>> {
    graph
        .get_bytes(&format!("users/{user_id}/photo/$value"))
        .await
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

#[cfg(test)]
mod contact_profile_tests {
    use super::*;

    fn person(name: &str, email: &str) -> Person {
        serde_json::from_value(serde_json::json!({
            "id": format!("person-{name}"),
            "displayName": name,
            "userPrincipalName": email,
            "scoredEmailAddresses": [{"address": email}]
        }))
        .expect("valid person fixture")
    }

    #[test]
    fn directory_user_maps_profile_and_phones() {
        let directory: DirectoryUserProfile = serde_json::from_value(serde_json::json!({
            "id": "61110baf-7952-404a-89ec-8354ca7099ef",
            "displayName": "Test User",
            "givenName": "Test",
            "surname": "User",
            "jobTitle": "Engineer",
            "department": "Development",
            "officeLocation": "Ostrava",
            "companyName": "Example",
            "mail": "test@example.com",
            "userPrincipalName": "test@example.com",
            "businessPhones": ["+420 555 111 222", "+420 555 333 444"],
            "mobilePhone": "+420 777 111 222"
        }))
        .expect("valid directory user fixture");

        let person =
            directory_user_into_person(directory).expect("directory profile maps to person");

        assert_eq!(person.job_title.as_deref(), Some("Engineer"));
        assert_eq!(person.department.as_deref(), Some("Development"));
        assert_eq!(person.office_location.as_deref(), Some("Ostrava"));
        assert_eq!(person.phones.len(), 3);
        assert_eq!(person.phones[0].phone_type.as_deref(), Some("business"));
        assert_eq!(person.phones[0].number.as_deref(), Some("+420 555 111 222"));
        assert_eq!(person.phones[2].phone_type.as_deref(), Some("mobile"));
        assert_eq!(person.phones[2].number.as_deref(), Some("+420 777 111 222"));
    }

    #[test]
    fn profile_match_never_falls_back_to_first_result() {
        let people = vec![
            person("Signed In User", "me@example.com"),
            person("Someone Else", "other@example.com"),
        ];

        assert_eq!(
            exact_profile_match_index(
                &people,
                Some("opaque_external_id"),
                None,
                "External Contact"
            ),
            None
        );
    }

    #[test]
    fn profile_match_accepts_exact_email() {
        let people = vec![
            person("Signed In User", "me@example.com"),
            person("External Contact", "external@example.net"),
        ];

        assert_eq!(
            exact_profile_match_index(
                &people,
                None,
                Some("external@example.net"),
                "External Contact"
            ),
            Some(1)
        );
    }

    #[test]
    fn profile_match_rejects_ambiguous_exact_name() {
        let people = vec![
            person("Alex Smith", "alex.one@example.com"),
            person("Alex Smith", "alex.two@example.com"),
        ];

        assert_eq!(
            exact_profile_match_index(&people, None, None, "Alex Smith"),
            None
        );
    }
}
