//! Read-only diagnostic path for the internal Teams presence stack.
//!
//! This module intentionally exposes only sanitized stage results. Access
//! tokens, Skype tokens, lookup addresses and MRIs stay local to the async
//! probe and are never returned to the TUI.

use std::time::Duration;

use serde::Deserialize;
use serde_json::{json, Value};

use crate::auth::Authenticator;

pub const SKYPE_SCOPE: &str = "https://api.spaces.skype.com/.default";
const AUTHZ_HOSTS: &[&str] = &[
    "https://teams.cloud.microsoft/api/authsvc/v1.0/authz",
    "https://teams.microsoft.com/api/authsvc/v1.0/authz",
];
const PRESENCE_FALLBACK: &str = "https://presence.teams.microsoft.com";
const CLIENT_VERSION: &str = "1415/1.0.0.2023031528";

#[derive(Debug, Clone)]
pub enum DiagnosticStep {
    Ok(String),
    Skipped(String),
    Error(String),
}

#[derive(Debug, Clone)]
pub struct Diagnostics {
    pub lookup_address_available: bool,
    pub resource_token: DiagnosticStep,
    pub authz: DiagnosticStep,
    pub middle_tier_lookup: DiagnosticStep,
    pub resolved_mri_shape: String,
    pub ups_presence: DiagnosticStep,
}

impl Diagnostics {
    fn stopped(
        lookup_address_available: bool,
        resource_token: DiagnosticStep,
        authz: DiagnosticStep,
        middle_tier_lookup: DiagnosticStep,
        resolved_mri_shape: String,
        ups_presence: DiagnosticStep,
    ) -> Self {
        Self {
            lookup_address_available,
            resource_token,
            authz,
            middle_tier_lookup,
            resolved_mri_shape,
            ups_presence,
        }
    }
}

struct TeamsSession {
    skype_token: String,
    region: String,
    presence_host: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SearchUser {
    #[serde(default)]
    mri: String,
    #[serde(default)]
    email: String,
    #[serde(default)]
    user_principal_name: String,
}

#[derive(Debug, Deserialize)]
struct PresenceEnvelope {
    #[serde(default)]
    mri: String,
    #[serde(default)]
    presence: Option<PresenceBody>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct PresenceBody {
    #[serde(default)]
    availability: String,
    #[serde(default)]
    activity: String,
}

#[derive(Debug, Clone)]
pub struct ConsentProbe {
    pub interactive_consent: DiagnosticStep,
    pub existing_refresh_token: DiagnosticStep,
}

/// Interactive consent probe for the Teams/Skype resource.
///
/// The interactive token is discarded. After consent, the method deliberately
/// re-tests the pre-existing Graph refresh token; success there proves normal
/// F7 diagnostics can proceed later without storing a second token cache.
pub async fn consent_probe<F>(auth: &Authenticator, on_prompt: F) -> ConsentProbe
where
    F: FnOnce(crate::auth::DeviceCodePrompt),
{
    match auth.resource_consent_probe(SKYPE_SCOPE, on_prompt).await {
        Ok(()) => {
            let existing_refresh_token = match auth.resource_access_token(SKYPE_SCOPE).await {
                Ok(token) => {
                    drop(token);
                    DiagnosticStep::Ok(
                        "existing Graph refresh token can acquire resource token".into(),
                    )
                }
                Err(error) => DiagnosticStep::Error(format!(
                    "consent succeeded; cached refresh-token test: {}",
                    explain_resource_error(&error.to_string())
                )),
            };
            ConsentProbe {
                interactive_consent: DiagnosticStep::Ok(
                    "consent completed; probe token discarded".into(),
                ),
                existing_refresh_token,
            }
        }
        Err(error) => ConsentProbe {
            interactive_consent: DiagnosticStep::Error(explain_resource_error(
                &error.to_string(),
            )),
            existing_refresh_token: DiagnosticStep::Skipped(
                "interactive consent did not complete".into(),
            ),
        },
    }
}

pub fn explain_resource_error(raw: &str) -> String {
    if raw.contains("AADSTS65001") {
        "consent required (AADSTS65001)".into()
    } else if raw.contains("AADSTS65002") {
        "resource requires Microsoft preauthorization (AADSTS65002)".into()
    } else if raw.contains("AADSTS650057") {
        "resource missing from app registration (AADSTS650057)".into()
    } else if raw.contains("AADSTS65005") {
        "client is not configured for this resource (AADSTS65005)".into()
    } else if raw.contains("AADSTS65004") || raw.contains("declined") {
        "consent declined (AADSTS65004)".into()
    } else if raw.contains("AADSTS70011") {
        "invalid resource scope combination (AADSTS70011)".into()
    } else if raw.contains("400 Bad Request") {
        "resource token request rejected (HTTP 400)".into()
    } else if raw.contains("401 Unauthorized") {
        "resource token request unauthorized (HTTP 401)".into()
    } else if raw.contains("403 Forbidden") {
        "resource token request forbidden (HTTP 403)".into()
    } else {
        "resource token request failed".into()
    }
}

pub async fn diagnose(
    auth: &Authenticator,
    lookup_address: Option<&str>,
    existing_mri: Option<&str>,
    prefer_consumer: bool,
) -> Diagnostics {
    let lookup_address = lookup_address
        .map(str::trim)
        .filter(|value| !value.is_empty());
    let existing_mri = existing_mri
        .map(str::trim)
        .filter(|value| is_user_mri(value));

    let token = match auth.resource_access_token(SKYPE_SCOPE).await {
        Ok(token) => token,
        Err(error) => {
            return Diagnostics::stopped(
                lookup_address.is_some(),
                DiagnosticStep::Error(explain_resource_error(&error.to_string())),
                DiagnosticStep::Skipped("resource token unavailable".into()),
                DiagnosticStep::Skipped("resource token unavailable".into()),
                existing_mri.map(mri_shape).unwrap_or("none").to_string(),
                DiagnosticStep::Skipped("resource token unavailable".into()),
            );
        }
    };

    let client = match reqwest::Client::builder()
        .timeout(Duration::from_secs(12))
        .build()
    {
        Ok(client) => client,
        Err(_) => {
            return Diagnostics::stopped(
                lookup_address.is_some(),
                DiagnosticStep::Ok("acquired".into()),
                DiagnosticStep::Error("HTTP client setup failed".into()),
                DiagnosticStep::Skipped("Teams session unavailable".into()),
                existing_mri.map(mri_shape).unwrap_or("none").to_string(),
                DiagnosticStep::Skipped("Teams session unavailable".into()),
            );
        }
    };

    let session = match acquire_teams_session(&client, &token).await {
        Ok(session) => session,
        Err(error) => {
            return Diagnostics::stopped(
                lookup_address.is_some(),
                DiagnosticStep::Ok("acquired".into()),
                DiagnosticStep::Error(error),
                DiagnosticStep::Skipped("Teams authz unavailable".into()),
                existing_mri.map(mri_shape).unwrap_or("none").to_string(),
                DiagnosticStep::Skipped("Teams authz unavailable".into()),
            );
        }
    };

    let (mri, middle_tier_lookup) = if let Some(mri) = existing_mri {
        (
            Some(mri.to_string()),
            DiagnosticStep::Skipped("MRI already available from chat identity".into()),
        )
    } else if let Some(address) = lookup_address {
        match lookup_mri(&client, &token, &session, address, prefer_consumer).await {
            Ok(Some(mri)) => (
                Some(mri),
                DiagnosticStep::Ok("user MRI resolved".into()),
            ),
            Ok(None) => (
                None,
                DiagnosticStep::Error("no unambiguous user MRI returned".into()),
            ),
            Err(error) => (None, DiagnosticStep::Error(error)),
        }
    } else {
        (
            None,
            DiagnosticStep::Skipped("no lookup address available".into()),
        )
    };

    let Some(mri) = mri else {
        return Diagnostics::stopped(
            lookup_address.is_some(),
            DiagnosticStep::Ok("acquired".into()),
            DiagnosticStep::Ok("Skype token/session acquired".into()),
            middle_tier_lookup,
            "none".into(),
            DiagnosticStep::Skipped("no user MRI available".into()),
        );
    };

    let shape = mri_shape(&mri).to_string();
    let ups_presence = match read_presence(&client, &token, &session, &mri).await {
        Ok(Some((availability, activity))) => {
            DiagnosticStep::Ok(format!("{availability} · {activity}"))
        }
        Ok(None) => DiagnosticStep::Error("UPS returned no presence for the MRI".into()),
        Err(error) => DiagnosticStep::Error(error),
    };

    Diagnostics {
        lookup_address_available: lookup_address.is_some(),
        resource_token: DiagnosticStep::Ok("acquired".into()),
        authz: DiagnosticStep::Ok("Skype token/session acquired".into()),
        middle_tier_lookup,
        resolved_mri_shape: shape,
        ups_presence,
    }
}

async fn acquire_teams_session(
    client: &reqwest::Client,
    resource_token: &str,
) -> Result<TeamsSession, String> {
    let mut last_status = None;

    for host in AUTHZ_HOSTS {
        let response = client
            .post(*host)
            .bearer_auth(resource_token)
            .header("content-type", "application/json")
            .body("{}")
            .send()
            .await
            .map_err(|_| "Teams authz transport error".to_string())?;

        let status = response.status();
        if !status.is_success() {
            last_status = Some(status.to_string());
            continue;
        }

        let value: Value = response
            .json()
            .await
            .map_err(|_| "Teams authz response could not be parsed".to_string())?;

        let skype_token = value
            .pointer("/tokens/skypeToken")
            .and_then(Value::as_str)
            .filter(|value| !value.is_empty())
            .ok_or_else(|| "Teams authz returned no Skype token".to_string())?
            .to_string();

        let region = value
            .get("region")
            .and_then(Value::as_str)
            .filter(|value| safe_region(value))
            .unwrap_or("emea")
            .to_ascii_lowercase();

        let presence_host = value
            .pointer("/regionGtms/unifiedPresence")
            .and_then(Value::as_str)
            .filter(|value| safe_presence_host(value))
            .unwrap_or(PRESENCE_FALLBACK)
            .trim_end_matches('/')
            .to_string();

        return Ok(TeamsSession {
            skype_token,
            region,
            presence_host,
        });
    }

    Err(match last_status {
        Some(status) => format!("Teams authz failed: HTTP {status}"),
        None => "Teams authz failed".into(),
    })
}

async fn lookup_mri(
    client: &reqwest::Client,
    resource_token: &str,
    session: &TeamsSession,
    lookup_address: &str,
    prefer_consumer: bool,
) -> Result<Option<String>, String> {
    let base = format!(
        "https://teams.microsoft.com/api/mt/{}/beta/users/",
        session.region
    );
    let mut url = reqwest::Url::parse(&base)
        .map_err(|_| "Middle Tier lookup URL could not be built".to_string())?;
    {
        let mut segments = url
            .path_segments_mut()
            .map_err(|_| "Middle Tier lookup URL could not be built".to_string())?;
        segments.push(lookup_address);
        segments.push("externalsearchv3");
    }
    url.query_pairs_mut()
        .append_pair("includeTFLUsers", "true");

    let response = client
        .get(url)
        .bearer_auth(resource_token)
        .header("x-ms-client-version", CLIENT_VERSION)
        .send()
        .await
        .map_err(|_| "Middle Tier lookup transport error".to_string())?;

    let status = response.status();
    if !status.is_success() {
        return Err(format!("Middle Tier lookup failed: HTTP {status}"));
    }

    let users: Vec<SearchUser> = response
        .json()
        .await
        .map_err(|_| "Middle Tier lookup response could not be parsed".to_string())?;

    let mut exact = Vec::new();
    let mut candidates = Vec::new();

    for user in users {
        if !is_user_mri(&user.mri) {
            continue;
        }
        if prefer_consumer && !user.mri.to_ascii_lowercase().starts_with("8:live:") {
            continue;
        }

        let address_matches = user.email.eq_ignore_ascii_case(lookup_address)
            || user
                .user_principal_name
                .eq_ignore_ascii_case(lookup_address);

        if address_matches {
            exact.push(user.mri.clone());
        }
        candidates.push(user.mri);
    }

    exact.sort();
    exact.dedup();
    candidates.sort();
    candidates.dedup();

    if exact.len() == 1 {
        Ok(exact.pop())
    } else if exact.is_empty() && candidates.len() == 1 {
        Ok(candidates.pop())
    } else {
        Ok(None)
    }
}

async fn read_presence(
    client: &reqwest::Client,
    resource_token: &str,
    session: &TeamsSession,
    mri: &str,
) -> Result<Option<(String, String)>, String> {
    let url = format!("{}/v1/presence/getpresence/", session.presence_host);
    let response = client
        .post(url)
        .bearer_auth(resource_token)
        .header("x-skypetoken", &session.skype_token)
        .header("content-type", "application/json")
        .json(&json!([{ "mri": mri }]))
        .send()
        .await
        .map_err(|_| "UPS presence transport error".to_string())?;

    let status = response.status();
    if !status.is_success() {
        return Err(format!("UPS presence failed: HTTP {status}"));
    }

    let items: Vec<PresenceEnvelope> = response
        .json()
        .await
        .map_err(|_| "UPS presence response could not be parsed".to_string())?;

    Ok(items
        .into_iter()
        .find(|item| item.mri.eq_ignore_ascii_case(mri))
        .and_then(|item| item.presence)
        .map(|presence| {
            let availability = nonempty_or_unknown(presence.availability);
            let activity = nonempty_or_unknown(presence.activity);
            (availability, activity)
        }))
}

fn nonempty_or_unknown(value: String) -> String {
    let value = value.trim();
    if value.is_empty() {
        "unknown".into()
    } else {
        value.chars().take(64).collect()
    }
}

fn safe_region(value: &str) -> bool {
    !value.is_empty()
        && value.len() <= 24
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-')
}

fn safe_presence_host(value: &str) -> bool {
    reqwest::Url::parse(value).ok().is_some_and(|url| {
        url.scheme() == "https"
            && url
                .host_str()
                .is_some_and(|host| {
                    host.eq_ignore_ascii_case("presence.teams.microsoft.com")
                        || host.ends_with(".teams.microsoft.com")
                })
    })
}

fn is_user_mri(value: &str) -> bool {
    let lower = value.trim().to_ascii_lowercase();
    lower.starts_with("8:live:")
        || lower.starts_with("8:orgid:")
        || lower.starts_with("8:teamsvisitor:")
}

fn mri_shape(value: &str) -> &'static str {
    let lower = value.trim().to_ascii_lowercase();
    if lower.starts_with("8:live:") {
        "MRI consumer (8:live:)"
    } else if lower.starts_with("8:orgid:") {
        "MRI orgid (8:orgid:)"
    } else if lower.starts_with("8:teamsvisitor:") {
        "MRI visitor (8:teamsvisitor:)"
    } else {
        "unknown MRI"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resource_auth_errors_are_actionable_and_redacted() {
        assert_eq!(
            explain_resource_error("invalid_grant something private AADSTS65001 more private"),
            "consent required (AADSTS65001)"
        );
        assert_eq!(
            explain_resource_error("AADSTS65002 secret details"),
            "resource requires Microsoft preauthorization (AADSTS65002)"
        );
        assert_eq!(
            explain_resource_error("AADSTS650057 secret details"),
            "resource missing from app registration (AADSTS650057)"
        );
        assert_eq!(
            explain_resource_error("AADSTS65005 secret details"),
            "client is not configured for this resource (AADSTS65005)"
        );
        assert_eq!(
            explain_resource_error("unexpected secret body"),
            "resource token request failed"
        );
    }

    #[test]
    fn mri_shapes_are_redacted_classifications() {
        assert_eq!(mri_shape("8:live:.cid.secret"), "MRI consumer (8:live:)");
        assert_eq!(
            mri_shape("8:orgid:123e4567-e89b-12d3-a456-426614174000"),
            "MRI orgid (8:orgid:)"
        );
        assert_eq!(
            mri_shape("8:teamsvisitor:secret"),
            "MRI visitor (8:teamsvisitor:)"
        );
    }

    #[test]
    fn region_and_presence_host_validation_are_strict() {
        assert!(safe_region("emea"));
        assert!(!safe_region("../emea"));
        assert!(safe_presence_host("https://presence.teams.microsoft.com"));
        assert!(safe_presence_host("https://foo.teams.microsoft.com"));
        assert!(!safe_presence_host("http://presence.teams.microsoft.com"));
        assert!(!safe_presence_host("https://example.com"));
    }
}
