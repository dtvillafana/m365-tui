//! Azure Communication Services: HMAC-signed REST for Call Automation.
//!
//! The client Calling SDK (WebRTC) is browser / Windows / mobile only. On
//! Linux the equivalent is this REST surface: join a Teams meeting or place a
//! call to Teams users, then stream PCM over a WebSocket that ACS opens to us.
//! Audio devices are someone else's problem (the TUI uses sox).

use anyhow::{Context, Result};
use chrono::{DateTime, Datelike, Timelike, Utc};
use hmac::{Hmac, Mac};
use serde::Deserialize;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

use crate::util::{base64_decode, base64_encode};

const API_VERSION: &str = "2025-06-15";

type HmacSha256 = Hmac<Sha256>;

/// Endpoint + access key parsed from an ACS connection string.
#[derive(Debug, Clone)]
pub struct AcsConfig {
    /// `https://<name>.communication.azure.com`, no trailing slash.
    pub endpoint: String,
    /// Base64 access key as it appears in the portal connection string.
    pub access_key: String,
}

impl AcsConfig {
    /// Parse `endpoint=https://…/;accesskey=…`. Keys are matched
    /// case-insensitively; extra segments are ignored.
    pub fn from_connection_string(s: &str) -> Result<Self> {
        let mut endpoint = None;
        let mut access_key = None;
        for part in s.split(';') {
            let part = part.trim();
            if part.is_empty() {
                continue;
            }
            let Some((k, v)) = part.split_once('=') else {
                continue;
            };
            match k.trim().to_ascii_lowercase().as_str() {
                "endpoint" => endpoint = Some(v.trim().trim_end_matches('/').to_string()),
                "accesskey" => access_key = Some(v.trim().to_string()),
                _ => {}
            }
        }
        let endpoint = endpoint.context("ACS connection string missing endpoint=")?;
        let access_key = access_key.context("ACS connection string missing accesskey=")?;
        if !endpoint.starts_with("https://") && !endpoint.starts_with("http://") {
            anyhow::bail!("ACS endpoint must be an http(s) URL, got {endpoint}");
        }
        Ok(Self {
            endpoint,
            access_key,
        })
    }

    fn host(&self) -> Result<&str> {
        self.endpoint
            .strip_prefix("https://")
            .or_else(|| self.endpoint.strip_prefix("http://"))
            .map(|h| h.split('/').next().unwrap_or(h))
            .filter(|h| !h.is_empty())
            .context("ACS endpoint has no host")
    }
}

/// Who to ring, or which meeting to join.
#[derive(Debug, Clone)]
pub enum CallTarget {
    /// Teams meeting join URL (`https://teams.microsoft.com/l/meetup-join/…`).
    Meeting { join_url: String },
    /// Entra object ids of Teams users to call.
    Users { ids: Vec<String> },
}

/// Identifiers returned when a call is created or connected.
#[derive(Debug, Clone)]
pub struct CallConnection {
    pub call_connection_id: String,
}

#[derive(Clone)]
pub struct AcsClient {
    http: reqwest::Client,
    config: AcsConfig,
}

impl AcsClient {
    pub fn new(config: AcsConfig) -> Self {
        Self {
            http: reqwest::Client::builder()
                .user_agent("m365-tui/0.1")
                .build()
                .expect("building reqwest client"),
            config,
        }
    }

    /// Join a Teams meeting, or place a 1:1 / group call to Teams users.
    ///
    /// `callback_uri` and `media_ws_uri` must be reachable from ACS (HTTPS /
    /// WSS). Media streaming is bidirectional PCM 16 kHz mono.
    pub async fn start_call(
        &self,
        target: &CallTarget,
        callback_uri: &str,
        media_ws_uri: &str,
        display_name: &str,
    ) -> Result<CallConnection> {
        let media = media_streaming_options(media_ws_uri);
        match target {
            CallTarget::Meeting { join_url } => {
                let body = json!({
                    "callbackUri": callback_uri,
                    "sourceDisplayName": display_name,
                    "callLocator": {
                        "kind": "teamsMeetingLink",
                        "teamsMeetingLink": join_url,
                    },
                    "mediaStreamingOptions": media,
                });
                self.post("/calling/callConnections:connect", &body)
                    .await
                    .map(parse_connection)
                    .and_then(|r| r)
            }
            CallTarget::Users { ids } => {
                let targets: Vec<Value> = ids
                    .iter()
                    .map(|id| {
                        json!({
                            "kind": "microsoftTeamsUser",
                            "microsoftTeamsUserId": id,
                            "cloud": "public",
                        })
                    })
                    .collect();
                let body = json!({
                    "callbackUri": callback_uri,
                    "sourceDisplayName": display_name,
                    "targets": targets,
                    "mediaStreamingOptions": media,
                });
                self.post("/calling/callConnections", &body)
                    .await
                    .map(parse_connection)
                    .and_then(|r| r)
            }
        }
    }

    /// Leave the call without ending it for everyone else.
    pub async fn hangup(&self, call_connection_id: &str) -> Result<()> {
        let path = format!("/calling/callConnections/{call_connection_id}:hangUp");
        let body = json!({ "forEveryone": false });
        self.post(&path, &body).await?;
        Ok(())
    }

    async fn post(&self, path: &str, body: &Value) -> Result<Vec<u8>> {
        let serialized = serde_json::to_vec(body).context("serializing ACS request")?;
        let path_and_query = format!("{path}?api-version={API_VERSION}");
        let url = format!("{}{path_and_query}", self.config.endpoint);
        let host = self.config.host()?.to_string();
        let date = format_rfc1123(Utc::now());
        let content_hash = sha256_b64(&serialized);
        let string_to_sign = format!("POST\n{path_and_query}\n{date};{host};{content_hash}");
        let signature = hmac_sign(&self.config.access_key, &string_to_sign)?;
        let authorization = format!(
            "HMAC-SHA256 SignedHeaders=x-ms-date;host;x-ms-content-sha256&Signature={signature}"
        );

        let resp = self
            .http
            .post(&url)
            .header("x-ms-date", &date)
            .header("x-ms-content-sha256", &content_hash)
            .header("Authorization", authorization)
            .header(reqwest::header::CONTENT_TYPE, "application/json")
            .header("Repeatability-Request-ID", uuid::Uuid::new_v4().to_string())
            .header("Repeatability-First-Sent", &date)
            .body(serialized)
            .send()
            .await
            .context("sending ACS request")?;

        let status = resp.status();
        let bytes = resp.bytes().await.context("reading ACS body")?;
        if !status.is_success() {
            let text = String::from_utf8_lossy(&bytes);
            anyhow::bail!("ACS request failed ({status}): {text}");
        }
        Ok(bytes.to_vec())
    }
}

fn media_streaming_options(media_ws_uri: &str) -> Value {
    json!({
        "transportUrl": media_ws_uri,
        "transportType": "websocket",
        "contentType": "audio",
        "audioChannelType": "mixed",
        "startMediaStreaming": true,
        "enableBidirectional": true,
        "audioFormat": "pcm16KMono",
    })
}

fn parse_connection(bytes: Vec<u8>) -> Result<CallConnection> {
    #[derive(Deserialize)]
    #[serde(rename_all = "camelCase")]
    struct Resp {
        call_connection_id: Option<String>,
    }
    let parsed: Resp = serde_json::from_slice(&bytes).context("decoding ACS call response")?;
    let call_connection_id = parsed
        .call_connection_id
        .filter(|s| !s.is_empty())
        .context("ACS response missing callConnectionId")?;
    Ok(CallConnection { call_connection_id })
}

fn sha256_b64(data: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(data);
    base64_encode(&hasher.finalize())
}

fn hmac_sign(access_key_b64: &str, string_to_sign: &str) -> Result<String> {
    let key = base64_decode(access_key_b64).map_err(|e| anyhow::anyhow!("ACS access key: {e}"))?;
    let mut mac = HmacSha256::new_from_slice(&key).context("HMAC key")?;
    mac.update(string_to_sign.as_bytes());
    Ok(base64_encode(&mac.finalize().into_bytes()))
}

/// HTTP-date (RFC 1123) in GMT, locale-independent.
pub fn format_rfc1123(dt: DateTime<Utc>) -> String {
    const DAYS: [&str; 7] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    const MONTHS: [&str; 12] = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    format!(
        "{}, {:02} {} {:04} {:02}:{:02}:{:02} GMT",
        DAYS[dt.weekday().num_days_from_monday() as usize],
        dt.day(),
        MONTHS[dt.month0() as usize],
        dt.year(),
        dt.hour(),
        dt.minute(),
        dt.second()
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    #[test]
    fn parses_a_portal_connection_string() {
        let c = AcsConfig::from_connection_string(
            "endpoint=https://contoso.communication.azure.com/;accesskey=abc+def==",
        )
        .unwrap();
        assert_eq!(c.endpoint, "https://contoso.communication.azure.com");
        assert_eq!(c.access_key, "abc+def==");
        assert_eq!(c.host().unwrap(), "contoso.communication.azure.com");
    }

    #[test]
    fn connection_string_is_case_insensitive_and_strips_slash() {
        let c = AcsConfig::from_connection_string(
            "Endpoint=https://x.communication.azure.com/;AccessKey=qq==",
        )
        .unwrap();
        assert_eq!(c.endpoint, "https://x.communication.azure.com");
        assert_eq!(c.access_key, "qq==");
    }

    #[test]
    fn empty_body_sha256_matches_the_known_digest() {
        // SHA-256 of the empty string, standard base64.
        assert_eq!(
            sha256_b64(b""),
            "47DEQpj8HBSa+/TImW+5JCeuQeRkm5NMpJWZG3hSuFU="
        );
    }

    #[test]
    fn hmac_is_stable_for_a_fixed_key_and_payload() {
        // key = base64("secret")
        let sig = hmac_sign("c2VjcmV0", "POST\n/x\ndate;host;hash").unwrap();
        let again = hmac_sign("c2VjcmV0", "POST\n/x\ndate;host;hash").unwrap();
        assert_eq!(sig, again);
        assert_ne!(
            hmac_sign("c2VjcmV0", "GET\n/x\ndate;host;hash").unwrap(),
            sig
        );
    }

    #[test]
    fn rfc1123_is_gmt_and_english() {
        let dt = Utc.with_ymd_and_hms(2026, 9, 9, 12, 3, 4).unwrap();
        assert_eq!(format_rfc1123(dt), "Wed, 09 Sep 2026 12:03:04 GMT");
    }

    #[test]
    fn rejects_a_connection_string_without_an_endpoint() {
        assert!(AcsConfig::from_connection_string("accesskey=qq==").is_err());
    }
}
