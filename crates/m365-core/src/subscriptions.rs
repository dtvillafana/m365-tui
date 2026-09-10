//! Microsoft Graph change-notification subscriptions.
//!
//! The TUI (which holds the delegated token) creates subscriptions pointing at
//! the tunnel-fronted webhook. We use `includeResourceData: false` so no
//! encryption certificate is needed — notifications merely *signal* a change and
//! the TUI then does a targeted delta fetch ("notify-then-delta").
//!
//! Note: Teams chat-message subscriptions expire in ~1h and require a
//! `lifecycleNotificationUrl` for longer lifetimes; a renewal loop in the TUI
//! keeps them alive.

use anyhow::Result;
use chrono::{Duration, Utc};
use serde::Deserialize;
use serde_json::json;

use crate::graph::GraphClient;

/// Well-known resources we subscribe to.
pub const RES_INBOX: &str = "me/mailFolders('inbox')/messages";

/// All chats the user takes part in.
///
/// Graph rejects the `/me/` shorthand here — it answers 403 "User may only
/// create user-scoped chat message subscriptions for their own messages" — so
/// the signed-in user's id has to be spelled out.
pub fn res_all_chats(user_id: &str) -> String {
    format!("users/{user_id}/chats/getAllMessages")
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Subscription {
    pub id: String,
    #[serde(default)]
    pub resource: Option<String>,
    #[serde(default)]
    pub change_type: Option<String>,
    #[serde(default)]
    pub expiration_date_time: Option<String>,
}

/// Create a subscription. `expiration_minutes` is clamped by Graph per resource.
#[allow(clippy::too_many_arguments)]
pub async fn create(
    graph: &GraphClient,
    resource: &str,
    change_type: &str,
    notification_url: &str,
    lifecycle_url: Option<&str>,
    client_state: &str,
    expiration_minutes: i64,
) -> Result<Subscription> {
    let expiry = (Utc::now() + Duration::minutes(expiration_minutes))
        .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

    let mut body = json!({
        "changeType": change_type,
        "notificationUrl": notification_url,
        "resource": resource,
        "expirationDateTime": expiry,
        "clientState": client_state,
        "includeResourceData": false,
    });
    if let Some(lc) = lifecycle_url {
        body["lifecycleNotificationUrl"] = json!(lc);
    }

    graph.post_json("subscriptions", &body).await
}

/// Extend a subscription's expiry (call well before it lapses).
pub async fn renew(graph: &GraphClient, id: &str, expiration_minutes: i64) -> Result<()> {
    let expiry = (Utc::now() + Duration::minutes(expiration_minutes))
        .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);
    graph
        .patch(
            &format!("subscriptions/{id}"),
            &json!({ "expirationDateTime": expiry }),
        )
        .await
}

pub async fn delete(graph: &GraphClient, id: &str) -> Result<()> {
    graph.delete(&format!("subscriptions/{id}")).await
}

/// List existing subscriptions owned by this app for the signed-in user.
pub async fn list(graph: &GraphClient) -> Result<Vec<Subscription>> {
    graph.get_collection("subscriptions").await
}
