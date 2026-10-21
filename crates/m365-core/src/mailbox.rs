//! Outlook mailbox settings used by the Out of Office UI.

use anyhow::Result;
use serde_json::json;

use crate::graph::GraphClient;
use crate::models::{AutomaticRepliesSetting, MailboxSettingsSnapshot};

/// Fetch the mailbox preferred time zone together with automatic replies.
pub async fn settings(graph: &GraphClient) -> Result<MailboxSettingsSnapshot> {
    graph
        .get_json("me/mailboxSettings?$select=timeZone,automaticRepliesSetting")
        .await
}

/// Replace the signed-in user's automatic-reply settings.
pub async fn set_automatic_replies(
    graph: &GraphClient,
    setting: &AutomaticRepliesSetting,
) -> Result<()> {
    graph
        .patch(
            "me/mailboxSettings",
            &json!({ "automaticRepliesSetting": setting }),
        )
        .await
}

/// Disable automatic replies without changing unrelated mailbox settings.
pub async fn clear_automatic_replies(graph: &GraphClient) -> Result<()> {
    graph
        .patch(
            "me/mailboxSettings",
            &json!({
              "automaticRepliesSetting": {
                "status": "disabled"
              }
            }),
        )
        .await
}
