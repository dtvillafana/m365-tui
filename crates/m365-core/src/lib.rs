//! Shared engine for the M365 TUI.
//!
//! [`Session`] bundles configuration, the authenticator, and a ready
//! [`GraphClient`]. The `mail`, `calendar`, `chats`, `channels`, and `people`
//! modules are thin endpoint wrappers over that client; `subscriptions` and
//! `events` implement the notify-then-delta real-time path.

pub mod acs;
pub mod auth;
pub mod calendar;
pub mod channels;
pub mod chats;
pub mod config;
pub mod events;
pub mod graph;
pub mod hosted;
pub mod mail;
pub mod models;
pub mod people;
pub mod subscriptions;
pub mod util;

use std::sync::Arc;

use anyhow::Result;

pub use acs::{AcsClient, AcsConfig, CallConnection, CallTarget};
pub use auth::{Authenticator, DeviceCodePrompt};
pub use config::Config;
pub use graph::GraphClient;
pub use hosted::{HostedImage, OutgoingBody};

/// Everything a frontend needs: config, auth, and a Graph client.
#[derive(Clone)]
pub struct Session {
    pub config: Arc<Config>,
    pub auth: Arc<Authenticator>,
    pub graph: GraphClient,
}

impl Session {
    /// Build a session from configuration. Does not perform any network I/O;
    /// call [`Session::ensure_logged_in`] before making Graph calls.
    pub fn new(config: Config) -> Self {
        let auth = Arc::new(Authenticator::new(&config));
        let graph = GraphClient::new(auth.clone());
        Self {
            config: Arc::new(config),
            auth,
            graph,
        }
    }

    /// Convenience constructor: load `.env`, then the environment.
    pub fn from_env() -> Result<Self> {
        Config::load_dotenv();
        Ok(Self::new(Config::from_env()?))
    }

    /// Ensure we hold a usable token, running device-code login (invoking
    /// `on_prompt`) only if there are no cached credentials.
    pub async fn ensure_logged_in<F>(&self, on_prompt: F) -> Result<()>
    where
        F: FnOnce(DeviceCodePrompt),
    {
        if self.auth.has_credentials().await && self.auth.access_token().await.is_ok() {
            return Ok(());
        }
        self.auth.login(on_prompt).await
    }

    /// Fetch the signed-in user (`/me`).
    pub async fn whoami(&self) -> Result<models::User> {
        people::me(&self.graph).await
    }
}
