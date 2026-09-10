//! Teams calling: ACS Call Automation for signaling, sox for the audio device.
//!
//! ACS opens a WebSocket to a small local server (published via `cloudflared`
//! or `M365_CALL_PUBLIC_URL`) and streams bidirectional PCM 16 kHz mono. sox
//! plays the far end and captures the microphone. There is no Linux client
//! Calling SDK; this is the REST equivalent.

use std::process::Stdio;
use std::sync::Arc;
use std::time::Duration;

use anyhow::{Context, Result};
use axum::extract::ws::{Message, WebSocket, WebSocketUpgrade};
use axum::extract::State;
use axum::http::StatusCode;
use axum::response::IntoResponse;
use axum::routing::{get, post};
use axum::Router;
use m365_core::{AcsClient, CallTarget};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};
use tokio::process::{Child, Command};
use tokio::sync::{mpsc, watch, Mutex, Notify};
use tokio::task::JoinHandle;

/// 20 ms of 16 kHz 16-bit mono PCM.
const FRAME_BYTES: usize = 640;

// Dropping a JoinHandle detaches it; call-scoped tasks must stop on every exit.
struct CallTask<T>(JoinHandle<T>);

impl<T> Drop for CallTask<T> {
    fn drop(&mut self) {
        self.0.abort();
    }
}

/// Mid-call signal back to the UI. The function returns when the call is over.
pub enum CallEvent {
    /// Audio is flowing.
    Live,
}

pub struct CallOpts {
    pub acs: AcsClient,
    pub target: CallTarget,
    pub display_name: String,
    pub public_url: Option<String>,
}

/// Handles the user can poke while a call is up.
pub struct CallControls {
    hangup: watch::Sender<bool>,
    muted: watch::Sender<bool>,
}

impl CallControls {
    pub fn hangup(&self) {
        let _ = self.hangup.send(true);
    }

    pub fn set_muted(&self, muted: bool) {
        let _ = self.muted.send(muted);
    }
}

pub fn new_controls() -> (CallControls, watch::Receiver<bool>, watch::Receiver<bool>) {
    let (hangup_tx, hangup_rx) = watch::channel(false);
    let (muted_tx, muted_rx) = watch::channel(false);
    (
        CallControls {
            hangup: hangup_tx,
            muted: muted_tx,
        },
        hangup_rx,
        muted_rx,
    )
}

pub fn sox_available() -> bool {
    have_bin("sox")
}

pub fn can_publish() -> bool {
    have_bin("cloudflared")
        || std::env::var("M365_CALL_PUBLIC_URL")
            .map(|s| !s.trim().is_empty())
            .unwrap_or(false)
}

fn have_bin(name: &str) -> bool {
    std::process::Command::new(name)
        .arg("--version")
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .ok()
        .map(|s| s.success())
        .unwrap_or(false)
}

/// Run a call until hangup, remote disconnect, or error.
pub async fn run_call(
    opts: CallOpts,
    events: mpsc::Sender<CallEvent>,
    mut hangup: watch::Receiver<bool>,
    muted: watch::Receiver<bool>,
) -> Result<()> {
    if !sox_available() {
        anyhow::bail!("sox is not installed — needed to play and record call audio");
    }

    let (to_speaker_tx, to_speaker_rx) = mpsc::channel::<Vec<u8>>(32);
    let (from_mic_tx, from_mic_rx) = mpsc::channel::<Vec<u8>>(32);
    let ws_live = Arc::new(Notify::new());
    let failed = Arc::new(Mutex::new(None::<Result<(), String>>));

    let state = CallHttp {
        to_speaker: to_speaker_tx,
        from_mic: Arc::new(Mutex::new(Some(from_mic_rx))),
        muted: muted.clone(),
        ws_live: ws_live.clone(),
        failed: failed.clone(),
    };

    // Random port when cloudflared will advertise it; a stable one when the
    // user is pointing their own tunnel at us.
    let bind = if opts.public_url.is_some() {
        let port: u16 = std::env::var("M365_CALL_BIND")
            .unwrap_or_else(|_| "8788".into())
            .parse()
            .context("M365_CALL_BIND must be a port number")?;
        format!("127.0.0.1:{port}")
    } else {
        "127.0.0.1:0".into()
    };
    let listener = tokio::net::TcpListener::bind(&bind)
        .await
        .with_context(|| format!("binding the local calling server on {bind}"))?;
    let local_port = listener.local_addr()?.port();

    let probe = format!("/health/{local_port}");
    let app = Router::new()
        .route(&probe, get(|| async { "m365-call-ready" }))
        .route("/callback", post(callback))
        .route("/media", get(media))
        .with_state(state);

    let (shutdown_tx, shutdown_rx) = tokio::sync::oneshot::channel::<()>();
    let _server = CallTask(tokio::spawn(async move {
        axum::serve(listener, app)
            .with_graceful_shutdown(async {
                let _ = shutdown_rx.await;
            })
            .await
    }));

    let mut tunnel = None;
    let public = match opts.public_url.clone() {
        Some(url) => url,
        None => {
            let (url, child) = spawn_cloudflared(local_port).await?;
            tunnel = Some(child);
            url
        }
    };
    let public = public.trim_end_matches('/').to_string();
    tokio::select! {
        result = wait_for_public_url(&public, &probe, local_port, tunnel.is_some()) => { result?; }
        _ = hangup.changed() => return Ok(()),
    }
    let callback_uri = format!("{public}/callback");
    let media_ws_uri = http_to_ws(&public) + "/media";

    let mut play = spawn_sox_play().context("starting sox playback")?;
    let mut rec = spawn_sox_rec().context("starting sox capture")?;
    let mut play_in = play.stdin.take().context("sox play stdin")?;
    let mut rec_out = rec.stdout.take().context("sox rec stdout")?;

    // Device failures usually happen just after spawn; report them before ringing.
    tokio::time::sleep(Duration::from_millis(150)).await;
    check_audio(&mut play, &mut rec).await?;
    if *hangup.borrow() {
        return Ok(());
    }

    let play_task = CallTask(tokio::spawn(async move {
        let mut rx = to_speaker_rx;
        while let Some(frame) = rx.recv().await {
            if play_in.write_all(&frame).await.is_err() {
                break;
            }
        }
    }));
    let rec_task = CallTask(tokio::spawn(async move {
        let mut buf = vec![0u8; FRAME_BYTES];
        while rec_out.read_exact(&mut buf).await.is_ok() {
            if from_mic_tx.send(buf.clone()).await.is_err() {
                break;
            }
        }
    }));

    let started = opts
        .acs
        .start_call(
            &opts.target,
            &callback_uri,
            &media_ws_uri,
            &opts.display_name,
        )
        .await;
    let connection = match started {
        Ok(c) => c,
        Err(e) => {
            let _ = shutdown_tx.send(());
            let _ = play.kill().await;
            let _ = rec.kill().await;
            if let Some(mut t) = tunnel {
                let _ = t.kill().await;
            }
            return Err(e);
        }
    };

    let acs = opts.acs.clone();
    let conn_id = connection.call_connection_id.clone();
    let stream_failed = failed.clone();
    let live_task = CallTask(tokio::spawn(async move {
        if tokio::time::timeout(Duration::from_secs(90), ws_live.notified())
            .await
            .is_ok()
        {
            let _ = events.send(CallEvent::Live).await;
        } else {
            let mut outcome = stream_failed.lock().await;
            if outcome.is_none() {
                *outcome = Some(Err("ACS did not open the media WebSocket within 90 seconds; check tunnel routing and Teams interop".into()));
            }
        }
    }));
    let outcome = wait_until_done(&mut hangup, &mut play, &mut rec, &mut tunnel, &failed).await;

    let _ = acs.hangup(&conn_id).await;
    let _ = shutdown_tx.send(());
    let _ = play.kill().await;
    let _ = rec.kill().await;
    if let Some(mut t) = tunnel {
        let _ = t.kill().await;
    }
    drop((play_task, rec_task, live_task));
    outcome
}

async fn wait_until_done(
    hangup: &mut watch::Receiver<bool>,
    play: &mut Child,
    rec: &mut Child,
    tunnel: &mut Option<Child>,
    failed: &Arc<Mutex<Option<Result<(), String>>>>,
) -> Result<()> {
    loop {
        if *hangup.borrow() {
            return Ok(());
        }
        tokio::select! {
            changed = hangup.changed() => {
                if changed.is_err() || *hangup.borrow() {
                    return Ok(());
                }
            }
            _ = tokio::time::sleep(Duration::from_millis(250)) => {
                check_audio(play, rec).await?;
                if let Some(child) = tunnel.as_mut() {
                    if let Some(status) = child.try_wait().context("polling cloudflared")? {
                        anyhow::bail!("cloudflared exited ({status})");
                    }
                }
                let reason = failed.lock().await.clone();
                if let Some(reason) = reason {
                    return reason.map_err(anyhow::Error::msg);
                }
            }
        }
    }
}

/// New quick-tunnel records can be hidden by a local resolver's negative cache.
/// Fall back only for this call's generated hostname; TLS still verifies it.
struct QuickTunnelDns {
    host: String,
    http: reqwest::Client,
}

impl reqwest::dns::Resolve for QuickTunnelDns {
    fn resolve(&self, name: reqwest::dns::Name) -> reqwest::dns::Resolving {
        let host = self.host.clone();
        let http = self.http.clone();
        Box::pin(async move {
            let local = tokio::time::timeout(
                Duration::from_secs(2),
                tokio::net::lookup_host((name.as_str(), 0)),
            )
            .await;
            match local {
                Ok(Ok(addrs)) => {
                    let addrs: Vec<_> = addrs.collect();
                    if !addrs.is_empty() {
                        return Ok(Box::new(addrs.into_iter()) as reqwest::dns::Addrs);
                    }
                }
                Ok(Err(error)) if name.as_str() != host => return Err(error.into()),
                Err(error) if name.as_str() != host => return Err(error.into()),
                _ => {}
            }
            if name.as_str() != host {
                return Err(std::io::Error::other("DNS returned no addresses").into());
            }

            tracing::debug!(%host, "local DNS lookup failed; trying DNS-over-HTTPS for calling tunnel");
            let response: serde_json::Value = http
                .get("https://cloudflare-dns.com/dns-query")
                .header("Accept", "application/dns-json")
                .query(&[("name", host.as_str()), ("type", "A")])
                .send()
                .await?
                .error_for_status()?
                .json()
                .await?;
            let addrs: Vec<std::net::SocketAddr> = response
                .get("Answer")
                .and_then(|answers| answers.as_array())
                .into_iter()
                .flatten()
                .filter(|answer| answer.get("type").and_then(|ty| ty.as_u64()) == Some(1))
                .filter_map(|answer| {
                    answer
                        .get("data")?
                        .as_str()?
                        .parse::<std::net::Ipv4Addr>()
                        .ok()
                })
                .map(|ip| std::net::SocketAddr::from((ip, 0)))
                .collect();
            if response.get("Status").and_then(|status| status.as_u64()) != Some(0)
                || addrs.is_empty()
            {
                return Err(std::io::Error::other(format!(
                    "DNS-over-HTTPS returned no addresses for {host}"
                ))
                .into());
            }
            Ok(Box::new(addrs.into_iter()) as reqwest::dns::Addrs)
        })
    }
}

async fn wait_for_public_url(
    public: &str,
    probe: &str,
    local_port: u16,
    quick_tunnel: bool,
) -> Result<reqwest::Client> {
    let url = reqwest::Url::parse(public).context("invalid M365_CALL_PUBLIC_URL")?;
    anyhow::ensure!(
        url.scheme() == "https"
            && url.host_str().is_some()
            && url.query().is_none()
            && url.fragment().is_none(),
        "calling requires a public HTTPS base URL without a query or fragment"
    );
    let mut builder = reqwest::Client::builder()
        .timeout(Duration::from_secs(10))
        .redirect(reqwest::redirect::Policy::none());
    if quick_tunnel {
        builder = builder.dns_resolver(Arc::new(QuickTunnelDns {
            host: url.host_str().unwrap().to_string(),
            http: reqwest::Client::builder()
                .timeout(Duration::from_secs(5))
                .build()?,
        }));
    }
    let http = builder.build()?;
    let mut last_error = "no response".to_string();
    let wait = Duration::from_secs(if quick_tunnel { 120 } else { 30 });
    let ready = tokio::time::timeout(wait, async {
        loop {
            match http.get(format!("{public}{probe}")).send().await {
                Ok(response) => {
                    let status = response.status();
                    if status.is_success()
                        && response.text().await.unwrap_or_default() == "m365-call-ready"
                    {
                        return;
                    }
                    last_error = format!("HTTP {status}, unexpected health response");
                }
                Err(error) => last_error = format!("{:#}", anyhow::Error::new(error)),
            }
            tokio::time::sleep(Duration::from_millis(500)).await;
        }
    })
    .await;
    ready.with_context(|| format!("calling tunnel {public} is not reachable after {} seconds ({last_error}): check DNS resolution and forwarding to 127.0.0.1:{local_port} for /health/*, /callback, and /media (WebSocket upgrades)", wait.as_secs()))?;
    Ok(http)
}

async fn check_audio(play: &mut Child, rec: &mut Child) -> Result<()> {
    for (name, child) in [("playback", play), ("capture", rec)] {
        if let Some(status) = child
            .try_wait()
            .with_context(|| format!("polling sox {name}"))?
        {
            anyhow::bail!(
                "sox {name} exited ({status}): {}",
                child_stderr(child).await
            );
        }
    }
    Ok(())
}

async fn child_stderr(child: &mut Child) -> String {
    let mut text = String::new();
    if let Some(stderr) = child.stderr.take() {
        let _ = stderr.take(8192).read_to_string(&mut text).await;
    }
    text.trim().to_string()
}

fn http_to_ws(url: &str) -> String {
    if let Some(rest) = url.strip_prefix("https://") {
        format!("wss://{rest}")
    } else if let Some(rest) = url.strip_prefix("http://") {
        format!("ws://{rest}")
    } else {
        url.to_string()
    }
}

#[derive(Clone)]
struct CallHttp {
    to_speaker: mpsc::Sender<Vec<u8>>,
    from_mic: Arc<Mutex<Option<mpsc::Receiver<Vec<u8>>>>>,
    muted: watch::Receiver<bool>,
    ws_live: Arc<Notify>,
    failed: Arc<Mutex<Option<Result<(), String>>>>,
}

async fn callback(State(state): State<CallHttp>, body: String) -> impl IntoResponse {
    if let Some(reason) = callback_failure(&body) {
        *state.failed.lock().await = Some(reason);
    }
    StatusCode::OK
}

/// Pull a failure reason out of an ACS callback payload, if any.
fn callback_failure(body: &str) -> Option<Result<(), String>> {
    let value: serde_json::Value = serde_json::from_str(body).ok()?;
    let events = if value.is_array() {
        value.as_array()?.clone()
    } else {
        vec![value]
    };
    for ev in events {
        let ty = ev.get("type").and_then(|t| t.as_str()).unwrap_or("");
        if ty.ends_with("Failed") {
            let data = ev.get("data").unwrap_or(&ev);
            let detail = data
                .get("resultInformation")
                .map(|r| r.to_string())
                .unwrap_or_else(|| "no resultInformation".into());
            return Some(Err(format!("{ty}: {detail}")));
        }
        if ty.ends_with("CallDisconnected") {
            return Some(Ok(()));
        }
    }
    None
}

async fn media(ws: WebSocketUpgrade, State(state): State<CallHttp>) -> impl IntoResponse {
    ws.on_upgrade(move |socket| handle_media(socket, state))
}

async fn handle_media(mut socket: WebSocket, state: CallHttp) {
    let mut from_mic = {
        let mut slot = state.from_mic.lock().await;
        match slot.take() {
            Some(rx) => rx,
            None => return, // already consumed by a previous socket
        }
    };
    state.ws_live.notify_one();
    let mut muted = state.muted.clone();
    loop {
        tokio::select! {
            msg = socket.recv() => {
                match msg {
                    Some(Ok(Message::Text(text))) => {
                        if let Some(pcm) = audio_payload(&text) {
                            let _ = state.to_speaker.send(pcm).await;
                        }
                    }
                    Some(Ok(Message::Close(_))) | None => break,
                    Some(Ok(_)) => {}
                    Some(Err(_)) => break,
                }
            }
            frame = from_mic.recv() => {
                let Some(mut pcm) = frame else { break };
                if *muted.borrow() {
                    pcm.fill(0);
                }
                let json = outbound_audio(&pcm);
                if socket.send(Message::Text(json.into())).await.is_err() {
                    break;
                }
            }
            _ = muted.changed() => {}
        }
    }
    let mut outcome = state.failed.lock().await;
    if outcome.is_none() {
        *outcome = Some(Err("ACS media WebSocket closed".into()));
    }
}

fn audio_payload(text: &str) -> Option<Vec<u8>> {
    let v: serde_json::Value = serde_json::from_str(text).ok()?;
    let kind = v.get("kind").and_then(|k| k.as_str()).unwrap_or("");
    if !kind.eq_ignore_ascii_case("AudioData") {
        return None;
    }
    let data = v
        .get("audioData")
        .and_then(|a| a.get("data"))
        .and_then(|d| d.as_str())?;
    if v.get("audioData")
        .and_then(|a| a.get("silent"))
        .and_then(|s| s.as_bool())
        == Some(true)
    {
        return None;
    }
    m365_core::util::base64_decode(data).ok()
}

fn outbound_audio(pcm: &[u8]) -> String {
    let data = m365_core::util::base64_encode(pcm);
    format!(r#"{{"kind":"AudioData","audioData":{{"data":"{data}"}}}}"#)
}

fn spawn_sox_play() -> Result<Child> {
    Command::new("sox")
        .args([
            "-q",
            "--buffer",
            "640",
            "-t",
            "raw",
            "-r",
            "16000",
            "-e",
            "signed-integer",
            "-b",
            "16",
            "-c",
            "1",
            "-L",
            "-",
            "-d",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::null())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .context("spawning sox (play)")
}

fn spawn_sox_rec() -> Result<Child> {
    Command::new("sox")
        .args([
            "-q",
            "--buffer",
            "640",
            "-d",
            "-t",
            "raw",
            "-r",
            "16000",
            "-e",
            "signed-integer",
            "-b",
            "16",
            "-c",
            "1",
            "-L",
            "-",
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .context("spawning sox (record)")
}

async fn spawn_cloudflared(port: u16) -> Result<(String, Child)> {
    if !have_bin("cloudflared") {
        anyhow::bail!(
            "cloudflared is not installed, and M365_CALL_PUBLIC_URL is unset — ACS cannot reach this machine"
        );
    }
    let mut child = Command::new("cloudflared")
        .args([
            "tunnel",
            "--url",
            &format!("http://127.0.0.1:{port}"),
            "--no-autoupdate",
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .spawn()
        .context("spawning cloudflared")?;

    let stdout = child.stdout.take();
    let stderr = child.stderr.take();
    let (url_tx, mut url_rx) = mpsc::channel::<String>(1);

    let scrape = |stream: Option<tokio::process::ChildStdout>,
                  err: Option<tokio::process::ChildStderr>,
                  tx: mpsc::Sender<String>| async move {
        // cloudflared prints the URL on stderr; read both to be safe.
        let mut handles: Vec<JoinHandle<()>> = Vec::new();
        if let Some(out) = stream {
            let tx = tx.clone();
            handles.push(tokio::spawn(async move {
                let mut lines = BufReader::new(out).lines();
                while let Ok(Some(line)) = lines.next_line().await {
                    tracing::debug!("cloudflared: {line}");
                    if let Some(url) = parse_tunnel_url(&line) {
                        let _ = tx.try_send(url);
                    }
                }
            }));
        }
        if let Some(err) = err {
            let tx = tx.clone();
            handles.push(tokio::spawn(async move {
                let mut lines = BufReader::new(err).lines();
                while let Ok(Some(line)) = lines.next_line().await {
                    tracing::debug!("cloudflared: {line}");
                    if let Some(url) = parse_tunnel_url(&line) {
                        let _ = tx.try_send(url);
                    }
                }
            }));
        }
        for h in handles {
            let _ = h.await;
        }
    };
    tokio::spawn(scrape(stdout, stderr, url_tx));

    match tokio::time::timeout(Duration::from_secs(25), url_rx.recv()).await {
        Ok(Some(url)) => Ok((url, child)),
        _ => {
            let _ = child.kill().await;
            anyhow::bail!("cloudflared did not print a public URL")
        }
    }
}

/// Only accept quick-tunnel hostnames, never documentation links in the logs.
pub fn parse_tunnel_url(line: &str) -> Option<String> {
    let start = line.find("https://")?;
    let rest = &line[start..];
    let end = rest
        .find(|c: char| c.is_whitespace() || matches!(c, '|' | '"' | '\'' | ')' | ']'))
        .unwrap_or(rest.len());
    let url = rest[..end].trim_end_matches('/');
    let parsed = reqwest::Url::parse(url).ok()?;
    if parsed.host_str()?.ends_with(".trycloudflare.com")
        && parsed.path() == "/"
        && parsed.query().is_none()
        && parsed.fragment().is_none()
    {
        Some(url.to_string())
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scrapes_a_trycloudflare_url_out_of_log_noise() {
        let line = "2026-09-09T12:00:00Z INF |  https://random-words-here.trycloudflare.com  |";
        assert_eq!(
            parse_tunnel_url(line).as_deref(),
            Some("https://random-words-here.trycloudflare.com")
        );
    }

    #[test]
    fn ignores_lines_without_a_url() {
        assert!(parse_tunnel_url("starting tunnel").is_none());
        assert!(parse_tunnel_url("Terms: https://www.cloudflare.com/website-terms/").is_none());
        assert!(parse_tunnel_url("https://fake.trycloudflare.com.example.org").is_none());
    }

    #[test]
    fn http_to_ws_rewrites_the_scheme() {
        assert_eq!(http_to_ws("https://ex.com"), "wss://ex.com");
        assert_eq!(http_to_ws("http://localhost:8"), "ws://localhost:8");
    }

    #[test]
    fn silent_frames_are_dropped() {
        let json = r#"{"kind":"AudioData","audioData":{"data":"AAA=","silent":true}}"#;
        assert!(audio_payload(json).is_none());
    }

    #[test]
    fn audio_frames_decode() {
        let pcm = vec![0u8, 1, 2];
        let json = outbound_audio(&pcm);
        assert_eq!(audio_payload(&json), Some(pcm));
    }

    #[test]
    fn failed_callback_is_detected() {
        let body = r#"[{"type":"Microsoft.Communication.CreateCallFailed","resultInformation":{"message":"user not found"}}]"#;
        let failure = callback_failure(body).unwrap().unwrap_err();
        assert!(failure.contains("CreateCallFailed"));
        assert!(failure.contains("user not found"));
        assert!(
            callback_failure(r#"[{"type":"Microsoft.Communication.CallConnected"}]"#).is_none()
        );
    }

    #[test]
    fn cloud_event_failure_preserves_acs_codes() {
        let body = r#"[{"type":"Microsoft.Communication.CreateCallFailed","data":{"resultInformation":{"code":403,"subCode":12345,"message":"Forbidden"}}}]"#;
        let reason = callback_failure(body).unwrap().unwrap_err();
        assert!(reason.contains("403"));
        assert!(reason.contains("12345"));
        assert!(reason.contains("Forbidden"));
        assert_eq!(
            callback_failure(r#"{"type":"Microsoft.Communication.CallDisconnected","data":{}}"#),
            Some(Ok(()))
        );
    }

    #[tokio::test]
    #[ignore = "requires cloudflared and outbound network; does not place a call"]
    async fn live_quick_tunnel_reaches_local_server() {
        let _ = tracing_subscriber::fmt()
            .with_env_filter("m365=debug")
            .try_init();
        let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
        let port = listener.local_addr().unwrap().port();
        let app = Router::new()
            .route("/health/test", get(|| async { "m365-call-ready" }))
            .route(
                "/media",
                get(|ws: WebSocketUpgrade| async {
                    ws.on_upgrade(|mut socket| async move {
                        let _ = socket.send(Message::Close(None)).await;
                    })
                }),
            );
        let _server = CallTask(tokio::spawn(
            async move { axum::serve(listener, app).await },
        ));
        let (public, mut tunnel) = spawn_cloudflared(port).await.unwrap();
        let result = async {
            let http = wait_for_public_url(&public, "/health/test", port, true).await?;
            let response = http
                .get(format!("{public}/media"))
                .timeout(Duration::from_secs(10))
                .header("Connection", "Upgrade")
                .header("Upgrade", "websocket")
                .header("Sec-WebSocket-Version", "13")
                .header("Sec-WebSocket-Key", "dGhlIHNhbXBsZSBub25jZQ==")
                .send()
                .await?;
            anyhow::ensure!(
                response.status() == StatusCode::SWITCHING_PROTOCOLS,
                "public WebSocket upgrade failed: {}",
                response.status()
            );
            Ok::<_, anyhow::Error>(())
        }
        .await;
        let _ = tunnel.kill().await;
        result.unwrap();
    }
}
