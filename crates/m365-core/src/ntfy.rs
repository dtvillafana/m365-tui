//! ntfy notification forwarding.
//!
//! Publishes JSON to the configured ntfy server root. The topic is included in
//! the JSON body, so UTF-8 titles/messages do not depend on HTTP header encoding.

use anyhow::{Context, Result};

pub async fn send(
    server: &str,
    topic: &str,
    token: Option<&str>,
    tag: Option<&str>,
    priority: Option<u8>,
    title: &str,
    body: &str,
) -> Result<()> {
    let payload = message_payload(topic, tag, priority, title, body);

    let client = reqwest::Client::new();
    let mut request = client
        .post(server.trim_end_matches('/'))
        .timeout(std::time::Duration::from_secs(5))
        .json(&payload);

    if let Some(token) = token.map(str::trim).filter(|token| !token.is_empty()) {
        request = request.bearer_auth(token);
    }

    let response = request.send().await.with_context(|| {
        format!(
            "publishing ntfy message to {}",
            server.trim_end_matches('/')
        )
    })?;

    let status = response.status();
    anyhow::ensure!(
        status.is_success(),
        "ntfy server {} returned HTTP {status}",
        server.trim_end_matches('/')
    );

    Ok(())
}

fn message_payload(
    topic: &str,
    tag: Option<&str>,
    priority: Option<u8>,
    title: &str,
    body: &str,
) -> serde_json::Value {
    let mut payload = serde_json::json!({
        "topic": topic,
        "title": title,
        "message": summarise(body),
    });

    if let Some(tag) = tag.map(str::trim).filter(|tag| !tag.is_empty()) {
        payload["tags"] = serde_json::json!([tag]);
    }
    if let Some(priority) = priority {
        payload["priority"] = serde_json::json!(priority);
    }

    payload
}

fn summarise(body: &str) -> String {
    let flat: String = body.split_whitespace().collect::<Vec<_>>().join(" ");
    if flat.chars().count() <= 140 {
        return flat;
    }
    let head: String = flat.chars().take(137).collect();
    format!("{head}…")
}

#[cfg(test)]
mod tests {
    use super::*;
    use tokio::io::{AsyncReadExt, AsyncWriteExt};

    #[test]
    fn summary_is_bounded_and_single_line() {
        assert_eq!(summarise("hello\n\n  world   again "), "hello world again");
        let long = summarise(&"x".repeat(500));
        assert_eq!(long.chars().count(), 138);
    }

    #[test]
    fn payload_encodes_optional_ntfy_metadata() {
        let tagged = message_payload("m365", Some("email"), Some(4), "Subject", "Body");
        assert_eq!(tagged["tags"], serde_json::json!(["email"]));
        assert_eq!(tagged["priority"], serde_json::json!(4));

        let plain = message_payload("m365", None, None, "Subject", "Body");
        assert!(plain.get("tags").is_none());
        assert!(plain.get("priority").is_none());
    }

    #[tokio::test]
    async fn publishes_ntfy_json_to_server_root() {
        let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();
        let server = format!("http://{addr}");

        let receiver = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.unwrap();
            let mut request = Vec::new();
            let mut buf = [0_u8; 4096];
            let mut expected_len = None;

            loop {
                let n = stream.read(&mut buf).await.unwrap();
                assert!(n > 0, "connection closed before full request arrived");
                request.extend_from_slice(&buf[..n]);

                if expected_len.is_none() {
                    if let Some(header_end) = request.windows(4).position(|w| w == b"\r\n\r\n") {
                        let headers = String::from_utf8_lossy(&request[..header_end]);
                        let content_len = headers
                            .lines()
                            .find_map(|line| {
                                let (name, value) = line.split_once(':')?;
                                name.eq_ignore_ascii_case("content-length")
                                    .then(|| value.trim().parse::<usize>().ok())
                                    .flatten()
                            })
                            .expect("reqwest request must carry Content-Length");
                        expected_len = Some(header_end + 4 + content_len);
                    }
                }

                if expected_len.is_some_and(|len| request.len() >= len) {
                    break;
                }
            }

            let request_text = String::from_utf8(request).unwrap();
            assert!(
                request_text.starts_with("POST / HTTP/1.1\r\n"),
                "ntfy JSON publish must POST to server root"
            );
            assert!(
                request_text
                    .to_ascii_lowercase()
                    .contains("authorization: bearer secret-token\r\n"),
                "Bearer token missing from ntfy request"
            );

            let (_, body) = request_text.split_once("\r\n\r\n").unwrap();
            let json: serde_json::Value = serde_json::from_str(body).unwrap();
            assert_eq!(json["topic"], "m365");
            assert_eq!(json["title"], "Test title");
            assert_eq!(json["message"], "hello world");
            assert_eq!(json["tags"], serde_json::json!(["email"]));
            assert_eq!(json["priority"], serde_json::json!(4));

            stream
                .write_all(
                    b"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: 2\r\nConnection: close\r\n\r\n{}",
                )
                .await
                .unwrap();
        });

        send(
            &server,
            "m365",
            Some("secret-token"),
            Some("email"),
            Some(4),
            "Test title",
            "hello\nworld",
        )
        .await
        .unwrap();

        receiver.await.unwrap();
    }
}
