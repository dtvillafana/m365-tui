//! Small shared helpers.

/// Standard base64 with padding — Graph wants attachment bytes this way.
pub fn base64_encode(input: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(input.len().div_ceil(3) * 4);
    for chunk in input.chunks(3) {
        let b = [
            chunk[0],
            chunk.get(1).copied().unwrap_or(0),
            chunk.get(2).copied().unwrap_or(0),
        ];
        let n = ((b[0] as u32) << 16) | ((b[1] as u32) << 8) | b[2] as u32;
        out.push(ALPHABET[(n >> 18 & 63) as usize] as char);
        out.push(ALPHABET[(n >> 12 & 63) as usize] as char);
        out.push(if chunk.len() > 1 {
            ALPHABET[(n >> 6 & 63) as usize] as char
        } else {
            '='
        });
        out.push(if chunk.len() > 2 {
            ALPHABET[(n & 63) as usize] as char
        } else {
            '='
        });
    }
    out
}

/// Decode standard padded base64. Whitespace is ignored.
pub fn base64_decode(input: &str) -> Result<Vec<u8>, &'static str> {
    fn val(b: u8) -> Result<u8, &'static str> {
        match b {
            b'A'..=b'Z' => Ok(b - b'A'),
            b'a'..=b'z' => Ok(b - b'a' + 26),
            b'0'..=b'9' => Ok(b - b'0' + 52),
            b'+' => Ok(62),
            b'/' => Ok(63),
            _ => Err("invalid base64"),
        }
    }

    let filtered: Vec<u8> = input.bytes().filter(|b| !b.is_ascii_whitespace()).collect();
    if filtered.is_empty() {
        return Ok(Vec::new());
    }
    if !filtered.len().is_multiple_of(4) {
        return Err("invalid base64 length");
    }
    let mut out = Vec::with_capacity(filtered.len() / 4 * 3);
    for chunk in filtered.chunks(4) {
        let pad = chunk.iter().rev().take_while(|&&c| c == b'=').count();
        if pad > 2 || chunk[..4 - pad].contains(&b'=') {
            return Err("invalid base64 padding");
        }
        let v0 = val(chunk[0])? as u32;
        let v1 = val(chunk[1])? as u32;
        let v2 = if pad == 2 { 0 } else { val(chunk[2])? as u32 };
        let v3 = if pad >= 1 { 0 } else { val(chunk[3])? as u32 };
        let n = (v0 << 18) | (v1 << 12) | (v2 << 6) | v3;
        out.push((n >> 16) as u8);
        if pad < 2 {
            out.push((n >> 8) as u8);
        }
        if pad < 1 {
            out.push(n as u8);
        }
    }
    Ok(out)
}

/// Escape text for embedding in an HTML mail body.
pub fn html_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            '\n' => out.push_str("<br>"),
            _ => out.push(c),
        }
    }
    out
}

/// Escape text for an HTML attribute value (`alt`, etc.). Newlines stay as
/// spaces so they cannot inject a `<br>`.
pub fn html_escape_attr(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            '\n' | '\r' => out.push(' '),
            _ => out.push(c),
        }
    }
    out
}

/// Pull the human-readable sentence out of a Graph error blob.
///
/// Graph returns a wall of JSON; the status bar wants the one sentence that
/// says what went wrong.
pub fn graph_error_summary(err: &str) -> String {
    if let Some(start) = err.find("\"message\":\"") {
        let rest = &err[start + 11..];
        if let Some(end) = rest.find('"') {
            return rest[..end].to_string();
        }
    }
    let flat: String = err.split_whitespace().collect::<Vec<_>>().join(" ");
    flat.chars().take(160).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn base64_matches_known_vectors() {
        assert_eq!(base64_encode(b""), "");
        assert_eq!(base64_encode(b"f"), "Zg==");
        assert_eq!(base64_encode(b"fo"), "Zm8=");
        assert_eq!(base64_encode(b"foo"), "Zm9v");
        assert_eq!(base64_encode(b"foob"), "Zm9vYg==");
        assert_eq!(base64_encode(b"fooba"), "Zm9vYmE=");
        assert_eq!(base64_encode(b"foobar"), "Zm9vYmFy");
        for input in [
            b"".as_slice(),
            b"f",
            b"fo",
            b"foo",
            b"foob",
            b"fooba",
            b"foobar",
        ] {
            assert_eq!(base64_decode(&base64_encode(input)).unwrap(), input);
        }
        assert_eq!(base64_decode(" Zm9 v\n").unwrap(), b"foo");
        assert!(base64_decode("!!!").is_err());
    }

    #[test]
    fn summarises_graph_errors() {
        let raw = r#"Graph request failed (403 Forbidden): {"error":{"code":"Forbidden","message":"Missing scope permissions on the request.","innerError":{"date":"2026-08-04"}}}"#;
        assert_eq!(
            graph_error_summary(raw),
            "Missing scope permissions on the request."
        );
        // Non-Graph errors are flattened and bounded.
        assert_eq!(
            graph_error_summary("connection refused"),
            "connection refused"
        );
        assert_eq!(graph_error_summary(&"x".repeat(500)).chars().count(), 160);
    }

    #[test]
    fn escapes_html_and_newlines() {
        assert_eq!(
            html_escape("a <b> & \"c\"\nd"),
            "a &lt;b&gt; &amp; &quot;c&quot;<br>d"
        );
        assert_eq!(html_escape_attr("a \"b\"\nc"), "a &quot;b&quot; c");
    }
}
