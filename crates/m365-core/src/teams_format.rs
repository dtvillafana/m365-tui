//! Safe Markdown-like formatting for outgoing Teams messages.
//!
//! The composer accepts a deliberately small Markdown subset and converts it
//! to HTML understood by Teams. Raw HTML is always escaped.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TeamsBody {
    pub content_type: &'static str,
    pub content: String,
}

impl TeamsBody {
    pub fn html_fragment(&self) -> String {
        if self.content_type == "html" {
            self.content.clone()
        } else {
            format!(
                "<p>{}</p>",
                escape_html(&self.content).replace('\n', "<br>")
            )
        }
    }
}

pub fn outgoing_body(text: &str, markdown: bool) -> TeamsBody {
    if !markdown {
        return TeamsBody {
            content_type: "text",
            content: text.to_string(),
        };
    }

    let (html, changed) = render_blocks(text);
    if changed {
        TeamsBody {
            content_type: "html",
            content: html,
        }
    } else {
        TeamsBody {
            content_type: "text",
            content: text.to_string(),
        }
    }
}

fn render_blocks(text: &str) -> (String, bool) {
    let lines: Vec<&str> = text.lines().collect();
    let mut out = String::new();
    let mut changed = false;
    let mut index = 0;

    while index < lines.len() {
        let line = lines[index];

        if line.trim_start().starts_with("```") {
            changed = true;
            index += 1;
            let mut code = String::new();
            while index < lines.len() && !lines[index].trim_start().starts_with("```") {
                if !code.is_empty() {
                    code.push('\n');
                }
                code.push_str(lines[index]);
                index += 1;
            }
            if index < lines.len() {
                index += 1;
            }
            out.push_str("<pre><code>");
            out.push_str(&escape_html(&code));
            out.push_str("</code></pre>");
            continue;
        }

        if line.starts_with("> ") || line == ">" {
            changed = true;
            out.push_str("<blockquote>");
            let mut first = true;
            while index < lines.len() && (lines[index].starts_with("> ") || lines[index] == ">") {
                if !first {
                    out.push_str("<br>");
                }
                first = false;
                let content = lines[index].strip_prefix("> ").unwrap_or("");
                let (rendered, inline_changed) = render_inline(content);
                changed |= inline_changed;
                out.push_str(&rendered);
                index += 1;
            }
            out.push_str("</blockquote>");
            continue;
        }

        if line.starts_with("- ") {
            changed = true;
            out.push_str("<ul>");
            while index < lines.len() && lines[index].starts_with("- ") {
                let (rendered, inline_changed) = render_inline(&lines[index][2..]);
                changed |= inline_changed;
                out.push_str("<li>");
                out.push_str(&rendered);
                out.push_str("</li>");
                index += 1;
            }
            out.push_str("</ul>");
            continue;
        }

        if line.is_empty() {
            index += 1;
            if !out.is_empty() {
                out.push_str("<br>");
            }
            continue;
        }

        let (rendered, inline_changed) = render_inline(line);
        changed |= inline_changed;
        out.push_str("<p>");
        out.push_str(&rendered);
        out.push_str("</p>");
        index += 1;
    }

    (out, changed)
}

fn render_inline(text: &str) -> (String, bool) {
    let mut out = String::new();
    let mut rest = text;
    let mut changed = false;

    while !rest.is_empty() {
        if rest.starts_with('\\') {
            let mut chars = rest.chars();
            chars.next();
            if let Some(next) = chars.next() {
                changed = true;
                push_escaped_char(&mut out, next);
                rest = &rest[1 + next.len_utf8()..];
                continue;
            }
        }

        if let Some(after) = rest.strip_prefix("**") {
            if let Some(end) = after.find("**") {
                changed = true;
                let (inner, _) = render_inline(&after[..end]);
                out.push_str("<strong>");
                out.push_str(&inner);
                out.push_str("</strong>");
                rest = &after[end + 2..];
                continue;
            }
        }

        if let Some(after) = rest.strip_prefix("~~") {
            if let Some(end) = after.find("~~") {
                changed = true;
                let (inner, _) = render_inline(&after[..end]);
                out.push_str("<s>");
                out.push_str(&inner);
                out.push_str("</s>");
                rest = &after[end + 2..];
                continue;
            }
        }

        if let Some(after) = rest.strip_prefix('`') {
            if let Some(end) = after.find('`') {
                changed = true;
                out.push_str("<code>");
                out.push_str(&escape_html(&after[..end]));
                out.push_str("</code>");
                rest = &after[end + 1..];
                continue;
            }
        }

        if let Some(after) = rest.strip_prefix('[') {
            if let Some(mid) = after.find("](") {
                let after_url = &after[mid + 2..];
                if let Some(end) = after_url.find(')') {
                    let label = &after[..mid];
                    let url = &after_url[..end];
                    if safe_link(url) {
                        changed = true;
                        let (label, _) = render_inline(label);
                        out.push_str("<a href=\"");
                        out.push_str(&escape_html_attr(url));
                        out.push_str("\">");
                        out.push_str(&label);
                        out.push_str("</a>");
                        rest = &after_url[end + 1..];
                        continue;
                    }
                }
            }
        }

        if let Some(after) = rest.strip_prefix('*') {
            if let Some(end) = after.find('*') {
                if end > 0 {
                    changed = true;
                    let (inner, _) = render_inline(&after[..end]);
                    out.push_str("<em>");
                    out.push_str(&inner);
                    out.push_str("</em>");
                    rest = &after[end + 1..];
                    continue;
                }
            }
        }

        let ch = rest.chars().next().expect("non-empty input");
        push_escaped_char(&mut out, ch);
        rest = &rest[ch.len_utf8()..];
    }

    (out, changed)
}

fn safe_link(url: &str) -> bool {
    let value = url.trim().to_ascii_lowercase();
    value.starts_with("https://") || value.starts_with("http://") || value.starts_with("mailto:")
}

fn push_escaped_char(out: &mut String, ch: char) {
    match ch {
        '&' => out.push_str("&amp;"),
        '<' => out.push_str("&lt;"),
        '>' => out.push_str("&gt;"),
        '"' => out.push_str("&quot;"),
        '\'' => out.push_str("&#39;"),
        _ => out.push(ch),
    }
}

fn escape_html(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        push_escaped_char(&mut out, ch);
    }
    out
}

fn escape_html_attr(text: &str) -> String {
    escape_html(text)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn plain_text_stays_plain() {
        assert_eq!(
            outgoing_body("hello <world>", true),
            TeamsBody {
                content_type: "text",
                content: "hello <world>".into(),
            }
        );
    }

    #[test]
    fn formats_supported_markdown_and_escapes_html() {
        let body = outgoing_body(
            "**bold** *italic* ~~strike~~ `code` [link](https://example.com) <b>raw</b>",
            true,
        );
        assert_eq!(body.content_type, "html");
        assert!(body.content.contains("<strong>bold</strong>"));
        assert!(body.content.contains("<em>italic</em>"));
        assert!(body.content.contains("<s>strike</s>"));
        assert!(body.content.contains("<code>code</code>"));
        assert!(body
            .content
            .contains("<a href=\"https://example.com\">link</a>"));
        assert!(body.content.contains("&lt;b&gt;raw&lt;/b&gt;"));
    }

    #[test]
    fn formats_blocks() {
        let body = outgoing_body("> quote\n- one\n- two\n```bash\necho <hello>\n```", true);
        assert_eq!(body.content_type, "html");
        assert!(body.content.contains("<blockquote>quote</blockquote>"));
        assert!(body.content.contains("<ul><li>one</li><li>two</li></ul>"));
        assert!(body
            .content
            .contains("<pre><code>echo &lt;hello&gt;</code></pre>"));
    }

    #[test]
    fn disabled_markdown_is_exact_plain_text() {
        let input = "**bold** <b>raw</b>";
        let body = outgoing_body(input, false);
        assert_eq!(body.content_type, "text");
        assert_eq!(body.content, input);
    }

    #[test]
    fn unsafe_link_is_not_emitted_as_html_link() {
        let body = outgoing_body("[x](javascript:alert(1))", true);
        assert_eq!(body.content_type, "text");
        assert_eq!(body.content, "[x](javascript:alert(1))");
    }
}
