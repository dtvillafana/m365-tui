//! Exact word wrapping for styled lines.
//!
//! Ratatui's `Paragraph` can wrap for us, but it won't say how many rows that
//! produced — and scrolling needs that number. Estimating it (characters ÷
//! width) always undercounts, because words don't fill a row exactly, which
//! leaves the last messages hidden below the bottom edge.
//!
//! Wrapping here instead makes the row count exact by construction: the caller
//! renders the returned rows *without* `Wrap`, and scroll offsets are then just
//! indices into them.

use ratatui::text::{Line, Span};

/// Wrap one styled line into rows of at most `width` columns.
///
/// Breaks at spaces where possible, hard-splits words longer than the width, and
/// indents continuation rows to match the original's leading whitespace so
/// wrapped message text stays in its column.
pub fn wrap_line(line: &Line<'static>, width: usize) -> Vec<Line<'static>> {
    if width == 0 {
        return vec![line.clone()];
    }
    let total: usize = line.spans.iter().map(|s| s.content.chars().count()).sum();
    if total <= width {
        return vec![line.clone()];
    }

    // Hang continuation rows under the original indent, but never so far that
    // there's no room left to write in.
    let indent_len = leading_spaces(line).min(width.saturating_sub(8));
    let indent = " ".repeat(indent_len);

    let mut rows: Vec<Line<'static>> = Vec::new();
    let mut cur: Vec<Span<'static>> = Vec::new();
    let mut cur_w = 0usize;

    let mut newline = |cur: &mut Vec<Span<'static>>, cur_w: &mut usize| {
        rows.push(Line::from(std::mem::take(cur)));
        if indent_len > 0 {
            cur.push(Span::raw(indent.clone()));
        }
        *cur_w = indent_len;
    };

    for span in &line.spans {
        for chunk in split_keep_spaces(&span.content) {
            let is_space = chunk.chars().all(char::is_whitespace);
            let mut rest: Vec<char> = chunk.chars().collect();

            loop {
                let avail = width.saturating_sub(cur_w);
                if rest.len() <= avail {
                    if !rest.is_empty() {
                        cur_w += rest.len();
                        cur.push(Span::styled(rest.iter().collect::<String>(), span.style));
                    }
                    break;
                }
                // Doesn't fit. A space at a break point is dropped rather than
                // carried to the start of the next row.
                if is_space {
                    newline(&mut cur, &mut cur_w);
                    break;
                }
                // A word that would fit on a fresh row moves there whole.
                if rest.len() <= width.saturating_sub(indent_len) && cur_w > indent_len {
                    newline(&mut cur, &mut cur_w);
                    continue;
                }
                // Otherwise it's longer than a row: hard-split it.
                if avail == 0 {
                    newline(&mut cur, &mut cur_w);
                    continue;
                }
                let head: String = rest.drain(..avail).collect();
                cur.push(Span::styled(head, span.style));
                newline(&mut cur, &mut cur_w);
            }
        }
    }
    if !cur.is_empty() {
        rows.push(Line::from(cur));
    }
    if rows.is_empty() {
        rows.push(Line::from(""));
    }
    rows
}

/// Wrap every line, reporting where each original line begins in the output.
pub fn wrap_all(lines: &[Line<'static>], width: usize) -> (Vec<Line<'static>>, Vec<usize>) {
    let mut rows = Vec::with_capacity(lines.len());
    let mut starts = Vec::with_capacity(lines.len());
    for line in lines {
        starts.push(rows.len());
        rows.extend(wrap_line(line, width));
    }
    (rows, starts)
}

fn leading_spaces(line: &Line) -> usize {
    let mut n = 0;
    for span in &line.spans {
        for c in span.content.chars() {
            if c == ' ' {
                n += 1;
            } else {
                return n;
            }
        }
    }
    n
}

/// Split into alternating runs of whitespace and non-whitespace, keeping both.
fn split_keep_spaces(s: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut chars = s.char_indices().peekable();
    while let Some((i, c)) = chars.next() {
        let in_space = c.is_whitespace();
        let end = loop {
            match chars.peek() {
                Some(&(j, d)) if d.is_whitespace() == in_space => {
                    chars.next();
                    let _ = j;
                }
                Some(&(j, _)) => break j,
                None => break s.len(),
            }
        };
        out.push(&s[start.max(i)..end]);
        start = end;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use ratatui::style::{Color, Style};

    fn plain(rows: &[Line]) -> Vec<String> {
        rows.iter()
            .map(|r| r.spans.iter().map(|s| s.content.as_ref()).collect())
            .collect()
    }

    #[test]
    fn short_lines_are_untouched() {
        let line = Line::from("hello");
        assert_eq!(wrap_line(&line, 20).len(), 1);
    }

    #[test]
    fn breaks_at_spaces() {
        let line = Line::from("the quick brown fox jumps");
        let rows = wrap_line(&line, 10);
        for r in plain(&rows) {
            assert!(r.chars().count() <= 10, "row too wide: {r:?}");
        }
        assert_eq!(
            plain(&rows).join(" ").replace("  ", " ").trim(),
            "the quick brown fox jumps"
        );
    }

    #[test]
    fn hard_splits_a_word_longer_than_the_width() {
        let line = Line::from("aaaaaaaaaaaaaaaaaaaa");
        let rows = wrap_line(&line, 6);
        assert!(rows.len() >= 4);
        for r in plain(&rows) {
            assert!(r.chars().count() <= 6, "row too wide: {r:?}");
        }
        assert_eq!(plain(&rows).concat(), "aaaaaaaaaaaaaaaaaaaa");
    }

    #[test]
    fn continuation_rows_keep_the_indent() {
        // Message bodies arrive already indented under their timestamp gutter.
        let line = Line::from("        some fairly long message text here");
        let rows = wrap_line(&line, 20);
        assert!(rows.len() > 1);
        for r in plain(&rows).iter().skip(1) {
            assert!(r.starts_with("        "), "lost the indent: {r:?}");
        }
    }

    #[test]
    fn styles_survive_wrapping() {
        let line = Line::from(vec![
            Span::styled("aaaa ", Style::default().fg(Color::Red)),
            Span::styled("bbbb cccc", Style::default().fg(Color::Green)),
        ]);
        let rows = wrap_line(&line, 6);
        let colours: Vec<Option<Color>> = rows
            .iter()
            .flat_map(|r| r.spans.iter().map(|s| s.style.fg))
            .collect();
        assert!(colours.contains(&Some(Color::Red)));
        assert!(colours.contains(&Some(Color::Green)));
    }

    #[test]
    fn every_row_fits_the_width() {
        // The property that matters: nothing exceeds the width, so a caller can
        // trust the row count for scrolling.
        let line = Line::from(
            "Não, não implicam restart do serviço, mas de qualquer maneira, acho \
             que os unattended upgrades trataram disso, pelo menos em 4 das máquinas",
        );
        for width in [8, 20, 40, 80] {
            for row in plain(&wrap_line(&line, width)) {
                assert!(
                    row.chars().count() <= width,
                    "width {width}: row {row:?} too wide"
                );
            }
        }
    }

    #[test]
    fn wrap_all_reports_where_each_line_starts() {
        let lines = vec![
            Line::from("one"),
            Line::from("a much longer line that will certainly wrap"),
            Line::from("three"),
        ];
        let (rows, starts) = wrap_all(&lines, 10);
        assert_eq!(starts[0], 0);
        assert_eq!(starts[1], 1);
        assert!(starts[2] > 2, "the middle line occupies several rows");
        assert_eq!(rows.len(), starts[2] + 1);
    }
}
