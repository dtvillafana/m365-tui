//! A small text-editing buffer with the movement and deletion keys people
//! expect when writing an email: arrows, Home/End, word jumps, word/line
//! deletion, and multi-line editing with soft wrapping.
//!
//! Positions are tracked as **character** indices (not bytes) so accented text
//! behaves correctly.

/// One display row produced by soft-wrapping, tagged with the character index
/// in the full text where it starts. The renderer and the cursor mapper share
/// this so they can never disagree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Row {
    pub start: usize,
    pub text: String,
}

#[derive(Debug, Default, Clone)]
pub struct TextInput {
    chars: Vec<char>,
    /// Cursor position as a character index in `0..=chars.len()`.
    cursor: usize,
}

impl TextInput {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn text(&self) -> String {
        self.chars.iter().collect()
    }

    pub fn is_empty(&self) -> bool {
        self.chars.is_empty()
    }

    pub fn cursor(&self) -> usize {
        self.cursor
    }

    pub fn chars(&self) -> &[char] {
        &self.chars
    }

    /// Replace `start..end` (character indices) and leave the cursor after the
    /// inserted text.
    pub fn replace_range(&mut self, start: usize, end: usize, s: &str) {
        let start = start.min(self.chars.len());
        let end = end.min(self.chars.len()).max(start);
        self.chars.drain(start..end);
        for (i, c) in s.chars().enumerate() {
            self.chars.insert(start + i, c);
        }
        self.cursor = start + s.chars().count();
    }

    pub fn clear(&mut self) {
        self.chars.clear();
        self.cursor = 0;
    }

    // -- editing -----------------------------------------------------------

    pub fn insert(&mut self, c: char) {
        self.chars.insert(self.cursor, c);
        self.cursor += 1;
    }

    /// Insert pasted text. Normalises CRLF so pastes don't leave stray \r.
    pub fn insert_str(&mut self, s: &str) {
        for c in s.chars().filter(|&c| c != '\r') {
            self.insert(c);
        }
    }

    /// Delete the character before the cursor.
    pub fn backspace(&mut self) {
        if self.cursor > 0 {
            self.cursor -= 1;
            self.chars.remove(self.cursor);
        }
    }

    /// Delete the character under the cursor.
    pub fn delete(&mut self) {
        if self.cursor < self.chars.len() {
            self.chars.remove(self.cursor);
        }
    }

    /// Ctrl+W: delete the word before the cursor (trailing spaces included).
    pub fn delete_word_before(&mut self) {
        let start = self.word_start();
        self.chars.drain(start..self.cursor);
        self.cursor = start;
    }

    /// Ctrl+U: delete from the start of the current line to the cursor.
    pub fn delete_to_line_start(&mut self) {
        let start = self.line_start();
        self.chars.drain(start..self.cursor);
        self.cursor = start;
    }

    /// Ctrl+K: delete from the cursor to the end of the current line.
    pub fn delete_to_line_end(&mut self) {
        let end = self.line_end();
        // On an empty line, swallow the newline itself.
        let end = if end == self.cursor && end < self.chars.len() {
            end + 1
        } else {
            end
        };
        self.chars.drain(self.cursor..end);
    }

    // -- movement ----------------------------------------------------------

    pub fn left(&mut self) {
        self.cursor = self.cursor.saturating_sub(1);
    }

    pub fn right(&mut self) {
        if self.cursor < self.chars.len() {
            self.cursor += 1;
        }
    }

    pub fn word_left(&mut self) {
        self.cursor = self.word_start();
    }

    pub fn word_right(&mut self) {
        let mut i = self.cursor;
        while i < self.chars.len() && self.chars[i].is_whitespace() {
            i += 1;
        }
        while i < self.chars.len() && !self.chars[i].is_whitespace() {
            i += 1;
        }
        self.cursor = i;
    }

    pub fn home(&mut self) {
        self.cursor = self.line_start();
    }

    pub fn end(&mut self) {
        self.cursor = self.line_end();
    }

    pub fn start_of_text(&mut self) {
        self.cursor = 0;
    }

    pub fn end_of_text(&mut self) {
        self.cursor = self.chars.len();
    }

    /// Move up/down one *display* row, keeping the column where possible.
    pub fn move_row(&mut self, delta: isize, width: usize) {
        let rows = self.wrap(width);
        let (row, col) = cursor_rowcol(&rows, self.cursor);
        let target = row as isize + delta;
        if target < 0 || target as usize >= rows.len() {
            return;
        }
        let target = &rows[target as usize];
        let len = target.text.chars().count();
        self.cursor = target.start + col.min(len);
    }

    // -- wrapping ----------------------------------------------------------

    /// Soft-wrap the text into display rows of at most `width` columns.
    pub fn wrap(&self, width: usize) -> Vec<Row> {
        wrap_chars(&self.chars, width.max(1))
    }

    /// The cursor's (row, column) among the wrapped display rows.
    pub fn cursor_position(&self, width: usize) -> (usize, usize) {
        cursor_rowcol(&self.wrap(width), self.cursor)
    }

    // -- internals ---------------------------------------------------------

    fn line_start(&self) -> usize {
        self.chars[..self.cursor]
            .iter()
            .rposition(|&c| c == '\n')
            .map(|i| i + 1)
            .unwrap_or(0)
    }

    fn line_end(&self) -> usize {
        self.chars[self.cursor..]
            .iter()
            .position(|&c| c == '\n')
            .map(|i| self.cursor + i)
            .unwrap_or(self.chars.len())
    }

    fn word_start(&self) -> usize {
        let mut i = self.cursor;
        while i > 0 && self.chars[i - 1].is_whitespace() {
            i -= 1;
        }
        while i > 0 && !self.chars[i - 1].is_whitespace() {
            i -= 1;
        }
        i
    }
}

impl From<&str> for TextInput {
    fn from(s: &str) -> Self {
        let chars: Vec<char> = s.chars().collect();
        let cursor = chars.len();
        Self { chars, cursor }
    }
}

fn wrap_chars(chars: &[char], width: usize) -> Vec<Row> {
    let mut rows = Vec::new();
    let mut i = 0;
    loop {
        let mut j = i;
        while j < chars.len() && chars[j] != '\n' {
            j += 1;
        }
        wrap_segment(&chars[i..j], i, width, &mut rows);
        if j >= chars.len() {
            break;
        }
        i = j + 1;
        // A trailing newline leaves an empty final row for the cursor to sit on.
        if i == chars.len() {
            rows.push(Row {
                start: i,
                text: String::new(),
            });
            break;
        }
    }
    if rows.is_empty() {
        rows.push(Row {
            start: 0,
            text: String::new(),
        });
    }
    rows
}

/// Greedy word wrap of a single logical line, breaking at spaces where possible
/// and hard-splitting words longer than `width`.
fn wrap_segment(seg: &[char], offset: usize, width: usize, rows: &mut Vec<Row>) {
    if seg.is_empty() {
        rows.push(Row {
            start: offset,
            text: String::new(),
        });
        return;
    }
    let mut pos = 0;
    while pos < seg.len() {
        if seg.len() - pos <= width {
            rows.push(Row {
                start: offset + pos,
                text: seg[pos..].iter().collect(),
            });
            return;
        }
        // Prefer breaking at the last space that fits.
        let limit = pos + width;
        let brk = (pos..limit).rev().find(|&k| seg[k] == ' ');
        match brk {
            Some(k) if k > pos => {
                rows.push(Row {
                    start: offset + pos,
                    text: seg[pos..k].iter().collect(),
                });
                pos = k + 1; // the break space is consumed by the wrap
            }
            _ => {
                rows.push(Row {
                    start: offset + pos,
                    text: seg[pos..limit].iter().collect(),
                });
                pos = limit;
            }
        }
    }
}

fn cursor_rowcol(rows: &[Row], cursor: usize) -> (usize, usize) {
    // The cursor belongs to the last row that starts at or before it.
    let idx = rows.iter().rposition(|r| r.start <= cursor).unwrap_or(0);
    let row = &rows[idx];
    let len = row.text.chars().count();
    (idx, (cursor.saturating_sub(row.start)).min(len))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inserts_and_deletes_at_the_cursor() {
        let mut t = TextInput::from("helo");
        t.left(); // between 'l' and 'o'
        t.insert('l');
        assert_eq!(t.text(), "hello");
        t.end();
        t.backspace();
        assert_eq!(t.text(), "hell");
    }

    #[test]
    fn handles_multibyte_characters() {
        // Portuguese accents must not corrupt the buffer or the cursor.
        let mut t = TextInput::from("ação");
        assert_eq!(t.cursor(), 4);
        t.backspace();
        assert_eq!(t.text(), "açã");
        t.start_of_text();
        t.right();
        t.insert('X');
        assert_eq!(t.text(), "aXçã");
    }

    #[test]
    fn word_and_line_deletion() {
        let mut t = TextInput::from("hello brave world");
        t.delete_word_before();
        assert_eq!(t.text(), "hello brave ");

        let mut t = TextInput::from("keep this");
        t.delete_to_line_start();
        assert_eq!(t.text(), "");

        let mut t = TextInput::from("one\ntwo");
        t.home();
        t.delete_to_line_end();
        assert_eq!(t.text(), "one\n");
    }

    #[test]
    fn home_and_end_are_per_line() {
        let mut t = TextInput::from("first\nsecond");
        t.home();
        assert_eq!(t.cursor(), 6, "start of the second line");
        t.end();
        assert_eq!(t.cursor(), 12);
    }

    #[test]
    fn word_movement_skips_whitespace() {
        let mut t = TextInput::from("alpha beta");
        t.start_of_text();
        t.word_right();
        assert_eq!(t.cursor(), 5);
        t.word_right();
        assert_eq!(t.cursor(), 10);
        t.word_left();
        assert_eq!(t.cursor(), 6);
    }

    #[test]
    fn wraps_on_word_boundaries() {
        let t = TextInput::from("the quick brown fox");
        let rows = t.wrap(10);
        assert_eq!(rows[0].text, "the quick");
        assert_eq!(rows[1].text, "brown fox");
    }

    #[test]
    fn hard_splits_words_longer_than_the_width() {
        let t = TextInput::from("aaaaaaaaaaaaa");
        let rows = t.wrap(5);
        assert_eq!(rows[0].text, "aaaaa");
        assert_eq!(rows[1].text, "aaaaa");
        assert_eq!(rows[2].text, "aaa");
    }

    #[test]
    fn cursor_maps_to_the_right_row_and_column() {
        let mut t = TextInput::from("the quick brown fox");
        t.start_of_text();
        assert_eq!(t.cursor_position(10), (0, 0));
        t.end_of_text();
        let (row, col) = t.cursor_position(10);
        assert_eq!((row, col), (1, 9), "end of 'brown fox'");
    }

    #[test]
    fn explicit_newlines_start_new_rows() {
        let t = TextInput::from("a\n\nb");
        let rows = t.wrap(20);
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[1].text, "");
        assert_eq!(rows[2].text, "b");
    }

    #[test]
    fn trailing_newline_leaves_a_row_to_type_on() {
        let t = TextInput::from("a\n");
        let rows = t.wrap(20);
        assert_eq!(rows.len(), 2);
        assert_eq!(t.cursor_position(20), (1, 0));
    }

    #[test]
    fn moves_between_display_rows_keeping_the_column() {
        let mut t = TextInput::from("the quick brown fox");
        t.end_of_text();
        t.move_row(-1, 10); // up into "the quick"
        let (row, _) = t.cursor_position(10);
        assert_eq!(row, 0);
        t.move_row(1, 10);
        assert_eq!(t.cursor_position(10).0, 1);
        t.move_row(5, 10); // past the end: no movement
        assert_eq!(t.cursor_position(10).0, 1);
    }

    #[test]
    fn paste_strips_carriage_returns() {
        let mut t = TextInput::new();
        t.insert_str("line one\r\nline two");
        assert_eq!(t.text(), "line one\nline two");
    }

    #[test]
    fn replace_range_updates_the_cursor() {
        let mut t = TextInput::from("see @./sh please");
        t.replace_range(4, 9, "@./shot.png");
        assert_eq!(t.text(), "see @./shot.png please");
        assert_eq!(t.cursor(), 4 + "@./shot.png".chars().count());
    }
}
