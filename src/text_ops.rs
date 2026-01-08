use std::collections::HashSet;

/// Detects the end-of-line style: "\r\n" or "\n".
/// Defaults to "\n" if mixed or none found.
fn detect_eol(s: &str) -> &str {
    if s.contains("\r\n") { "\r\n" } else { "\n" }
}

/// Helper to split text into content and optional trailing newline.
/// Returns (content_without_trailing, has_trailing_newline)
fn split_trailing_newline(text: &str) -> (&str, bool) {
    if text.ends_with("\r\n") {
        return (&text[..text.len() - 2], true);
    }
    if text.ends_with('\n') {
        return (&text[..text.len() - 1], true);
    }
    (text, false)
}

/// Splits a string by EOLs, preserving empty lines.
/// Handles \r\n and \n.
fn split_lines_keep_empty(text: &str) -> Vec<&str> {
    // split_inclusive might be better, but we want to strip the EOLs
    // to compare content, then re-join with the detected EOL.
    // However, we need to handle mixed line endings gracefully if possible,
    // though the spec says "Normalize only line endings for comparison".

    // Simple approach: split by \n, then strip \r.
    // This handles both \r\n and \n.
    text.split('\n')
        .map(|line| line.strip_suffix('\r').unwrap_or(line))
        .collect()
}

pub fn remove_duplicate_lines(scope: &str) -> String {
    if scope.is_empty() {
        return String::new();
    }

    let eol = detect_eol(scope);
    let (content, trailing) = split_trailing_newline(scope);

    // If splitting by '\n' results in an empty string at the end (e.g. "a\n"),
    // split() returns ["a", ""].
    // If scope is "a\nb", split is ["a", "b"].
    // If scope is "a\n", content is "a", split is ["a"].
    // If scope is "a\n\n", content is "a\n", split is ["a", ""].

    // We work on `content`.
    let mut lines = split_lines_keep_empty(content);

    // Special case: if the original string ended with a newline, `split_trailing_newline` removed it.
    // If `content` was empty (e.g. scope was just "\n"), then lines is [""] which is correct for a blank line?
    // references:
    // "a" -> content="a", trailing=false. split=["a"]
    // "a\n" -> content="a", trailing=true. split=["a"]
    // "\n" -> content="", trailing=true. split=[""] (one empty line)

    let mut seen = HashSet::new();
    let mut out_lines = Vec::new();

    for line in lines {
        // Compare lines by exact text content (do NOT trim spaces, do NOT ignore case).
        if seen.insert(line) {
            out_lines.push(line);
        }
    }

    let mut out = out_lines.join(eol);
    if trailing {
        out.push_str(eol);
    }
    out
}

pub fn remove_duplicate_consecutive_lines(scope: &str) -> String {
    if scope.is_empty() {
        return String::new();
    }

    let eol = detect_eol(scope);
    let (content, trailing) = split_trailing_newline(scope);

    let mut lines = split_lines_keep_empty(content);
    let mut out_lines = Vec::new();

    let mut last_line: Option<&str> = None;

    for line in lines {
        match last_line {
            Some(last) if last == line => {
                // Duplicate, skip
            }
            _ => {
                out_lines.push(line);
                last_line = Some(line);
            }
        }
    }

    let mut out = out_lines.join(eol);
    if trailing {
        out.push_str(eol);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_empty() {
        assert_eq!(remove_duplicate_lines(""), "");
        assert_eq!(remove_duplicate_consecutive_lines(""), "");
    }

    #[test]
    fn test_single_line_no_newline() {
        assert_eq!(remove_duplicate_lines("abc"), "abc");
        assert_eq!(remove_duplicate_consecutive_lines("abc"), "abc");
    }

    #[test]
    fn test_single_line_with_newline() {
        assert_eq!(remove_duplicate_lines("abc\n"), "abc\n");
        assert_eq!(remove_duplicate_consecutive_lines("abc\n"), "abc\n");
    }

    #[test]
    fn test_duplicates_non_consecutive() {
        let input = "a\nb\na\n";
        // Global removes second 'a'
        assert_eq!(remove_duplicate_lines(input), "a\nb\n");
        // Consecutive keeps both 'a's as they are not adjacent
        assert_eq!(remove_duplicate_consecutive_lines(input), "a\nb\na\n");
    }

    #[test]
    fn test_duplicates_consecutive() {
        let input = "a\na\nb\n";
        assert_eq!(remove_duplicate_lines(input), "a\nb\n");
        assert_eq!(remove_duplicate_consecutive_lines(input), "a\nb\n");
    }

    #[test]
    fn test_whitespace_significant() {
        let input = "a\na \n";
        assert_eq!(remove_duplicate_lines(input), "a\na \n");
        assert_eq!(remove_duplicate_consecutive_lines(input), "a\na \n");
    }

    #[test]
    fn test_crlf_preserved() {
        let input = "a\r\na\r\nb";
        assert_eq!(remove_duplicate_lines(input), "a\r\nb");
        assert_eq!(remove_duplicate_consecutive_lines(input), "a\r\nb");
    }

    #[test]
    fn test_mixed_eol_uses_first_found() {
        // First EOL is \r\n, so output uses \r\n
        let input = "a\r\nb\nc";
        // split: "a", "b", "c". Join with \r\n.
        assert_eq!(remove_duplicate_lines(input), "a\r\nb\r\nc");
    }

    #[test]
    fn test_trailing_newline_preserved() {
        assert_eq!(remove_duplicate_lines("a\na\n"), "a\n");
        assert_eq!(remove_duplicate_lines("a\na"), "a");
    }

    #[test]
    fn test_empty_lines_deduplication() {
        // "\n\n" -> two empty lines.
        // Global: keep one empty line.
        // Consecutive: keep one empty line.
        assert_eq!(remove_duplicate_lines("\n\n"), "\n");
        assert_eq!(remove_duplicate_consecutive_lines("\n\n"), "\n");

        // "\n\na\n\n"
        // Global: "" (first), "a", "" (duplicate of first) -> removed.
        // Result: "\na" (one empty line, then 'a').
        // Wait, join("\n") of ["", "a"] is "\n" + "\n" + "a" ? No.
        // Join ["", "a"] with "\n" -> "\na".
        // Original: "" (line 1), "" (line 2), "a" (line 3), "" (line 4 -- trailing newline for line 3? No, line 4 is empty).
        // "\n\na\n\n" -> content "\n\na\n". trailing=true.
        // lines: ["", "", "a", ""].
        // Global: seen "", insert. seen "", skip. seen "a", insert. seen "", skip.
        // Out: ["", "a"]. Join("\n") -> "\na". Trailing=true -> "\na\n".
        assert_eq!(remove_duplicate_lines("\n\na\n\n"), "\na\n");
    }
}
