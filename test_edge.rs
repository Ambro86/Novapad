fn main() {
    let text = "Ciao <pause ms=\"500\"/> dopo";
    // I need to use the functions from tts_engine but they are private or need the crate.
    // Let me just copy the logic.
    let cleaned = sanitize_edge_text(text);
    let escaped = escape_xml(&cleaned);
    println!("escaped: {}", escaped);
    let rendered = render_edge_ssml_text_with_pause_tags(&escaped);
    println!("rendered: {}", rendered);
}

fn sanitize_edge_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        let code = ch as u32;
        if (0..=8).contains(&code) || (11..=12).contains(&code) || (14..=31).contains(&code) {
            out.push(' ');
        } else {
            out.push(ch);
        }
    }
    out
}

fn escape_xml(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

fn parse_pause_tag_milliseconds(tag: &str) -> Option<u32> {
    let trimmed = tag.trim();
    let inner = trimmed
        .strip_prefix('<')?
        .strip_suffix('>')?
        .trim()
        .trim_end_matches('/')
        .trim();
    let rest = inner.strip_prefix("pause")?.trim();
    if rest.is_empty() {
        return None;
    }
    for token in rest.split_whitespace() {
        let value = token
            .strip_prefix("ms=")
            .or_else(|| token.strip_prefix("milliseconds="))
            .unwrap_or(token)
            .trim_matches(['"', '\'']);
        if let Ok(ms) = value.parse::<u32>()
            && (50..=60_000).contains(&ms)
        {
            return Some(ms);
        }
    }
    None
}

fn decode_basic_xml_entities(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn render_edge_ssml_text_with_pause_tags(text: &str) -> String {
    let lower = text.to_ascii_lowercase();
    let mut out = String::with_capacity(text.len());
    let mut cursor = 0usize;
    let mut i = 0usize;
    let bytes = lower.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'<' {
            let remaining = &lower[i..];
            if remaining.starts_with("<pause")
                && let Some(end_rel) = remaining.find('>')
            {
                let end = i + end_rel + 1;
                if let Some(ms) = parse_pause_tag_milliseconds(&lower[i..end]) {
                    out.push_str(&text[cursor..i]);
                    out.push_str(&format!("<break time=\"{}ms\"/>", ms));
                    cursor = end;
                    i = end;
                    continue;
                }
            }
        } else if bytes[i] == b'&' {
            let remaining = &lower[i..];
            if remaining.starts_with("&lt;pause")
                && let Some(end_rel) = remaining.find("&gt;")
            {
                let end = i + end_rel + "&gt;".len();
                let decoded = decode_basic_xml_entities(&lower[i..end]);
                if let Some(ms) = parse_pause_tag_milliseconds(&decoded) {
                    out.push_str(&text[cursor..i]);
                    out.push_str(&format!("<break time=\"{}ms\"/>", ms));
                    cursor = end;
                    i = end;
                    continue;
                }
            }
        }
        i += 1;
    }
    if cursor < text.len() {
        out.push_str(&text[cursor..]);
    }
    out
}
