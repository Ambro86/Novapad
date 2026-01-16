use scraper::{Html, Selector};

#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct ArticleContent {
    pub title: String,
    pub content: String,
    pub excerpt: String,
}

fn decode_unicode(input: &str) -> String {
    let mut result = String::new();
    let mut chars = input.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '\\' && chars.peek() == Some(&'u') {
            chars.next();
            let mut hex = String::new();
            for _ in 0..4 {
                if let Some(h) = chars.next() {
                    hex.push(h);
                }
            }
            if let Ok(code) = u32::from_str_radix(&hex, 16) {
                if let Some(decoded_char) = std::char::from_u32(code) {
                    result.push(decoded_char);
                    continue;
                }
            }
            result.push_str("\\u");
            result.push_str(&hex);
        } else {
            result.push(c);
        }
    }
    result
}

/// Estrae una stringa JSON gestendo correttamente gli escape (\" \\ \n ecc.)
/// Ritorna la stringa decodificata e la posizione dopo la virgoletta di chiusura
fn extract_json_string(s: &str) -> Option<(String, usize)> {
    let mut result = String::new();
    let mut chars = s.char_indices().peekable();

    while let Some((i, c)) = chars.next() {
        if c == '\\' {
            // Carattere escaped
            if let Some((_, next_c)) = chars.next() {
                match next_c {
                    '"' => result.push('"'),
                    '\\' => result.push('\\'),
                    'n' => result.push('\n'),
                    'r' => result.push('\r'),
                    't' => result.push('\t'),
                    'u' => {
                        // Unicode escape \uXXXX
                        let mut hex = String::new();
                        for _ in 0..4 {
                            if let Some((_, h)) = chars.next() {
                                hex.push(h);
                            }
                        }
                        if let Ok(code) = u32::from_str_radix(&hex, 16) {
                            if let Some(decoded_char) = std::char::from_u32(code) {
                                result.push(decoded_char);
                            }
                        }
                    }
                    _ => {
                        result.push('\\');
                        result.push(next_c);
                    }
                }
            }
        } else if c == '"' {
            // Fine della stringa
            return Some((result, i + 1));
        } else {
            result.push(c);
        }
    }

    // Stringa non chiusa
    if !result.is_empty() {
        Some((result, s.len()))
    } else {
        None
    }
}

pub fn clean_text(input: &str) -> String {
    let decoded = decode_unicode(input);
    // Pulizia encoding Mediaset/TGCOM24
    let mut text = decoded
        .replace("Ã¨", "è")
        .replace("Ã ", "à")
        .replace("Ã¹", "ù")
        .replace("Ã²", "ò")
        .replace("Ã¬", "ì")
        .replace("Â ", " ")
        .replace("Ã©", "é")
        .replace("Â", "");

    text = text
        .replace("&nbsp;", " ")
        .replace("&#160;", " ")
        .replace("\u{00a0}", " ");
    text = text
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&apos;", "'");
    text = text
        .replace("\\\"", "\"")
        .replace("\\n", "\n")
        .replace("\\/", "/");

    let mut cleaned = String::new();
    let mut in_tag = false;
    for c in text.chars() {
        if c == '<' {
            in_tag = true;
        } else if c == '>' {
            in_tag = false;
            cleaned.push(' ');
        } else if !in_tag {
            cleaned.push(c);
        }
    }
    cleaned
}

pub fn reader_mode_extract(html_content: &str) -> Option<ArticleContent> {
    let document = Html::parse_document(html_content);
    let title = pick_title(&document);

    let mut body_acc = String::new();
    let mut author_info = String::new();
    let mut found_anything = false;

    // 1. ESTRAZIONE DA JSON-LD (Schema.org) - MOLTO RICCO SU TGCOM24
    if let Ok(s) = Selector::parse("script[type='application/ld+json']") {
        for element in document.select(&s) {
            let json = element.text().collect::<Vec<_>>().join("");

            // Cerchiamo Autore e Data
            if author_info.is_empty() {
                if let Some(a_idx) = json.find("\"name\":\"") {
                    let part = &json[a_idx + 8..];
                    if let Some((name, _)) = extract_json_string(part) {
                        author_info.push_str(&name);
                    }
                }
                if let Some(d_idx) = json.find("\"datePublished\":\"") {
                    let part = &json[d_idx + 17..];
                    if let Some((date_str, _)) = extract_json_string(part) {
                        let date = if date_str.len() >= 10 {
                            &date_str[..10]
                        } else {
                            &date_str
                        };
                        author_info.push_str(&format!(" ({})", date));
                    }
                }
            }

            // Cerchiamo description e articleBody
            for key in [
                "\"description\":\"",
                "\"articleBody\":\"",
                "\"subtitle\":\"",
            ] {
                let mut search_pos = 0;
                while let Some(key_pos) = json[search_pos..].find(key) {
                    let abs_start = search_pos + key_pos + key.len();
                    if abs_start < json.len() {
                        if let Some((val, end_pos)) = extract_json_string(&json[abs_start..]) {
                            if val.len() > 40 && !val.contains("http") && !body_acc.contains(&val) {
                                body_acc.push_str(&val);
                                body_acc.push_str("\n\n");
                                found_anything = true;
                            }
                            search_pos = abs_start + end_pos;
                        } else {
                            break;
                        }
                    } else {
                        break;
                    }
                }
            }
        }
    }

    // 2. ESTRAZIONE DA NEXT_DATA (WSJ / Altri)
    if !found_anything {
        if let Ok(next_selector) = Selector::parse("script#__NEXT_DATA__") {
            if let Some(element) = document.select(&next_selector).next() {
                let json_text = element.text().collect::<Vec<_>>().join("");

                // WSJ: estrai testi dai blocchi "content":[...] dei paragrafi
                let mut seen_paragraphs = std::collections::HashSet::new();
                for content_block in json_text.split("\"type\":\"paragraph\"") {
                    if let Some(content_start) = content_block.find("\"content\":[") {
                        let after_content = &content_block[content_start..];
                        // Estrai tutti i "text":"..." dal blocco content usando extract_json_string
                        let mut para_text = String::new();
                        let mut search_pos = 0;
                        while let Some(text_start) = after_content[search_pos..].find("\"text\":\"")
                        {
                            let abs_start = search_pos + text_start + 8; // dopo "text":"
                            if abs_start < after_content.len() {
                                if let Some((val, end_pos)) =
                                    extract_json_string(&after_content[abs_start..])
                                {
                                    if !val.is_empty() && !val.starts_with('{') {
                                        para_text.push_str(&val);
                                    }
                                    search_pos = abs_start + end_pos;
                                } else {
                                    break;
                                }
                            } else {
                                break;
                            }
                        }
                        // Evita duplicati
                        if para_text.len() > 20 && !seen_paragraphs.contains(&para_text) {
                            seen_paragraphs.insert(para_text.clone());
                            body_acc.push_str(&para_text);
                            body_acc.push_str("\n\n");
                            found_anything = true;
                        }
                    }
                }

                // Fallback: vecchio metodo per altri siti (con extract_json_string)
                if !found_anything {
                    let mut search_pos = 0;
                    while let Some(text_start) = json_text[search_pos..].find("\"text\":\"") {
                        let abs_start = search_pos + text_start + 8;
                        if abs_start < json_text.len() {
                            if let Some((val, end_pos)) =
                                extract_json_string(&json_text[abs_start..])
                            {
                                if val.len() > 30 && !val.contains("http") && !val.contains("{") {
                                    body_acc.push_str(&val);
                                    body_acc.push_str("\n\n");
                                    found_anything = true;
                                }
                                search_pos = abs_start + end_pos;
                            } else {
                                break;
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
        }
    }

    // 3. FALLBACK CSS
    if !found_anything || body_acc.len() < 300 {
        let content_selectors = [
            "p[data-type='paragraph']", // WSJ modern
            ".wsj-article-body p",
            "article p",
            ".atext",
            ".art-text",
            ".story-content p",
            ".article-body p",
        ];
        for sel_str in content_selectors {
            if let Ok(selector) = Selector::parse(sel_str) {
                let mut sel_acc = String::new();
                for element in document.select(&selector) {
                    let text = element.text().collect::<Vec<_>>().join(" ");
                    if text.to_lowercase().contains("enable js") {
                        continue;
                    }
                    sel_acc.push_str(&text);
                    sel_acc.push_str("\n\n");
                }
                if sel_acc.len() > 200 {
                    body_acc.push_str(&sel_acc);
                    break;
                }
            }
        }
    }

    let mut final_text = String::new();
    if !author_info.is_empty() {
        final_text.push_str(&format!("Di {}\n\n", author_info));
    }
    final_text.push_str(&body_acc);

    let content = clean_text(&final_text);
    let final_content = collapse_blank_lines(&content);
    let excerpt = final_content.chars().take(300).collect::<String>();

    Some(ArticleContent {
        title: title.trim().to_string(),
        content: final_content,
        excerpt,
    })
}

fn pick_title(document: &Html) -> String {
    let title_selectors = ["meta[property='og:title']", "h1", "title"];
    for sel in title_selectors {
        if let Ok(s) = Selector::parse(sel) {
            if let Some(el) = document.select(&s).next() {
                let t = if sel.contains("meta") {
                    el.value().attr("content").unwrap_or("").to_string()
                } else {
                    el.text().collect::<Vec<_>>().join(" ")
                };
                let clean_t = t.trim();
                if clean_t.len() > 5 && !clean_t.to_lowercase().ends_with(".com") {
                    return decode_unicode(clean_t);
                }
            }
        }
    }
    "No Title".to_string()
}

pub fn collapse_blank_lines(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut blank_run = 0usize;
    for line in s.lines() {
        let l = line.trim();
        if l.is_empty() {
            blank_run += 1;
            if blank_run <= 1 {
                out.push('\n');
            }
        } else {
            blank_run = 0;
            out.push_str(l);
            out.push('\n');
        }
    }
    out.trim_end_matches('\n').to_string()
}

#[allow(dead_code)]
pub fn extract_article_links_from_html(_: &str, _: &str, _: usize) -> Vec<(String, String)> {
    Vec::new()
}
#[allow(dead_code)]
pub fn extract_hub_links_from_html(_: &str, _: &str, _: usize) -> Vec<String> {
    Vec::new()
}
#[allow(dead_code)]
pub fn extract_feed_links_from_html(_: &str, _: &str) -> Vec<String> {
    Vec::new()
}
#[allow(dead_code)]
pub fn extract_page_title(html: &str, _: &str) -> String {
    pick_title(&Html::parse_document(html))
}
