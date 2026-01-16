use scraper::{Html, Selector};

#[derive(Debug, Clone)]
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
                    if let Some(end) = part.find("\"") {
                        author_info.push_str(&part[..end]);
                    }
                }
                if let Some(d_idx) = json.find("\"datePublished\":\"") {
                    let part = &json[d_idx + 17..];
                    if let Some(end) = part.find("\"") {
                        let date = &part[..10]; // Prendi solo YYYY-MM-DD
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
                for part in json.split(key) {
                    if let Some(end) = part.find("\"") {
                        let val = &part[..end];
                        if val.len() > 40 && !val.contains("http") && !body_acc.contains(val) {
                            body_acc.push_str(val);
                            body_acc.push_str("\n\n");
                            found_anything = true;
                        }
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
                for part in json_text.split("\"text\":\"") {
                    if let Some(end_idx) = part.find("\"") {
                        let val = &part[..end_idx];
                        if val.len() > 30 && !val.contains("http") && !val.contains("{") {
                            body_acc.push_str(val);
                            body_acc.push_str("\n\n");
                            found_anything = true;
                        }
                    }
                }
            }
        }
    }

    // 3. FALLBACK CSS
    if !found_anything || body_acc.len() < 300 {
        let content_selectors = [
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
                    found_anything = true;
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

pub fn extract_article_links_from_html(_: &str, _: &str, _: usize) -> Vec<(String, String)> {
    Vec::new()
}
pub fn extract_hub_links_from_html(_: &str, _: &str, _: usize) -> Vec<String> {
    Vec::new()
}
pub fn extract_feed_links_from_html(_: &str, _: &str) -> Vec<String> {
    Vec::new()
}
pub fn extract_page_title(html: &str, _: &str) -> String {
    pick_title(&Html::parse_document(html))
}
