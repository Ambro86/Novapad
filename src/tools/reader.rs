use crate::settings::Language;
use scraper::{Html, Selector};

#[derive(Debug, Clone)]
pub struct ArticleContent {
    pub title: String,
    pub content: String,
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
            if let Ok(code) = u32::from_str_radix(&hex, 16)
                && let Some(decoded_char) = std::char::from_u32(code)
            {
                result.push(decoded_char);
                continue;
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
                        if let Ok(code) = u32::from_str_radix(&hex, 16)
                            && let Some(decoded_char) = std::char::from_u32(code)
                        {
                            result.push(decoded_char);
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

fn decode_html_entities(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let mut chars = input.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '&' {
            let mut entity = String::new();
            while let Some(&next) = chars.peek() {
                chars.next();
                if next == ';' {
                    break;
                }
                if entity.len() >= 16 {
                    entity.push(next);
                    break;
                }
                entity.push(next);
            }
            if entity.starts_with("#x") || entity.starts_with("#X") {
                if let Ok(code) = u32::from_str_radix(&entity[2..], 16)
                    && let Some(decoded) = std::char::from_u32(code)
                {
                    out.push(decoded);
                    continue;
                }
            } else if let Some(num) = entity.strip_prefix('#')
                && let Ok(code) = num.parse::<u32>()
                && let Some(decoded) = std::char::from_u32(code)
            {
                out.push(decoded);
                continue;
            } else {
                match entity.as_str() {
                    "nbsp" => out.push(' '),
                    "amp" => out.push('&'),
                    "quot" => out.push('"'),
                    "apos" => out.push('\''),
                    "hellip" => out.push('…'),
                    "ndash" => out.push('–'),
                    "mdash" => out.push('—'),
                    "rsquo" => out.push('’'),
                    "lsquo" => out.push('‘'),
                    "rdquo" => out.push('”'),
                    "ldquo" => out.push('“'),
                    _ => {
                        out.push('&');
                        out.push_str(&entity);
                        out.push(';');
                    }
                }
                continue;
            }
            out.push('&');
            out.push_str(&entity);
            out.push(';');
        } else {
            out.push(c);
        }
    }
    out
}

fn looks_like_teaser(value: &str) -> bool {
    let v = value.trim();
    if v.len() < 120 {
        return true;
    }
    v.contains("&hellip;")
        || v.contains("&#8230;")
        || v.contains("[&hellip;]")
        || v.contains("[…]")
        || v.contains("[...]")
        || v.ends_with("…")
        || v.ends_with("...")
}

fn looks_like_ui_chrome(value: &str) -> bool {
    let v = value.to_lowercase();
    v.contains("cookie")
        || v.contains("privacy policy")
        || v.contains("terms and conditions")
        || v.contains("sign up")
        || v.contains("log in")
        || v.contains("subscribe")
        || v.contains("newsletter")
        || v.contains("all rights reserved")
        || v.contains("enable js")
        || v.contains("advert")
        || v.contains("sponsored")
        || v.contains("consent")
}

fn count_sentences(value: &str) -> usize {
    value
        .chars()
        .filter(|c| matches!(c, '.' | '!' | '?'))
        .count()
}

fn extract_json_values(json_text: &str, key: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut search_pos = 0;
    while let Some(text_start) = json_text[search_pos..].find(key) {
        let abs_start = search_pos + text_start + key.len();
        if abs_start < json_text.len() {
            if let Some((val, end_pos)) = extract_json_string(&json_text[abs_start..]) {
                out.push(val);
                search_pos = abs_start + end_pos;
            } else {
                break;
            }
        } else {
            break;
        }
    }
    out
}

fn pick_best_json_article_text(json_text: &str) -> Option<String> {
    let keys = [
        "\"articleBody\":\"",
        "\"body\":\"",
        "\"bodyHtml\":\"",
        "\"content\":\"",
        "\"contentHtml\":\"",
        "\"full_text\":\"",
        "\"text\":\"",
    ];
    let mut best = String::new();
    for key in keys {
        for val in extract_json_values(json_text, key) {
            if val.len() < 80 {
                continue;
            }
            let cleaned = clean_text(&val);
            let cleaned = collapse_blank_lines(&cleaned);
            let trimmed = cleaned.trim();
            if trimmed.len() < 300 {
                continue;
            }
            if looks_like_teaser(trimmed) || looks_like_ui_chrome(trimmed) {
                continue;
            }
            if count_sentences(trimmed) < 2 {
                continue;
            }
            if trimmed.len() > best.len() {
                best = trimmed.to_string();
            }
        }
    }
    if best.is_empty() { None } else { Some(best) }
}

pub fn clean_text(input: &str) -> String {
    let decoded = decode_html_entities(&decode_unicode(input));
    // Pulizia encoding Mediaset/TGCOM24
    let mut text = decoded
        .replace("ÃƒÂ¨", "Ã¨")
        .replace("ÃƒÂ ", "Ã ")
        .replace("ÃƒÂ¹", "Ã¹")
        .replace("ÃƒÂ²", "Ã²")
        .replace("ÃƒÂ¬", "Ã¬")
        .replace("Ã‚Â ", " ")
        .replace("ÃƒÂ©", "Ã©")
        .replace("Ã‚", "");

    text = text
        .replace("&nbsp;", " ")
        .replace("&#160;", " ")
        .replace("\u{00a0}", " ");
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

fn author_prefix(language: Language) -> &'static str {
    match language {
        Language::English => "By",
        Language::Italian => "Di",
        Language::French => "Par",
        Language::Spanish => "Por",
        Language::Portuguese => "Por",
        Language::Swedish => "Av",
        Language::Czech => "Od",
        Language::Polish => "Autor",
        Language::Vietnamese => "Bởi",
    }
}

pub fn reader_mode_extract(html_content: &str, language: Language) -> Option<ArticleContent> {
    let document = Html::parse_document(html_content);
    let title = pick_title(&document);

    let mut body_acc = String::new();
    let mut author_info = String::new();
    let mut found_anything = false;

    // 1. ESTRAZIONE DA JSON-LD (Schema.org) - MOLTO RICCO SU TGCOM24
    if let Ok(s) = Selector::parse("script[type='application/ld+json']") {
        for element in document.select(&s) {
            let json = element.text().collect::<Vec<_>>().join("");
            let mut has_rich_json_body = false;

            // Cerchiamo Autore e Data
            if author_info.is_empty() {
                if let Some(author_idx) = json.find("\"author\"") {
                    let author_part = &json[author_idx..];
                    if let Some(name_idx) = author_part.find("\"name\":\"") {
                        let part = &author_part[name_idx + 8..];
                        if let Some((name, _)) = extract_json_string(part) {
                            let trimmed = name.trim();
                            if !trimmed.eq_ignore_ascii_case(&title)
                                && !trimmed.eq_ignore_ascii_case("home")
                                && trimmed.len() >= 3
                            {
                                author_info.push_str(trimmed);
                            }
                        }
                    }
                }
                if author_info.is_empty()
                    && let Some(a_idx) = json.find("\"name\":\"")
                {
                    let part = &json[a_idx + 8..];
                    if let Some((name, _)) = extract_json_string(part) {
                        let trimmed = name.trim();
                        if !trimmed.eq_ignore_ascii_case(&title)
                            && !trimmed.eq_ignore_ascii_case("home")
                            && trimmed.len() >= 3
                        {
                            author_info.push_str(trimmed);
                        }
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
                            if key == "\"description\":\"" {
                                let window_start = key_pos.saturating_sub(400);
                                let window = &json[window_start..key_pos];
                                let is_person_or_org = window.contains("\"@type\":\"Person\"")
                                    || window.contains("\"@type\":\"Organization\"");
                                let is_article = window.contains("\"@type\":\"Article\"")
                                    || window.contains("\"@type\":\"NewsArticle\"")
                                    || window.contains("\"@type\":\"TechArticle\"")
                                    || window.contains("\"@type\":\"BlogPosting\"");
                                if is_person_or_org || !is_article {
                                    search_pos = abs_start + end_pos;
                                    continue;
                                }
                            }
                            if val.len() > 40
                                && !val.contains("http")
                                && !body_acc.contains(&val)
                                && !looks_like_teaser(&val)
                            {
                                body_acc.push_str(&val);
                                body_acc.push_str("\n\n");
                                if key == "\"articleBody\":\"" {
                                    has_rich_json_body = true;
                                }
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
            if has_rich_json_body {
                found_anything = true;
            }
        }
    }

    // 2. ESTRAZIONE DA NEXT_DATA (WSJ / Altri)
    if !found_anything
        && let Ok(next_selector) = Selector::parse("script#__NEXT_DATA__")
        && let Some(element) = document.select(&next_selector).next()
    {
        let json_text = element.text().collect::<Vec<_>>().join("");

        // WSJ: estrai testi dai blocchi "content":[...] dei paragrafi
        let mut seen_paragraphs = std::collections::HashSet::new();
        for content_block in json_text.split("\"type\":\"paragraph\"") {
            if let Some(content_start) = content_block.find("\"content\":[") {
                let after_content = &content_block[content_start..];
                // Estrai tutti i "text":"..." dal blocco content usando extract_json_string
                let mut para_text = String::new();
                let mut search_pos = 0;
                while let Some(text_start) = after_content[search_pos..].find("\"text\":\"") {
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
                    if let Some((val, end_pos)) = extract_json_string(&json_text[abs_start..]) {
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
        if !found_anything && let Some(best) = pick_best_json_article_text(&json_text) {
            body_acc.push_str(&best);
            body_acc.push_str("\n\n");
            found_anything = true;
        }
    }

    // 3. FALLBACK CSS
    if !found_anything || body_acc.len() < 300 {
        let content_selectors = [
            ".node-text .textarea-content-body",
            ".node-summary",
            ".entry-content p",
            ".wp-block-post-content p",
            ".ifq-post__content p",
            ".ifq-post__content",
            "p[data-type='paragraph']", // WSJ modern
            "article [data-testid='article-body'] p",
            "article [data-testid='paragraph']",
            "article [data-type='paragraph']",
            ".prose p",
            ".wsj-article-body p",
            "article p",
            ".atext",
            ".art-text",
            ".story-content p",
            ".article-body p",
            "#col-sx-interna p",
        ];
        let mut best_sel_acc = String::new();
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
                if sel_acc.len() > best_sel_acc.len() {
                    best_sel_acc = sel_acc;
                }
            }
        }
        if best_sel_acc.len() > 200 {
            body_acc.push_str(&best_sel_acc);
        }
    }

    let mut final_text = String::new();
    if !author_info.is_empty() {
        let prefix = author_prefix(language);
        final_text.push_str(&format!("{prefix} {}\n\n", author_info));
    }
    final_text.push_str(&body_acc);

    let content = clean_text(&final_text);
    let mut final_content = collapse_blank_lines(&content);
    if let Some(meta_desc) = pick_meta_description(&document) {
        let should_fallback = body_acc.trim().len() < 120
            || count_sentences(&final_content) < 2
            || looks_like_ui_chrome(&final_content);
        if should_fallback {
            let mut fallback = String::new();
            if !author_info.is_empty() {
                let prefix = author_prefix(language);
                fallback.push_str(&format!("{prefix} {}\n\n", author_info));
            }
            fallback.push_str(meta_desc.trim());
            let fallback_content = collapse_blank_lines(&clean_text(&fallback));
            if fallback_content.len() > final_content.len() {
                final_content = fallback_content;
            }
        }
    }
    Some(ArticleContent {
        title: title.trim().to_string(),
        content: final_content,
    })
}

fn pick_title(document: &Html) -> String {
    let title_selectors = ["meta[property='og:title']", "h1", "title"];
    for sel in title_selectors {
        if let Ok(s) = Selector::parse(sel)
            && let Some(el) = document.select(&s).next()
        {
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
    "No Title".to_string()
}

fn pick_meta_description(document: &Html) -> Option<String> {
    let selectors = [
        "meta[name='description']",
        "meta[property='og:description']",
        "meta[name='twitter:description']",
    ];
    let mut best = String::new();
    for sel in selectors {
        if let Ok(s) = Selector::parse(sel)
            && let Some(el) = document.select(&s).next()
            && let Some(content) = el.value().attr("content")
        {
            let clean = decode_unicode(content.trim());
            if clean.len() > best.len() {
                best = clean;
            }
        }
    }
    if best.len() >= 40 { Some(best) } else { None }
}

pub fn collapse_blank_lines(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut blank_run = 0usize;
    let mut seen_short = std::collections::HashSet::new();
    for line in s.lines() {
        let l = line.trim();
        if l.is_empty() {
            blank_run += 1;
            if blank_run <= 1 {
                out.push('\n');
            }
        } else {
            if let Some(prev) = out.lines().last()
                && prev.eq_ignore_ascii_case(l)
            {
                continue;
            }
            if l.len() <= 40 {
                let key = l.to_ascii_lowercase();
                if !seen_short.insert(key) {
                    continue;
                }
            }
            blank_run = 0;
            out.push_str(l);
            out.push('\n');
        }
    }
    out.trim_end_matches('\n').to_string()
}
