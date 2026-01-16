use scraper::{Html, Selector};
use url::Url;

#[derive(Debug, Clone)]
pub struct ArticleContent {
    pub title: String,
    pub content: String,
    pub excerpt: String,
}

/// Funzione per pulire il testo da entità HTML, rimasugli JSON e spazi strani
pub fn clean_text(input: &str) -> String {
    let mut text = input.to_string();

    // 1. Rimpiazza entità HTML comuni
    text = text
        .replace("&nbsp;", " ")
        .replace("&#160;", " ")
        .replace("\u{00a0}", " ") // Carattere Unicode non-breaking space
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&#39;", "'")
        .replace("&#039;", "'")
        .replace("&#x27;", "'")
        .replace("&ndash;", "–")
        .replace("&mdash;", "—")
        .replace("&laquo;", "«")
        .replace("&raquo;", "»")
        .replace("&lsquo;", "‘")
        .replace("&rsquo;", "’")
        .replace("&ldquo;", "“")
        .replace("&rdquo;", "”");

    // 2. Rimpiazza sequenze di escape JSON (per estrazione da JSON-LD)
    text = text
        .replace("\\n", "\n")
        .replace("\\\"", "\"")
        .replace("\\u003c", "<")
        .replace("\\u003e", ">")
        .replace("\\u0026", "&")
        .replace("\\u0027", "'");

    // 3. Rimuove tag HTML rimasti (es. <br>, <b>)
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
    let raw_title = pick_title(&document);
    let title = clean_text(&raw_title);

    let mut combined_content = String::new();
    let mut found_anything = false;

    // 1. ESTRAZIONE DA JSON-LD (Schema.org) - La più pulita
    if let Ok(json_selector) = Selector::parse("script[type='application/ld+json']") {
        for element in document.select(&json_selector) {
            let json_text = element.text().collect::<Vec<_>>().join("");
            if let Some(start_idx) = json_text.find("\"articleBody\":\"") {
                let body_part = &json_text[start_idx + 15..];
                // Cerchiamo la fine del campo, facendo attenzione alle virgolette chiuse
                if let Some(end_idx) = body_part.find("\",\"") {
                    let body = &body_part[..end_idx];
                    if body.len() > 300 {
                        combined_content = body.to_string();
                        found_anything = true;
                        break;
                    }
                }
            }
        }
    }

    // 2. ESTRAZIONE DA SELETTORI CSS (per Corriere, Adnkronos, ecc.)
    if !found_anything {
        let content_selectors = [
            ".atext",         // Il Sole 24 Ore
            ".art-text",      // Adnkronos
            ".story-content", // Corriere della Sera
            ".item-text",     // Corriere
            "article p",      // Standard
            ".article-body p",
            ".entry-content p",
            ".post-content p",
        ];

        for sel_str in content_selectors {
            if let Ok(selector) = Selector::parse(sel_str) {
                let matches: Vec<_> = document.select(&selector).collect();
                if !matches.is_empty() {
                    for element in matches {
                        let text = element.text().collect::<Vec<_>>().join(" ");
                        let trimmed = text.trim();
                        if !trimmed.is_empty() {
                            combined_content.push_str(trimmed);
                            combined_content.push_str("\n\n");
                            found_anything = true;
                        }
                    }
                    if found_anything && (sel_str.starts_with('.') || sel_str.starts_with('[')) {
                        break;
                    }
                }
            }
        }
    }

    // 3. FALLBACK: Body
    if !found_anything {
        if let Ok(s) = Selector::parse("body") {
            if let Some(node) = document.select(&s).next() {
                combined_content = extract_text_cleanly(node);
            }
        }
    }

    // PULIZIA FINALE
    let content = clean_text(&combined_content);
    let content = collapse_blank_lines(&content);
    let excerpt = content.chars().take(300).collect::<String>();

    Some(ArticleContent {
        title: title.trim().to_string(),
        content,
        excerpt,
    })
}

fn pick_title(document: &Html) -> String {
    let title_selectors = [
        "h1",
        "title",
        "meta[property='og:title']",
        "meta[name='twitter:title']",
    ];
    for sel in title_selectors {
        if let Ok(s) = Selector::parse(sel) {
            if let Some(el) = document.select(&s).next() {
                let t = if sel.contains("meta") {
                    el.value().attr("content").unwrap_or("").to_string()
                } else {
                    el.text().collect::<Vec<_>>().join(" ")
                };
                if !t.trim().is_empty() {
                    return t;
                }
            }
        }
    }
    "No Title".to_string()
}

fn extract_text_cleanly(element: scraper::ElementRef) -> String {
    let mut out = String::new();
    let ignore = [
        "script", "style", "noscript", "nav", "footer", "header", "aside", "iframe",
    ];
    recursive_text_extract(element, &mut out, &ignore);
    out
}

fn recursive_text_extract(element: scraper::ElementRef, out: &mut String, ignore: &[&str]) {
    for child in element.children() {
        if let Some(el) = child.value().as_element() {
            let name = el.name();
            if ignore.contains(&name) {
                continue;
            }
            if is_block_element(name) {
                out.push('\n');
            }
            if let Some(child_ref) = scraper::ElementRef::wrap(child) {
                recursive_text_extract(child_ref, out, ignore);
            }
            if is_block_element(name) {
                out.push('\n');
            }
        } else if let Some(txt) = child.value().as_text() {
            out.push_str(txt.trim());
            out.push(' ');
        }
    }
}

fn is_block_element(name: &str) -> bool {
    matches!(
        name,
        "p" | "div"
            | "h1"
            | "h2"
            | "h3"
            | "h4"
            | "h5"
            | "h6"
            | "li"
            | "blockquote"
            | "section"
            | "article"
    )
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
