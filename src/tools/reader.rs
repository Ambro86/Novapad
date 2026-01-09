use scraper::{ElementRef, Html, Selector};
use std::collections::HashSet;
use url::Url;

#[derive(Debug, Clone)]
pub struct ArticleContent {
    pub title: String,
    pub content: String,
    #[allow(dead_code)]
    pub excerpt: String,
}

/// Extract *all* links (legacy helper used by older code paths).
/// Returns (absolute_url, label).
#[allow(dead_code)]
pub fn extract_links_from_html(base_url: &str, html_content: &str) -> Vec<(String, String)> {
    let document = Html::parse_document(html_content);
    let selector = Selector::parse("a[href]").unwrap();
    let base = Url::parse(base_url).ok();

    let mut links = Vec::new();
    let mut seen = HashSet::new();

    for element in document.select(&selector) {
        let Some(href) = element.value().attr("href") else {
            continue;
        };

        let full_url = match resolve_url(base.as_ref(), href) {
            Some(u) => u,
            None => continue,
        };
        if !is_http_scheme(&full_url) {
            continue;
        }

        let url_string = full_url.to_string();
        if !seen.insert(url_string.clone()) {
            continue;
        }

        let label = link_label(&element, &full_url);
        if label.trim().is_empty() {
            continue;
        }

        links.push((url_string, label));
        if links.len() >= 200 {
            break;
        }
    }

    links
}

/// Extract "article-looking" links from a generic website page (home / hub pages).
/// Returns (absolute_url, title).
pub fn extract_article_links_from_html(
    base_url: &str,
    html_content: &str,
    max_items: usize,
) -> Vec<(String, String)> {
    let document = Html::parse_document(html_content);
    let selector = Selector::parse("a[href]").unwrap();
    let base = Url::parse(base_url).ok();

    let mut seen = HashSet::new();
    let mut scored: Vec<(i32, String, String)> = Vec::new();

    for a in document.select(&selector) {
        let Some(href) = a.value().attr("href") else {
            continue;
        };
        if href.trim().starts_with('#') {
            continue;
        }
        let full = match resolve_url(base.as_ref(), href) {
            Some(u) => u,
            None => continue,
        };
        if !is_http_scheme(&full) {
            continue;
        }

        // Keep only same-site links (host match) to avoid cross-site noise.
        if let (Some(b), Some(h)) = (base.as_ref(), full.host_str()) {
            if !same_site(b, h) {
                continue;
            }
        }

        let url_s = full.to_string();
        if !seen.insert(url_s.clone()) {
            continue;
        }

        // Filter obvious non-article targets.
        if is_non_article_url(&full) {
            continue;
        }

        let label = link_label(&a, &full);
        if is_skip_link_label(&label) {
            continue;
        }
        let mut score = score_article_candidate(&full, &label);

        // Boost if the link appears within an article/main container.
        if has_articleish_ancestor(&a) {
            score += 5;
        }

        // Discard very low score candidates.
        // Some sites (including those with short titles or menu-driven structures)
        // produce otherwise valid article links that would score 1.
        if score < 1 {
            continue;
        }

        scored.push((score, url_s, label));
    }

    // Sort: higher score first.
    scored.sort_by(|a, b| b.0.cmp(&a.0));

    // Deduplicate by normalized URL path (aggressive).
    let mut out = Vec::new();
    let mut seen_norm = HashSet::new();
    for (_score, url, title) in scored {
        if out.len() >= max_items {
            break;
        }
        let norm = normalize_for_dedup(&url);
        if !seen_norm.insert(norm) {
            continue;
        }
        if title.trim().is_empty() {
            continue;
        }
        out.push((url, title));
    }

    out
}

/// Extract likely "hub" pages (blog, biblioteca, archive, categorie) from the homepage.
/// Returns absolute URLs.
pub fn extract_hub_links_from_html(
    base_url: &str,
    html_content: &str,
    max_hubs: usize,
) -> Vec<String> {
    let document = Html::parse_document(html_content);
    let selector = Selector::parse("a[href]").unwrap();
    let base = Url::parse(base_url).ok();

    let mut scored: Vec<(i32, String)> = Vec::new();
    let mut seen = HashSet::new();

    for a in document.select(&selector) {
        let Some(href) = a.value().attr("href") else {
            continue;
        };
        let full = match resolve_url(base.as_ref(), href) {
            Some(u) => u,
            None => continue,
        };
        if !is_http_scheme(&full) {
            continue;
        }

        if let (Some(b), Some(h)) = (base.as_ref(), full.host_str()) {
            if !same_site(b, h) {
                continue;
            }
        }

        if is_non_article_url(&full) {
            continue;
        }

        let url_s = full.to_string();
        if !seen.insert(url_s.clone()) {
            continue;
        }

        let label = link_label(&a, &full);
        let score = score_hub_candidate(&full, &label);
        if score <= 0 {
            continue;
        }

        scored.push((score, url_s));
    }

    scored.sort_by(|a, b| b.0.cmp(&a.0));

    let mut out = Vec::new();
    for (_score, url) in scored {
        if out.len() >= max_hubs {
            break;
        }
        out.push(url);
    }
    out
}

pub fn extract_feed_links_from_html(base_url: &str, html_content: &str) -> Vec<String> {
    let document = Html::parse_document(html_content);
    let selector = Selector::parse("link[href]").unwrap();
    let base = Url::parse(base_url).ok();

    let mut out = Vec::new();
    let mut seen = HashSet::new();

    for link in document.select(&selector) {
        let Some(href) = link.value().attr("href") else {
            continue;
        };
        let Some(rel) = link.value().attr("rel") else {
            continue;
        };
        if !rel.to_ascii_lowercase().contains("alternate") {
            continue;
        }
        let Some(typ) = link.value().attr("type") else {
            continue;
        };
        let typ_l = typ.to_ascii_lowercase();
        if !(typ_l.contains("rss") || typ_l.contains("atom+xml") || typ_l.contains("xml")) {
            continue;
        }

        let full = match resolve_url(base.as_ref(), href) {
            Some(u) => u,
            None => continue,
        };
        if !is_http_scheme(&full) {
            continue;
        }
        let s = full.to_string();
        if !seen.insert(s.clone()) {
            continue;
        }
        out.push(s);
        if out.len() >= 4 {
            break;
        }
    }

    out
}

/// Extract the best possible title from the HTML content.
pub fn extract_page_title(html_content: &str, _url: &str) -> String {
    let document = Html::parse_document(html_content);
    pick_title(&document)
}

pub fn reader_mode_extract(html_content: &str) -> Option<ArticleContent> {
    let document = Html::parse_document(html_content);

    let title = pick_title(&document);

    // Prefer article/main containers.
    let candidates = [
        "article",
        "main",
        "div#content",
        "div.content",
        "div#main",
        "body",
    ];
    let mut best_root: Option<ElementRef> = None;
    for sel in candidates {
        if let Ok(s) = Selector::parse(sel) {
            if let Some(node) = document.select(&s).next() {
                best_root = Some(node);
                break;
            }
        }
    }
    let root = best_root.unwrap_or_else(|| document.root_element());

    let mut content = extract_text_cleanly(root);
    content = collapse_blank_lines(&content);

    let excerpt = content.chars().take(300).collect::<String>();

    Some(ArticleContent {
        title: title.trim().to_string(),
        content,
        excerpt,
    })
}

fn pick_title(document: &Html) -> String {
    // Prefer H1 if present, else <title>.
    if let Ok(sel_h1) = Selector::parse("h1") {
        if let Some(h1) = document.select(&sel_h1).next() {
            let s = h1.text().collect::<Vec<_>>().join(" ");
            if !s.trim().is_empty() {
                return s;
            }
        }
    }
    if let Ok(sel_t) = Selector::parse("title") {
        if let Some(t) = document.select(&sel_t).next() {
            let s = t.text().collect::<Vec<_>>().join(" ");
            if !s.trim().is_empty() {
                return s;
            }
        }
    }
    "No Title".to_string()
}

fn extract_text_cleanly(element: ElementRef) -> String {
    let mut out = String::new();
    let ignore = [
        "script", "style", "noscript", "iframe", "svg", "nav", "footer", "header", "aside", "form",
    ];
    recursive_text_extract(element, &mut out, &ignore);
    out
}

fn recursive_text_extract(element: ElementRef, out: &mut String, ignore: &[&str]) {
    for child in element.children() {
        if let Some(el) = child.value().as_element() {
            if ignore.contains(&el.name()) {
                continue;
            }

            if is_block_element(el.name()) {
                out.push('\n');
            }

            if let Some(child_ref) = ElementRef::wrap(child) {
                recursive_text_extract(child_ref, out, ignore);
            }

            if is_block_element(el.name()) {
                out.push('\n');
            }
        } else if let Some(txt) = child.value().as_text() {
            let s = txt.trim();
            if !s.is_empty() {
                out.push_str(s);
                out.push(' ');
            }
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
            | "ul"
            | "ol"
            | "blockquote"
            | "pre"
            | "hr"
            | "table"
            | "tr"
            | "section"
            | "article"
    )
}

pub fn collapse_blank_lines(s: &str) -> String {
    // Reduce multiple blank lines to a single blank line.
    let mut out = String::with_capacity(s.len());
    let mut blank_run = 0usize;

    for line in s.lines() {
        let is_blank = line.trim().is_empty();
        if is_blank {
            blank_run += 1;
            if blank_run <= 1 {
                out.push('\n');
            }
        } else {
            blank_run = 0;
            out.push_str(line.trim_end());
            out.push('\n');
        }
    }

    out.trim_end_matches('\n').to_string()
}

fn resolve_url(base: Option<&Url>, href: &str) -> Option<Url> {
    if let Some(b) = base {
        b.join(href).ok()
    } else {
        Url::parse(href).ok()
    }
}

fn is_http_scheme(u: &Url) -> bool {
    matches!(u.scheme(), "http" | "https")
}

fn same_site(base: &Url, other_host: &str) -> bool {
    let Some(base_host) = base.host_str() else {
        return false;
    };
    // allow subdomains
    other_host == base_host || other_host.ends_with(&format!(".{base_host}"))
}

fn link_label(a: &scraper::element_ref::ElementRef<'_>, url: &Url) -> String {
    let text = a.text().collect::<Vec<_>>().join(" ");
    let mut label = text.trim().to_string();
    if label.is_empty() {
        if let Some(t) = a.value().attr("title") {
            label = t.trim().to_string();
        }
    }
    if label.is_empty() {
        label = url.path().trim_matches('/').replace('-', " ");
    }
    normalize_whitespace(&label)
}

fn normalize_whitespace(s: &str) -> String {
    s.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn is_skip_link_label(label: &str) -> bool {
    let l = label.to_lowercase();
    l.contains("skip to content")
        || l.contains("skip content")
        || l.contains("salta al contenuto")
        || l.contains("vai al contenuto")
}

fn normalize_for_dedup(url: &str) -> String {
    // Remove tracking queries/fragments for dedup.
    if let Ok(u) = Url::parse(url) {
        let mut b = format!(
            "{}://{}{}",
            u.scheme(),
            u.host_str().unwrap_or(""),
            u.path()
        );
        // keep first path segment maybe
        if b.ends_with('/') {
            b.pop();
        }
        return b;
    }
    url.to_string()
}

fn is_non_article_url(u: &Url) -> bool {
    let p = u.path().to_lowercase();
    if p.is_empty() || p == "/" {
        return false; // homepage can be article-ish, keep.
    }
    let bad_ext = [
        ".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg", ".pdf", ".zip", ".mp3", ".mp4", ".avi",
        ".mov",
    ];
    if bad_ext.iter().any(|e| p.ends_with(e)) {
        return true;
    }
    let bad_kw = [
        "login",
        "logout",
        "signup",
        "register",
        "privacy",
        "cookie",
        "cookies",
        "terms",
        "contatti",
        "contact",
        "about",
        "chi-siamo",
        "carrello",
        "cart",
        "checkout",
        "wp-admin",
        "wp-login",
        "feed",
        "rss",
        "sitemap",
        "search",
        "tag/",
        "categoria/",
        "category/",
        "author/",
        "comment",
        "forum",
        "account",
    ];
    bad_kw.iter().any(|k| p.contains(k))
}

fn score_article_candidate(u: &Url, label: &str) -> i32 {
    let p = u.path().to_lowercase();
    let mut score = 0;

    // Common article signals
    if p.contains("/blog") || p.contains("/news") || p.contains("/post") || p.contains("/artic") {
        score += 3;
    }
    // Date in path: /2024/03/12/
    if looks_like_date_path(&p) {
        score += 4;
    }
    // Long-ish slug
    if p.split('/').filter(|s| !s.is_empty()).count() >= 2 && p.len() >= 20 {
        score += 2;
    }
    // Label quality
    let l = label.trim();
    if l.len() >= 10 {
        score += 2;
    }
    if l.len() <= 3 {
        score -= 2;
    }

    score
}

fn score_hub_candidate(u: &Url, label: &str) -> i32 {
    let p = u.path().to_lowercase();
    let mut score = 0;

    let hub_kw = [
        "blog",
        "biblioteca",
        "articoli",
        "news",
        "archivio",
        "archive",
        "posts",
        "post",
        "categoria",
        "category",
    ];
    if hub_kw.iter().any(|k| p.contains(k)) {
        score += 5;
    }

    let l = label.to_lowercase();
    if hub_kw.iter().any(|k| l.contains(k)) {
        score += 3;
    }

    // Prefer shorter hub paths
    let segs = p.split('/').filter(|s| !s.is_empty()).count();
    if segs <= 2 {
        score += 1;
    }
    score
}

fn looks_like_date_path(path: &str) -> bool {
    // very lightweight: /YYYY/MM/ or /YYYY/MM/DD/
    let parts: Vec<&str> = path.split('/').filter(|s| !s.is_empty()).collect();
    for w in parts.windows(2) {
        if w[0].len() == 4
            && w[0].chars().all(|c| c.is_ascii_digit())
            && w[1].len() == 2
            && w[1].chars().all(|c| c.is_ascii_digit())
        {
            return true;
        }
    }
    for w in parts.windows(3) {
        if w[0].len() == 4
            && w[0].chars().all(|c| c.is_ascii_digit())
            && w[1].len() == 2
            && w[1].chars().all(|c| c.is_ascii_digit())
            && w[2].len() == 2
            && w[2].chars().all(|c| c.is_ascii_digit())
        {
            return true;
        }
    }
    false
}

fn has_articleish_ancestor(a: &scraper::element_ref::ElementRef<'_>) -> bool {
    // scraper doesn't provide parent traversal easily; approximate by checking if the anchor is inside
    // common wrappers by selecting closest in the subtree is not feasible.
    // We approximate: if the anchor itself has classes suggesting cards/entries.
    if let Some(class) = a.value().attr("class") {
        let c = class.to_lowercase();
        if c.contains("post") || c.contains("entry") || c.contains("title") || c.contains("article")
        {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extract_links() {
        let html = r#"
            <html>
                <body>
                    <a href="https://example.com/1">Link 1</a>
                    <a href="/relative">Link 2</a>
                    <a href="https://example.com/1">Duplicate</a>
                </body>
            </html>
        "#;
        let links = extract_links_from_html("https://example.com", html);
        assert!(links.iter().any(|(u, _)| u == "https://example.com/1"));
        assert!(
            links
                .iter()
                .any(|(u, _)| u == "https://example.com/relative")
        );
        assert_eq!(links.len(), 2);
    }

    #[test]
    fn test_reader_extract() {
        let html = r#"
            <html>
                <head><title>Test Article</title></head>
                <body>
                    <h1>Main Title</h1>
                    <article>
                        <p>This is the content.</p>
                    </article>
                </body>
            </html>
        "#;
        let article = reader_mode_extract(html).unwrap();
        assert_eq!(article.title, "Main Title");
        assert!(article.content.contains("This is the content."));
    }
}
