use crate::tools::reader;
use feed_rs::parser;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::error::Error;
use std::io::Cursor;
use std::time::Duration;
use tokio::time::sleep;
use url::Url;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum RssSourceType {
    Feed,    // RSS/Atom
    Article, // Single page (article)
    Site,    // Website (articles list)
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RssSource {
    pub title: String,
    pub url: String,
    pub kind: RssSourceType,
    #[serde(default)]
    pub user_title: bool,
}

#[derive(Debug, Clone)]
pub struct RssItem {
    pub title: String,
    pub link: String,
    pub description: String,
    pub is_folder: bool,
}

fn normalize_url(input: &str) -> String {
    let s = input.trim();
    if s.is_empty() {
        return String::new();
    }
    if s.starts_with("http://") || s.starts_with("https://") {
        return s.to_string();
    }
    format!("https://{s}")
}

fn canonicalize_url(u: &str) -> String {
    // Stable dedup key: ignore scheme, fragment, common tracking params, and trailing slash
    let normalized = normalize_url(u);
    if let Ok(mut url) = Url::parse(&normalized) {
        url.set_fragment(None);

        // Drop common tracking params
        if url.query().is_some() {
            let pairs: Vec<(String, String)> = url
                .query_pairs()
                .filter(|(k, _)| {
                    let k = k.to_ascii_lowercase();
                    !(k.starts_with("utm_")
                        || k == "gclid"
                        || k == "fbclid"
                        || k == "yclid"
                        || k == "mc_cid"
                        || k == "mc_eid")
                })
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect();

            url.query_pairs_mut().clear();
            for (k, v) in pairs {
                url.query_pairs_mut().append_pair(&k, &v);
            }
        }

        // Remove default ports
        let _ = url.set_port(None);

        let mut s = url.to_string();

        // ignore scheme for dedup
        if let Some(rest) = s.strip_prefix("https://") {
            s = rest.to_string();
        } else if let Some(rest) = s.strip_prefix("http://") {
            s = rest.to_string();
        }

        // strip trailing slash
        while s.ends_with('/') && s.len() > 1 {
            s.pop();
        }
        return s;
    }

    // Fallback: string-based canonicalization
    let mut s = normalized;
    if let Some(rest) = s.strip_prefix("https://") {
        s = rest.to_string();
    } else if let Some(rest) = s.strip_prefix("http://") {
        s = rest.to_string();
    }
    if let Some((left, _)) = s.split_once('#') {
        s = left.to_string();
    }
    if let Some((left, _)) = s.split_once('?') {
        s = left.to_string();
    }
    while s.ends_with('/') && s.len() > 1 {
        s.pop();
    }
    s
}

fn reqwest_client() -> Result<reqwest::Client, String> {
    reqwest::Client::builder()
        .user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36")
        .redirect(reqwest::redirect::Policy::limited(10))
        .timeout(Duration::from_secs(15))
        .build()
        .map_err(|e| e.to_string())
}

fn format_error_chain(e: &reqwest::Error) -> String {
    let mut msg = e.to_string();
    let mut cur: Option<&(dyn Error + 'static)> = e.source();
    while let Some(err) = cur {
        msg.push_str(" | caused by: ");
        msg.push_str(&err.to_string());
        cur = err.source();
    }
    msg
}

fn parse_feed_bytes(bytes: Vec<u8>, fallback_title: &str) -> Option<(String, Vec<RssItem>)> {
    let cursor = Cursor::new(bytes);
    let feed = parser::parse(cursor).ok()?;
    let title = feed
        .title
        .map(|t| t.content)
        .unwrap_or_else(|| fallback_title.to_string());
    let items = feed
        .entries
        .into_iter()
        .map(|entry| {
            let title = entry
                .title
                .map(|t| t.content)
                .unwrap_or_else(|| "No Title".to_string());
            let link = entry
                .links
                .first()
                .map(|l| l.href.clone())
                .unwrap_or_default();
            let description = entry.summary.map(|s| s.content).unwrap_or_default();
            RssItem {
                title,
                link,
                description,
                is_folder: false,
            }
        })
        .collect();
    Some((title, items))
}

fn is_library_hub(url: &str) -> bool {
    let u = url.to_ascii_lowercase();
    u.contains("biblioteca") || u.contains("library") || u.contains("materiale")
}

fn is_feed_url(url: &str) -> bool {
    let u = url.to_ascii_lowercase();
    u.contains("/feed") || u.contains("rss") || u.contains("atom") || u.ends_with(".xml")
}

fn is_resource_limit_body(bytes: &[u8]) -> bool {
    let s = String::from_utf8_lossy(bytes).to_ascii_lowercase();
    if !s.contains("resource limit") {
        return false;
    }
    s.contains("resource limit reached")
        || s.contains("resource limit exceeded")
        || s.contains("resource limit exhausted")
        || s.contains("resource limit exausted")
}

fn request_delay_ms(url: &str) -> u64 {
    if is_feed_url(url) {
        0
    } else if is_library_hub(url) {
        350
    } else {
        120
    }
}

const SITE_EXTRA_REQUESTS_TOTAL: usize = 8;
const SITE_EXTRA_BURST: usize = 2;
const SITE_EXTRA_PAUSE_MS: u64 = 2000;

async fn fetch_bytes(client: &reqwest::Client, url: &str) -> Result<Vec<u8>, String> {
    // Small throttle to reduce host-side rate limits.
    let delay_ms = request_delay_ms(url);
    if delay_ms > 0 {
        sleep(Duration::from_millis(delay_ms)).await;
    }
    let resp = client
        .get(url)
        .header(
            "Accept",
            "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        )
        .header("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7")
        .send()
        .await
        .map_err(|e| format_error_chain(&e))?;

    let bytes = resp.bytes().await.map_err(|e| e.to_string())?.to_vec();

    if is_resource_limit_body(&bytes) {
        sleep(Duration::from_millis(1200)).await;
        let resp = client
            .get(url)
            .header(
                "Accept",
                "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            )
            .header("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7")
            .send()
            .await
            .map_err(|e| format_error_chain(&e))?;
        let retry_bytes = resp.bytes().await.map_err(|e| e.to_string())?.to_vec();
        if is_resource_limit_body(&retry_bytes) {
            return Err("Resource limit reached".to_string());
        }
        return Ok(retry_bytes);
    }

    Ok(bytes)
}

async fn fetch_site_extra_bytes(
    client: &reqwest::Client,
    url: &str,
    extra_requests: &mut usize,
    burst_requests: &mut usize,
) -> Option<Vec<u8>> {
    if *extra_requests >= SITE_EXTRA_REQUESTS_TOTAL {
        return None;
    }
    if *burst_requests >= SITE_EXTRA_BURST {
        sleep(Duration::from_millis(SITE_EXTRA_PAUSE_MS)).await;
        *burst_requests = 0;
    }
    *extra_requests += 1;
    *burst_requests += 1;
    fetch_bytes(client, url).await.ok()
}

fn pagination_variants(base: &str, max_pages: usize) -> Vec<String> {
    // Generate a *small* set of pagination URL variants.
    // We keep this intentionally conservative to avoid excessive network requests.
    // Variants:
    //   /page/N/
    //   ?page=N
    let mut out = Vec::new();
    if max_pages <= 1 {
        return out;
    }

    let page_base = if base.ends_with('/') {
        base.to_string()
    } else {
        format!("{base}/")
    };

    let lower = base.to_lowercase();
    if lower.contains("/page/") || lower.contains("?page=") || lower.contains("?paged=") {
        return out;
    }

    for n in 2..=max_pages {
        out.push(format!("{page_base}page/{n}/"));
        if base.contains('?') {
            out.push(format!("{base}&page={n}"));
        } else {
            out.push(format!("{base}?page={n}"));
        }
    }

    out
}

fn common_hub_paths(url: &str) -> Vec<String> {
    let Ok(mut base) = Url::parse(url) else {
        return Vec::new();
    };
    base.set_path("/");
    base.set_query(None);
    base.set_fragment(None);

    let candidates = [
        "blog/",
        "blogs/",
        "articles/",
        "article/",
        "news/",
        "articoli/",
        "posts/",
        "biblioteca/",
        "biblioteca-sdag/",
        "biblioteca-proposta/",
        "library/",
        "materiale-di-studio/",
    ];

    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for c in candidates {
        if let Ok(u) = base.join(c) {
            let s = u.to_string();
            if seen.insert(s.clone()) {
                out.push(s);
            }
        }
    }
    out
}

/// Fetch an URL and return:
/// - SourceType
/// - title
/// - list of items (leaf articles; for Site returns a flat list of article links)
pub async fn fetch_and_parse(url: &str) -> Result<(RssSourceType, String, Vec<RssItem>), String> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err("Empty URL".to_string());
    }

    let client = reqwest_client()?;

    // Try HTTPS first; if it fails and URL was https, fallback to http once.
    let bytes = match fetch_bytes(&client, &url).await {
        Ok(b) => b,
        Err(e1) => {
            if url.starts_with("https://") {
                let http_url = url.replacen("https://", "http://", 1);
                match fetch_bytes(&client, &http_url).await {
                    Ok(b) => b,
                    Err(e2) => return Err(format!("{e1} | fallback-http failed: {e2}")),
                }
            } else {
                return Err(e1);
            }
        }
    };

    let mut feed_url: Option<String> = None;
    let mut feed_title: Option<String> = None;
    let mut feed_items: Vec<RssItem> = Vec::new();
    // Try parsing as RSS/Atom feed
    if let Some((title, items)) = parse_feed_bytes(bytes.clone(), &url) {
        feed_url = Some(url.clone());
        feed_title = Some(title);
        feed_items = items;
    }

    // HTML mode
    let html = String::from_utf8_lossy(&bytes).to_string();

    // Try discovering feed links from HTML to avoid heavy crawling.
    if feed_url.is_none() {
        let feed_links = reader::extract_feed_links_from_html(&url, &html);
        for candidate in feed_links {
            if let Ok(feed_bytes) = fetch_bytes(&client, &candidate).await {
                if let Some((title, items)) = parse_feed_bytes(feed_bytes.clone(), &candidate) {
                    feed_url = Some(candidate);
                    feed_title = Some(title);
                    feed_items = items;
                    break;
                }
            }
        }
    }

    // If we have a feed, return quickly with the first page items.
    if feed_url.is_some() {
        let mut seen = HashSet::new();
        let mut merged: Vec<RssItem> = Vec::new();
        for (idx, item) in feed_items.into_iter().enumerate() {
            let key = canonicalize_url(&item.link);
            if seen.insert(key) {
                merged.push(item);
            }
            if idx >= 300 {
                break;
            }
        }
        let title = feed_title.unwrap_or_else(|| url.clone());
        return Ok((RssSourceType::Feed, title, merged));
    }

    // Article heuristics (OpenGraph etc.)
    let is_article_meta = html.contains("property=\"og:type\" content=\"article\"")
        || html.contains("property='og:type' content='article'")
        || html.contains("name=\"twitter:card\" content=\"summary_large_image\"")
        || html.contains("name='twitter:card' content='summary_large_image'");

    // Extract a readable title (prefer meta/h1/title; fallback to readability; then URL)
    let mut page_title = reader::extract_page_title(&html, &url);
    if page_title.trim().is_empty() {
        page_title = reader::reader_mode_extract(&html)
            .map(|a| a.title)
            .unwrap_or_else(|| url.clone());
    }
    // First: try article links from homepage (site mode)
    // Try extracting article links from the homepage.
    let target_max: usize = 120;
    let mut article_links = reader::extract_article_links_from_html(&url, &html, target_max);

    // If homepage yields very few articles, do a lightweight "hub discovery" (1 level).
    if article_links.len() < 12 {
        // Lightweight hub discovery: visit a few hub pages (blog/biblioteca/archivio...)
        // and optionally a couple of paginated pages for each hub.
        let mut hubs = reader::extract_hub_links_from_html(&url, &html, 1);
        if hubs.is_empty() {
            hubs = common_hub_paths(&url);
        }
        let mut extra: Vec<(String, String)> = Vec::new();
        let mut extra_requests = 0usize;
        let mut burst_requests = 0usize;
        let mut hub_seen = HashSet::new();

        for hub in hubs {
            if !hub_seen.insert(canonicalize_url(&hub)) {
                continue;
            }
            // Stop early once we have a healthy batch.
            if article_links.len() + extra.len() >= 120 {
                break;
            }
            if extra.len() >= target_max {
                break;
            }

            // Fetch hub page itself.
            if let Some(hub_bytes) =
                fetch_site_extra_bytes(&client, &hub, &mut extra_requests, &mut burst_requests)
                    .await
            {
                let hub_html = String::from_utf8_lossy(&hub_bytes).to_string();
                let mut got = reader::extract_article_links_from_html(&hub, &hub_html, target_max);
                extra.append(&mut got);
                // If this looks like a library hub, try one level of sub-hubs.
                if is_library_hub(&hub) {
                    let sub_hubs = reader::extract_hub_links_from_html(&hub, &hub_html, 1);
                    for sub in sub_hubs {
                        if !hub_seen.insert(canonicalize_url(&sub)) {
                            continue;
                        }
                        let sub_bytes = fetch_site_extra_bytes(
                            &client,
                            &sub,
                            &mut extra_requests,
                            &mut burst_requests,
                        )
                        .await;
                        if let Some(sub_bytes) = sub_bytes {
                            let sub_html = String::from_utf8_lossy(&sub_bytes).to_string();
                            let mut sub_items = reader::extract_article_links_from_html(
                                &sub, &sub_html, target_max,
                            );
                            extra.append(&mut sub_items);
                            if article_links.len() + extra.len() >= 120 {
                                break;
                            }
                        }
                    }
                }
            }

            // Try a couple of paginated variants (common CMS patterns).
            if is_library_hub(&hub) {
                continue;
            }
            for purl in pagination_variants(&hub, 1) {
                if article_links.len() + extra.len() >= 120 {
                    break;
                }
                if extra.len() >= target_max {
                    break;
                }
                let p_bytes = fetch_site_extra_bytes(
                    &client,
                    &purl,
                    &mut extra_requests,
                    &mut burst_requests,
                )
                .await;
                if let Some(p_bytes) = p_bytes {
                    let p_html = String::from_utf8_lossy(&p_bytes).to_string();
                    let mut got =
                        reader::extract_article_links_from_html(&purl, &p_html, target_max);
                    extra.append(&mut got);
                }
            }
        }

        if !extra.is_empty() {
            article_links.extend(extra);
        }
    }

    // Dedup by link
    let mut seen = HashSet::new();
    let mut items = Vec::new();
    for (link, title) in article_links {
        let key = canonicalize_url(&link);
        if !seen.insert(key) {
            continue;
        }
        let t = title.trim();
        if t.is_empty() {
            continue;
        }
        items.push(RssItem {
            title: t.to_string(),
            link,
            description: String::new(),
            is_folder: false, // IMPORTANT: flat list, no navigation
        });
        if items.len() >= target_max {
            break;
        }
    }

    // If we found articles, treat as Site.
    if !feed_items.is_empty() {
        let mut merged = feed_items;
        let mut dedup = HashSet::new();
        for item in &merged {
            dedup.insert(canonicalize_url(&item.link));
        }
        for item in items {
            let key = canonicalize_url(&item.link);
            if dedup.insert(key) {
                merged.push(item);
            }
            if merged.len() >= target_max {
                break;
            }
        }
        let title = feed_title.unwrap_or_else(|| page_title.clone());
        return Ok((RssSourceType::Feed, title, merged));
    }

    if !items.is_empty() {
        return Ok((RssSourceType::Site, page_title, items));
    }

    // If no article links found, treat as Article (single page).
    // This matches your desired UX: pressing Enter imports the page.
    if is_article_meta {
        let items = vec![RssItem {
            title: page_title.clone(),
            link: url.clone(),
            description: String::new(),
            is_folder: false,
        }];
        return Ok((RssSourceType::Article, page_title, items));
    }

    // Last resort: still allow importing the page.
    let items = vec![RssItem {
        title: page_title.clone(),
        link: url.clone(),
        description: String::new(),
        is_folder: false,
    }];
    Ok((RssSourceType::Article, page_title, items))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    async fn test_normalize_url() {
        assert_eq!(normalize_url("example.com"), "https://example.com");
        assert_eq!(
            normalize_url(" https://example.com "),
            "https://example.com"
        );
    }
}
