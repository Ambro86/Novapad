use crate::log_debug;
use crate::podcast::chapters::Chapter;
use crate::tools::reader;
use feed_rs::parser;
use quick_xml::Reader;
use quick_xml::events::Event;
use rand::Rng;
use reqwest::StatusCode;
use reqwest::header::{
    ACCEPT, ACCEPT_LANGUAGE, CONNECTION, ETAG, IF_MODIFIED_SINCE, IF_NONE_MATCH, LAST_MODIFIED,
    REFERER, RETRY_AFTER, SET_COOKIE, UPGRADE_INSECURE_REQUESTS,
};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};
use std::error::Error;
use std::io::Cursor;
use std::sync::{Arc, OnceLock};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tokio::sync::{Mutex, OwnedSemaphorePermit, Semaphore};
use tokio::time::sleep;
use url::Url;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum RssSourceType {
    Feed,    // RSS/Atom
    Article, // Single page (article)
    Site,    // Website (articles list)
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct RssFeedCache {
    #[serde(default)]
    pub feed_url: Option<String>,
    #[serde(default)]
    pub etag: Option<String>,
    #[serde(default)]
    pub last_modified: Option<String>,
    #[serde(default)]
    pub last_fetch: Option<i64>,
    #[serde(default)]
    pub last_status: Option<u16>,
    #[serde(default)]
    pub consecutive_failures: u32,
    #[serde(default)]
    pub blocked_until_epoch_secs: Option<i64>,
    #[serde(default)]
    pub last_error_kind: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RssSource {
    pub title: String,
    pub url: String,
    pub kind: RssSourceType,
    #[serde(default)]
    pub user_title: bool,
    #[serde(default)]
    pub unread: bool,
    #[serde(default)]
    pub cache: RssFeedCache,
    /// GUID of the most recent article the user has seen (for cross-session unread tracking)
    #[serde(default)]
    pub last_seen_guid: Option<String>,
}

#[derive(Debug, Clone)]
pub struct RssItem {
    pub title: String,
    pub link: String,
    pub description: String,
    pub is_folder: bool,
    #[allow(dead_code)]
    pub guid: String,
    #[allow(dead_code)]
    pub published: Option<i64>,
}

#[derive(Debug, Clone)]
pub struct PodcastEpisode {
    pub title: String,
    pub link: String,
    #[allow(dead_code)]
    pub description: String,
    pub guid: String,
    pub published: Option<i64>,
    pub enclosure_url: Option<String>,
    #[allow(dead_code)]
    pub enclosure_type: Option<String>,
    pub chapters_url: Option<String>,
    pub chapters_type: Option<String>,
    pub podlove_chapters: Vec<Chapter>,
}

#[derive(Debug, Clone, Copy)]
pub struct RssFetchConfig {
    pub max_items_per_feed: usize,
    pub max_excerpt_chars: usize,
    pub cooldown_blocked_secs: u64,
    pub cooldown_not_found_secs: u64,
    pub cooldown_rate_limited_secs: u64,
}

impl Default for RssFetchConfig {
    fn default() -> Self {
        Self {
            max_items_per_feed: 5000,
            max_excerpt_chars: 512,
            cooldown_blocked_secs: 3600,
            cooldown_not_found_secs: 86400,
            cooldown_rate_limited_secs: 300,
        }
    }
}

#[derive(Debug, Clone)]
pub enum FeedFetchError {
    InCooldown {
        until: i64,
        kind: String,
        cache: RssFeedCache,
    },
    HttpStatus {
        status: u16,
        kind: String,
        cache: RssFeedCache,
    },
    Network {
        message: String,
        cache: RssFeedCache,
    },
}

impl std::fmt::Display for FeedFetchError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FeedFetchError::InCooldown { until, kind, .. } => {
                write!(f, "Feed in cooldown ({kind}) until {until}")
            }
            FeedFetchError::HttpStatus { status, kind, .. } => {
                write!(f, "HTTP {status} ({kind})")
            }
            FeedFetchError::Network { message, .. } => write!(f, "{message}"),
        }
    }
}

impl FeedFetchError {
    fn cache_clone(&self) -> RssFeedCache {
        match self {
            FeedFetchError::InCooldown { cache, .. } => cache.clone(),
            FeedFetchError::HttpStatus { cache, .. } => cache.clone(),
            FeedFetchError::Network { cache, .. } => cache.clone(),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct RssHttpConfig {
    pub global_max_concurrency: usize,
    pub per_host_max_concurrency: usize,
    pub per_host_rps: u32,
    pub per_host_burst: u32,
    pub max_retries: usize,
    pub backoff_max_secs: u64,
}

impl Default for RssHttpConfig {
    fn default() -> Self {
        Self {
            global_max_concurrency: 8,
            per_host_max_concurrency: 2,
            per_host_rps: 1,
            per_host_burst: 2,
            max_retries: 4,
            backoff_max_secs: 120,
        }
    }
}

#[derive(Debug, Clone)]
pub struct RssFetchOutcome {
    pub kind: RssSourceType,
    pub title: String,
    pub items: Vec<RssItem>,
    pub cache: RssFeedCache,
    pub not_modified: bool,
}

#[derive(Debug, Clone)]
pub struct PodcastFetchOutcome {
    #[allow(dead_code)]
    pub title: String,
    pub items: Vec<PodcastEpisode>,
    pub cache: RssFeedCache,
    #[allow(dead_code)]
    pub not_modified: bool,
}

struct RssHttp {
    client: reqwest::Client,
    global_sem: Arc<Semaphore>,
    per_host_sem: Mutex<HashMap<String, Arc<Semaphore>>>,
    rate_state: Mutex<HashMap<String, HostRateState>>,
    config: RssHttpConfig,
}

impl RssHttp {
    fn new(config: RssHttpConfig) -> Result<Self, String> {
        let mut headers = reqwest::header::HeaderMap::new();
        headers.insert(
            reqwest::header::USER_AGENT,
            reqwest::header::HeaderValue::from_static(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            ),
        );
        headers.insert(
            reqwest::header::ACCEPT_LANGUAGE,
            reqwest::header::HeaderValue::from_static("en-US,en;q=0.9"),
        );
        headers.insert(
            reqwest::header::ACCEPT,
            reqwest::header::HeaderValue::from_static(
                "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            ),
        );
        headers.insert(
            reqwest::header::ACCEPT_ENCODING,
            reqwest::header::HeaderValue::from_static("gzip, deflate, br"),
        );
        headers.insert(
            reqwest::header::CACHE_CONTROL,
            reqwest::header::HeaderValue::from_static("max-age=0"),
        );
        headers.insert(
            reqwest::header::UPGRADE_INSECURE_REQUESTS,
            reqwest::header::HeaderValue::from_static("1"),
        );
        headers.insert(
            "sec-ch-ua",
            reqwest::header::HeaderValue::from_static(
                "\"Not_A Brand\";v=\"8\", \"Chromium\";v=\"120\", \"Google Chrome\";v=\"120\"",
            ),
        );
        headers.insert(
            "sec-ch-ua-mobile",
            reqwest::header::HeaderValue::from_static("?0"),
        );
        headers.insert(
            "sec-ch-ua-platform",
            reqwest::header::HeaderValue::from_static("\"Windows\""),
        );
        headers.insert(
            "sec-fetch-dest",
            reqwest::header::HeaderValue::from_static("document"),
        );
        headers.insert(
            "sec-fetch-mode",
            reqwest::header::HeaderValue::from_static("navigate"),
        );
        headers.insert(
            "sec-fetch-site",
            reqwest::header::HeaderValue::from_static("none"),
        );
        headers.insert(
            "sec-fetch-user",
            reqwest::header::HeaderValue::from_static("?1"),
        );
        headers.insert(
            REFERER,
            reqwest::header::HeaderValue::from_static("https://news.google.com/"),
        );
        let client = reqwest::Client::builder()
            .default_headers(headers)
            .cookie_store(true)
            .redirect(reqwest::redirect::Policy::limited(10))
            .gzip(true)
            .brotli(true)
            .connect_timeout(Duration::from_secs(4))
            .timeout(Duration::from_secs(15))
            .pool_max_idle_per_host(8)
            .build()
            .map_err(|e| e.to_string())?;

        Ok(Self {
            client,
            global_sem: Arc::new(Semaphore::new(config.global_max_concurrency.max(1))),
            per_host_sem: Mutex::new(HashMap::new()),
            rate_state: Mutex::new(HashMap::new()),
            config,
        })
    }

    async fn acquire_permits(&self, host: &str) -> Result<RequestPermits, String> {
        let global = self
            .global_sem
            .clone()
            .acquire_owned()
            .await
            .map_err(|_| "Global concurrency limiter closed".to_string())?;

        let host_sem = {
            let mut map = self.per_host_sem.lock().await;
            map.entry(host.to_string())
                .or_insert_with(|| {
                    Arc::new(Semaphore::new(self.config.per_host_max_concurrency.max(1)))
                })
                .clone()
        };
        let host = host_sem
            .acquire_owned()
            .await
            .map_err(|_| "Per-host concurrency limiter closed".to_string())?;

        Ok(RequestPermits {
            _global: global,
            _host: host,
        })
    }
}

struct RequestPermits {
    _global: OwnedSemaphorePermit,
    _host: OwnedSemaphorePermit,
}

#[derive(Debug, Clone)]
struct HostRateState {
    tokens: f64,
    last: Instant,
}

static RSS_HTTP: OnceLock<Result<RssHttp, String>> = OnceLock::new();

pub fn init_http(config: RssHttpConfig) -> Result<(), String> {
    let res = RSS_HTTP.get_or_init(|| RssHttp::new(config));
    res.as_ref().map(|_| ()).map_err(|e| e.clone())
}

fn shared_http() -> Result<&'static RssHttp, String> {
    let res = RSS_HTTP.get_or_init(|| RssHttp::new(RssHttpConfig::default()));
    res.as_ref().map_err(|e| e.clone())
}

pub fn normalize_url(input: &str) -> String {
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
            if pairs.is_empty() {
                url.set_query(None);
            } else {
                for (k, v) in pairs {
                    url.query_pairs_mut().append_pair(&k, &v);
                }
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

fn host_from_url(url: &str) -> Option<String> {
    Url::parse(url)
        .ok()
        .and_then(|u| u.host_str().map(|h| h.to_string()))
}

fn article_referer(url: &str) -> Option<String> {
    let _ = url;
    Some("https://news.google.com/".to_string())
}

fn now_unix() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64
}

fn should_retry_status(status: StatusCode) -> bool {
    matches!(status.as_u16(), 429 | 502 | 503 | 504 | 508)
}

fn should_retry_error(err: &reqwest::Error) -> bool {
    if err.is_timeout() || err.is_connect() {
        return true;
    }
    if let Some(status) = err.status() {
        return should_retry_status(status);
    }
    err.is_request()
}

fn parse_retry_after(headers: &reqwest::header::HeaderMap) -> Option<Duration> {
    let value = headers.get(RETRY_AFTER)?.to_str().ok()?;
    if let Ok(seconds) = value.trim().parse::<u64>() {
        return Some(Duration::from_secs(seconds));
    }
    if let Ok(date) = httpdate::parse_http_date(value) {
        if let Ok(delta) = date.duration_since(SystemTime::now()) {
            return Some(delta);
        }
    }
    None
}

fn compute_backoff(attempt: usize, max_secs: u64) -> Duration {
    let secs = 1u64
        .checked_shl(attempt as u32)
        .unwrap_or(u64::MAX)
        .min(max_secs);
    Duration::from_secs(secs)
}

fn log_request_attempt(
    url: &str,
    host: &str,
    fetch_kind: &str,
    attempt: usize,
    status: Option<StatusCode>,
    not_modified: bool,
    backoff: Option<Duration>,
    err: Option<&str>,
) {
    let status_code = status.map(|s| s.as_u16()).unwrap_or(0);
    let backoff_ms = backoff.map(|d| d.as_millis()).unwrap_or(0);
    let err_msg = err.unwrap_or("");
    log_debug(&format!(
        "rss_request kind=\"{}\" url=\"{}\" host=\"{}\" attempt={} status={} not_modified={} backoff_ms={} error=\"{}\"",
        fetch_kind, url, host, attempt, status_code, not_modified, backoff_ms, err_msg
    ));
}

fn log_feed_cooldown(
    url: &str,
    host: &str,
    status: u16,
    until: i64,
    kind: &str,
    cooldown_secs: u64,
) {
    log_debug(&format!(
        "rss_cooldown kind=\"feed\" url=\"{}\" host=\"{}\" status={} cooldown_secs={} until_epoch={} reason=\"{}\"",
        url, host, status, cooldown_secs, until, kind
    ));
}

fn log_feed_cooldown_skip(url: &str, host: &str, until: i64, kind: &str) {
    log_debug(&format!(
        "rss_cooldown_skip kind=\"feed\" url=\"{}\" host=\"{}\" status=cooldown until_epoch={} reason=\"{}\"",
        url, host, until, kind
    ));
}

fn parse_feed_bytes(
    bytes: Vec<u8>,
    fallback_title: &str,
    max_excerpt_chars: usize,
) -> Option<(String, Vec<RssItem>)> {
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
                .as_ref()
                .map(|t| t.content.clone())
                .unwrap_or_else(|| "No Title".to_string());
            let link = select_entry_link(&entry);
            let guid = if !entry.id.trim().is_empty() {
                entry.id.clone()
            } else if !link.trim().is_empty() {
                link.clone()
            } else {
                title.clone()
            };
            let published = entry.published.or(entry.updated).map(|d| d.timestamp());
            let description = entry
                .summary
                .as_ref()
                .map(|s| s.content.clone())
                .unwrap_or_default();
            let description = truncate_excerpt(&description, max_excerpt_chars);
            RssItem {
                title,
                link,
                description,
                is_folder: false,
                guid,
                published,
            }
        })
        .collect();
    Some((title, items))
}

fn select_entry_enclosure(entry: &feed_rs::model::Entry) -> (Option<String>, Option<String>) {
    let mut candidate_audio: Option<(String, Option<String>)> = None;
    for media in &entry.media {
        for content in &media.content {
            let Some(url) = content.url.as_ref().map(|u| u.as_str()) else {
                continue;
            };
            let url = url.trim();
            if url.is_empty() {
                continue;
            }
            let media_type = content.content_type.as_ref().map(|m| m.to_string());
            let is_audio = media_type
                .as_deref()
                .map(|m| m.starts_with("audio/"))
                .unwrap_or(false);
            if is_audio {
                return (Some(url.to_string()), media_type);
            }
            if candidate_audio.is_none() && is_audio_extension(url) {
                candidate_audio = Some((url.to_string(), media_type));
            }
        }
    }
    if let Some(content) = entry.content.as_ref().and_then(|c| c.src.as_ref()) {
        let href = content.href.trim();
        if !href.is_empty() {
            let media_type = content.media_type.as_ref().map(|m| m.to_string());
            let is_audio = media_type
                .as_deref()
                .map(|m| m.starts_with("audio/"))
                .unwrap_or(false);
            if is_audio || candidate_audio.is_none() && is_audio_extension(href) {
                return (Some(href.to_string()), media_type);
            }
        }
    }
    for link in &entry.links {
        let href = link.href.trim();
        if href.is_empty() {
            continue;
        }
        let rel = link.rel.as_deref().unwrap_or("");
        let media_type = link.media_type.as_ref().map(|m| m.to_string());
        let is_audio = media_type
            .as_deref()
            .map(|m| m.starts_with("audio/"))
            .unwrap_or(false);
        if rel.eq_ignore_ascii_case("enclosure") || is_audio {
            return (Some(href.to_string()), media_type);
        }
        if candidate_audio.is_none() && is_audio_extension(href) {
            candidate_audio = Some((href.to_string(), media_type));
        }
    }
    if let Some((href, media_type)) = candidate_audio {
        return (Some(href), media_type);
    }
    (None, None)
}

fn is_audio_extension(url: &str) -> bool {
    let lower = url.to_ascii_lowercase();
    lower.ends_with(".mp3")
        || lower.ends_with(".m4a")
        || lower.ends_with(".aac")
        || lower.ends_with(".ogg")
        || lower.ends_with(".opus")
        || lower.ends_with(".wav")
        || lower.ends_with(".flac")
}

fn parse_podcast_feed_bytes(
    bytes: Vec<u8>,
    fallback_title: &str,
    max_excerpt_chars: usize,
) -> Option<(String, Vec<PodcastEpisode>)> {
    let chapters_map = parse_podcast_chapters_map(&bytes);
    let podlove_map = parse_podlove_chapters_map(&bytes);
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
                .as_ref()
                .map(|t| t.content.clone())
                .unwrap_or_else(|| "No Title".to_string());
            let link = select_entry_link(&entry);
            let guid = if !entry.id.trim().is_empty() {
                entry.id.clone()
            } else if !link.trim().is_empty() {
                link.clone()
            } else {
                title.clone()
            };
            let published = entry.published.or(entry.updated).map(|d| d.timestamp());
            let description = entry
                .summary
                .as_ref()
                .map(|s| s.content.clone())
                .unwrap_or_default();
            let description = truncate_excerpt(&description, max_excerpt_chars);
            let (enclosure_url, enclosure_type) = select_entry_enclosure(&entry);
            let chapters_key = podcast_chapters_guid_key(&guid);
            let chapters_link_key = podcast_chapters_link_key(&link);
            let chapters_title_key = podcast_chapters_title_key(&title);
            let chapters_entry = chapters_key
                .as_ref()
                .and_then(|key| chapters_map.get(key))
                .or_else(|| {
                    chapters_link_key
                        .as_ref()
                        .and_then(|key| chapters_map.get(key))
                })
                .or_else(|| {
                    chapters_title_key
                        .as_ref()
                        .and_then(|key| chapters_map.get(key))
                })
                .cloned()
                .unwrap_or((None, None));
            let (chapters_url, chapters_type) = chapters_entry;
            let podlove_entry = chapters_key
                .as_ref()
                .and_then(|key| podlove_map.get(key))
                .or_else(|| {
                    chapters_link_key
                        .as_ref()
                        .and_then(|key| podlove_map.get(key))
                })
                .or_else(|| {
                    chapters_title_key
                        .as_ref()
                        .and_then(|key| podlove_map.get(key))
                })
                .cloned()
                .unwrap_or_default();
            PodcastEpisode {
                title,
                link,
                description,
                guid,
                published,
                enclosure_url,
                enclosure_type,
                chapters_url,
                chapters_type,
                podlove_chapters: podlove_entry,
            }
        })
        .collect();
    Some((title, items))
}

fn podcast_chapters_guid_key(guid: &str) -> Option<String> {
    let trimmed = guid.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(format!("guid:{trimmed}"))
    }
}

fn podcast_chapters_link_key(link: &str) -> Option<String> {
    let trimmed = link.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(format!("link:{}", canonicalize_url(trimmed)))
    }
}

fn podcast_chapters_title_key(title: &str) -> Option<String> {
    let trimmed = title.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(format!("title:{trimmed}"))
    }
}

fn parse_podcast_chapters_map(bytes: &[u8]) -> HashMap<String, (Option<String>, Option<String>)> {
    let mut out = HashMap::new();
    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut current = None;
    let mut current_field: Option<ChaptersField> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                if is_item_start(&name) {
                    current = Some(ChaptersItem::default());
                } else if let Some(item) = current.as_mut() {
                    if eq_ignore_ascii(&name, b"guid") {
                        current_field = Some(ChaptersField::Guid);
                    } else if eq_ignore_ascii(&name, b"id") {
                        current_field = Some(ChaptersField::Guid);
                    } else if eq_ignore_ascii(&name, b"title") {
                        current_field = Some(ChaptersField::Title);
                    } else if eq_ignore_ascii(&name, b"link") {
                        current_field = Some(ChaptersField::Link);
                        if let Some(href) = attr_value(&reader, &e, b"href") {
                            if item.link.is_empty() {
                                item.link = href;
                            }
                        }
                    } else if is_podcast_chapters_tag(&name) {
                        if let Some(url) = attr_value(&reader, &e, b"url") {
                            item.chapters_url = Some(url);
                        }
                        if let Some(kind) = attr_value(&reader, &e, b"type") {
                            item.chapters_type = Some(kind);
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                if let Some(item) = current.as_mut() {
                    if eq_ignore_ascii(&name, b"link") {
                        if let Some(href) = attr_value(&reader, &e, b"href") {
                            if item.link.is_empty() {
                                item.link = href;
                            }
                        }
                    } else if is_podcast_chapters_tag(&name) {
                        if let Some(url) = attr_value(&reader, &e, b"url") {
                            item.chapters_url = Some(url);
                        }
                        if let Some(kind) = attr_value(&reader, &e, b"type") {
                            item.chapters_type = Some(kind);
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if let (Some(item), Some(field)) = (current.as_mut(), current_field) {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match field {
                        ChaptersField::Guid => {
                            if item.guid.is_empty() {
                                item.guid = text;
                            }
                        }
                        ChaptersField::Link => {
                            if item.link.is_empty() {
                                item.link = text;
                            }
                        }
                        ChaptersField::Title => {
                            if item.title.is_empty() {
                                item.title = text;
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                if is_item_end(&name) {
                    if let Some(item) = current.take() {
                        if item
                            .chapters_url
                            .as_ref()
                            .map(|u| !u.trim().is_empty())
                            .unwrap_or(false)
                        {
                            let entry = (item.chapters_url, item.chapters_type);
                            if let Some(key) = podcast_chapters_guid_key(&item.guid) {
                                out.entry(key).or_insert_with(|| entry.clone());
                            }
                            if let Some(key) = podcast_chapters_link_key(&item.link) {
                                out.entry(key).or_insert_with(|| entry.clone());
                            }
                            if let Some(key) = podcast_chapters_title_key(&item.title) {
                                out.entry(key).or_insert(entry);
                            }
                        }
                    }
                } else if eq_ignore_ascii(&name, b"guid")
                    || eq_ignore_ascii(&name, b"id")
                    || eq_ignore_ascii(&name, b"title")
                    || eq_ignore_ascii(&name, b"link")
                {
                    current_field = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[derive(Default)]
struct ChaptersItem {
    guid: String,
    link: String,
    title: String,
    chapters_url: Option<String>,
    chapters_type: Option<String>,
}

#[derive(Clone, Copy)]
enum ChaptersField {
    Guid,
    Link,
    Title,
}

fn is_item_start(name: &[u8]) -> bool {
    eq_ignore_ascii(name, b"item") || eq_ignore_ascii(name, b"entry")
}

fn is_item_end(name: &[u8]) -> bool {
    eq_ignore_ascii(name, b"item") || eq_ignore_ascii(name, b"entry")
}

fn is_podcast_chapters_tag(name: &[u8]) -> bool {
    eq_ignore_ascii(name, b"podcast:chapters")
}

fn eq_ignore_ascii(input: &[u8], expected: &[u8]) -> bool {
    input.len() == expected.len()
        && input
            .iter()
            .zip(expected)
            .all(|(a, b)| a.to_ascii_lowercase() == b.to_ascii_lowercase())
}

fn attr_value<R: std::io::BufRead>(
    reader: &Reader<R>,
    e: &quick_xml::events::BytesStart,
    key: &[u8],
) -> Option<String> {
    for attr in e.attributes().flatten() {
        if eq_ignore_ascii(attr.key.as_ref(), key) {
            return attr
                .decode_and_unescape_value(reader)
                .ok()
                .map(|v| v.to_string());
        }
    }
    None
}

fn parse_podlove_chapters_map(bytes: &[u8]) -> HashMap<String, Vec<Chapter>> {
    let mut out = HashMap::new();
    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut current = None;
    let mut current_field: Option<ChaptersField> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                if is_item_start(&name) {
                    current = Some(PodloveItem::default());
                } else if let Some(item) = current.as_mut() {
                    if eq_ignore_ascii(&name, b"guid") {
                        current_field = Some(ChaptersField::Guid);
                    } else if eq_ignore_ascii(&name, b"id") {
                        current_field = Some(ChaptersField::Guid);
                    } else if eq_ignore_ascii(&name, b"title") {
                        current_field = Some(ChaptersField::Title);
                    } else if eq_ignore_ascii(&name, b"link") {
                        current_field = Some(ChaptersField::Link);
                        if let Some(href) = attr_value(&reader, &e, b"href") {
                            if item.link.is_empty() {
                                item.link = href;
                            }
                        }
                    } else if eq_ignore_ascii(&name, b"psc:chapters") {
                        item.in_chapters = true;
                    } else if eq_ignore_ascii(&name, b"psc:chapter") && item.in_chapters {
                        if let Some(chapter) = parse_podlove_chapter(&reader, &e) {
                            item.chapters.push(chapter);
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                if let Some(item) = current.as_mut() {
                    if eq_ignore_ascii(&name, b"psc:chapter") && item.in_chapters {
                        if let Some(chapter) = parse_podlove_chapter(&reader, &e) {
                            item.chapters.push(chapter);
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if let (Some(item), Some(field)) = (current.as_mut(), current_field) {
                    let text = e.unescape().unwrap_or_default().to_string();
                    match field {
                        ChaptersField::Guid => {
                            if item.guid.is_empty() {
                                item.guid = text;
                            }
                        }
                        ChaptersField::Link => {
                            if item.link.is_empty() {
                                item.link = text;
                            }
                        }
                        ChaptersField::Title => {
                            if item.title.is_empty() {
                                item.title = text;
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                if is_item_end(&name) {
                    if let Some(item) = current.take() {
                        if !item.chapters.is_empty() {
                            let entry = item.chapters;
                            if let Some(key) = podcast_chapters_guid_key(&item.guid) {
                                out.entry(key).or_insert_with(|| entry.clone());
                            }
                            if let Some(key) = podcast_chapters_link_key(&item.link) {
                                out.entry(key).or_insert_with(|| entry.clone());
                            }
                            if let Some(key) = podcast_chapters_title_key(&item.title) {
                                out.entry(key).or_insert(entry);
                            }
                        }
                    }
                } else if eq_ignore_ascii(&name, b"psc:chapters") {
                    if let Some(item) = current.as_mut() {
                        item.in_chapters = false;
                    }
                } else if eq_ignore_ascii(&name, b"guid")
                    || eq_ignore_ascii(&name, b"id")
                    || eq_ignore_ascii(&name, b"title")
                    || eq_ignore_ascii(&name, b"link")
                {
                    current_field = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[derive(Default)]
struct PodloveItem {
    guid: String,
    link: String,
    title: String,
    in_chapters: bool,
    chapters: Vec<Chapter>,
}

fn parse_podlove_chapter<R: std::io::BufRead>(
    reader: &Reader<R>,
    e: &quick_xml::events::BytesStart,
) -> Option<Chapter> {
    let mut start = None;
    let mut title = None;
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let value = attr
            .decode_and_unescape_value(reader)
            .unwrap_or_default()
            .to_string();
        if eq_ignore_ascii(key, b"start") {
            start = Some(value);
        } else if eq_ignore_ascii(key, b"title") {
            title = Some(value);
        }
    }
    let start = start?;
    let title = title?.trim().to_string();
    if title.is_empty() {
        return None;
    }
    let start_ms = parse_podlove_time_ms(&start)?;
    Some(Chapter {
        start_ms,
        title,
        url: None,
        image: None,
    })
}

fn parse_podlove_time_ms(input: &str) -> Option<u64> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return None;
    }
    let parts: Vec<&str> = trimmed.split(':').collect();
    let seconds = match parts.len() {
        1 => parts[0].parse::<f64>().ok()?,
        2 => {
            let minutes = parts[0].parse::<f64>().ok()?;
            let secs = parts[1].parse::<f64>().ok()?;
            minutes * 60.0 + secs
        }
        3 => {
            let hours = parts[0].parse::<f64>().ok()?;
            let minutes = parts[1].parse::<f64>().ok()?;
            let secs = parts[2].parse::<f64>().ok()?;
            hours * 3600.0 + minutes * 60.0 + secs
        }
        _ => return None,
    };
    if seconds < 0.0 {
        return None;
    }
    Some((seconds * 1000.0).floor() as u64)
}

fn select_entry_link(entry: &feed_rs::model::Entry) -> String {
    for link in &entry.links {
        let href = link.href.trim();
        if href.is_empty() {
            continue;
        }
        let rel = link.rel.as_deref().unwrap_or("");
        if rel.is_empty() || rel.eq_ignore_ascii_case("alternate") {
            return href.to_string();
        }
    }
    if let Some(link) = entry.links.iter().find(|l| !l.href.trim().is_empty()) {
        return link.href.clone();
    }
    let id = entry.id.trim();
    if id.starts_with("http://") || id.starts_with("https://") {
        return id.to_string();
    }
    String::new()
}

fn item_dedup_key(item: &RssItem) -> String {
    if !item.guid.trim().is_empty() {
        return format!("guid:{}", item.guid.trim());
    }
    if !item.link.trim().is_empty() {
        return format!("link:{}", canonicalize_url(&item.link));
    }
    let mut hasher = Sha256::new();
    hasher.update(item.title.as_bytes());
    hasher.update(b"|");
    hasher.update(item.link.as_bytes());
    hasher.update(b"|");
    if let Some(ts) = item.published {
        hasher.update(ts.to_string().as_bytes());
    }
    format!("hash:{}", hex::encode(hasher.finalize()))
}

fn podcast_item_dedup_key(item: &PodcastEpisode) -> String {
    if !item.guid.trim().is_empty() {
        return format!("guid:{}", item.guid.trim());
    }
    if !item.link.trim().is_empty() {
        return format!("link:{}", canonicalize_url(&item.link));
    }
    let mut hasher = Sha256::new();
    hasher.update(item.title.as_bytes());
    hasher.update(b"|");
    hasher.update(item.link.as_bytes());
    hasher.update(b"|");
    if let Some(ts) = item.published {
        hasher.update(ts.to_string().as_bytes());
    }
    format!("hash:{}", hex::encode(hasher.finalize()))
}

fn truncate_excerpt(input: &str, max_chars: usize) -> String {
    if max_chars == 0 {
        return String::new();
    }
    let mut out = String::new();
    for (i, ch) in input.chars().enumerate() {
        if i >= max_chars {
            break;
        }
        out.push(ch);
    }
    out
}

fn dedup_items(items: Vec<RssItem>, max_items: usize) -> Vec<RssItem> {
    let max_items = max_items.max(1);
    let mut seen = HashSet::new();
    let mut out = Vec::new();
    for item in items {
        let key = item_dedup_key(&item);
        if seen.insert(key) {
            out.push(item);
            if out.len() >= max_items {
                break;
            }
        }
    }
    out
}

fn dedup_podcast_items(items: Vec<PodcastEpisode>, max_items: usize) -> Vec<PodcastEpisode> {
    if max_items == 0 {
        return Vec::new();
    }
    let mut seen: HashSet<String> = HashSet::new();
    let mut out = Vec::with_capacity(items.len().min(max_items));
    for item in items {
        let key = podcast_item_dedup_key(&item);
        if seen.insert(key) {
            out.push(item);
            if out.len() >= max_items {
                break;
            }
        }
    }
    out
}

fn is_library_hub(url: &str) -> bool {
    let u = url.to_ascii_lowercase();
    u.contains("biblioteca") || u.contains("library") || u.contains("materiale")
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

const SITE_EXTRA_REQUESTS_TOTAL: usize = 8;
const SITE_EXTRA_BURST: usize = 2;
const SITE_EXTRA_PAUSE_MS: u64 = 2000;

async fn wait_for_rate_limit(http: &RssHttp, host: &str) {
    let rps = http.config.per_host_rps.max(1) as f64;
    let burst = http.config.per_host_burst.max(1) as f64;
    loop {
        let wait = {
            let mut map = http.rate_state.lock().await;
            let state = map
                .entry(host.to_string())
                .or_insert_with(|| HostRateState {
                    tokens: burst,
                    last: Instant::now(),
                });
            let now = Instant::now();
            let elapsed = now.duration_since(state.last).as_secs_f64();
            state.tokens = (state.tokens + elapsed * rps).min(burst);
            state.last = now;
            if state.tokens >= 1.0 {
                state.tokens -= 1.0;
                None
            } else {
                let missing = 1.0 - state.tokens;
                Some(Duration::from_secs_f64(missing / rps))
            }
        };
        if let Some(delay) = wait {
            sleep(delay).await;
        } else {
            break;
        }
    }
}

struct FetchBytesOutcome {
    bytes: Vec<u8>,
    not_modified: bool,
}

fn apply_feed_cooldown(
    cache: &mut RssFeedCache,
    status: u16,
    retry_after: Option<Duration>,
    config: &RssFetchConfig,
    now: i64,
) -> (i64, String, u64) {
    let (kind, cooldown_secs) = match status {
        401 => ("blocked_401".to_string(), config.cooldown_blocked_secs),
        403 => ("blocked_403".to_string(), config.cooldown_blocked_secs),
        404 => ("not_found_404".to_string(), config.cooldown_not_found_secs),
        429 => {
            let retry_secs = retry_after.map(|d| d.as_secs()).unwrap_or(0);
            let secs = if retry_secs > 0 {
                retry_secs
            } else {
                config.cooldown_rate_limited_secs
            };
            ("rate_limited_429".to_string(), secs)
        }
        _ => ("unknown".to_string(), config.cooldown_blocked_secs),
    };
    let until = now.saturating_add(cooldown_secs as i64);
    cache.last_status = Some(status);
    cache.consecutive_failures = cache.consecutive_failures.saturating_add(1);
    cache.blocked_until_epoch_secs = Some(until);
    cache.last_error_kind = Some(kind.clone());
    (until, kind, cooldown_secs)
}

async fn fetch_bytes_with_retries(
    http: &RssHttp,
    url: &str,
    is_feed: bool,
    fetch_kind: &str,
    override_cooldown: bool,
    fetch_config: &RssFetchConfig,
    mut cache: Option<&mut RssFeedCache>,
) -> Result<FetchBytesOutcome, FeedFetchError> {
    let host = host_from_url(url).unwrap_or_else(|| "unknown".to_string());
    let max_attempts = http.config.max_retries + 1;

    if is_feed && !override_cooldown {
        if let Some(cache) = cache.as_deref() {
            if let Some(until) = cache.blocked_until_epoch_secs {
                let now = now_unix();
                if now < until {
                    let kind = cache
                        .last_error_kind
                        .clone()
                        .unwrap_or_else(|| "cooldown".to_string());
                    log_feed_cooldown_skip(url, &host, until, &kind);
                    return Err(FeedFetchError::InCooldown {
                        until,
                        kind,
                        cache: cache.clone(),
                    });
                }
            }
        }
    }

    for attempt in 1..=max_attempts {
        wait_for_rate_limit(http, &host).await;
        let response = {
            let _permits = match http.acquire_permits(&host).await {
                Ok(permits) => permits,
                Err(e) => {
                    let cache = cache
                        .as_deref()
                        .cloned()
                        .unwrap_or_else(RssFeedCache::default);
                    return Err(FeedFetchError::Network { message: e, cache });
                }
            };
            let mut req = http.client.get(url);
            if fetch_kind == "article" {
                req = req
                    .timeout(Duration::from_secs(30))
                    .header(
                        ACCEPT,
                        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                    )
                    .header(ACCEPT_LANGUAGE, "en-US,en;q=0.9")
                    .header("DNT", "1")
                    .header(CONNECTION, "keep-alive")
                    .header(UPGRADE_INSECURE_REQUESTS, "1")
                    .header("Sec-Fetch-Dest", "document")
                    .header("Sec-Fetch-Mode", "navigate")
                    .header("Sec-Fetch-Site", "none")
                    .header("Sec-Fetch-User", "?1");
                if let Some(referer) = article_referer(url) {
                    req = req.header(REFERER, referer);
                }
            } else {
                req = req
                    .header(
                        ACCEPT,
                        "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                    )
                    .header(ACCEPT_LANGUAGE, "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");
            }

            if is_feed {
                if let Some(cache) = cache.as_ref() {
                    if cache.feed_url.as_deref() == Some(url) {
                        if let Some(etag) = cache.etag.as_deref() {
                            req = req.header(IF_NONE_MATCH, etag);
                        }
                        if let Some(modified) = cache.last_modified.as_deref() {
                            req = req.header(IF_MODIFIED_SINCE, modified);
                        }
                    }
                }
            }

            req.send().await
        };

        match response {
            Ok(resp) => {
                let status = resp.status();
                let headers = resp.headers().clone();
                if let Some(cache) = cache.as_deref_mut() {
                    cache.last_fetch = Some(now_unix());
                    cache.last_status = Some(status.as_u16());
                }

                if is_feed {
                    let now = now_unix();
                    let code = status.as_u16();
                    if matches!(code, 401 | 403 | 404 | 429) {
                        if let Some(cache) = cache.as_deref_mut() {
                            let retry_after = parse_retry_after(&headers);
                            let (until, kind, cooldown_secs) =
                                apply_feed_cooldown(cache, code, retry_after, fetch_config, now);
                            log_feed_cooldown(url, &host, code, until, &kind, cooldown_secs);
                            return Err(FeedFetchError::HttpStatus {
                                status: code,
                                kind,
                                cache: cache.clone(),
                            });
                        }
                    }
                }

                if status == StatusCode::NOT_MODIFIED && is_feed {
                    if let Some(cache) = cache.as_deref_mut() {
                        cache.consecutive_failures = 0;
                        cache.blocked_until_epoch_secs = None;
                        cache.last_error_kind = None;
                        cache.feed_url = Some(url.to_string());
                        if let Some(etag) = headers.get(ETAG).and_then(|v| v.to_str().ok()) {
                            cache.etag = Some(etag.to_string());
                        }
                        if let Some(modified) =
                            headers.get(LAST_MODIFIED).and_then(|v| v.to_str().ok())
                        {
                            cache.last_modified = Some(modified.to_string());
                        }
                    }
                    log_request_attempt(
                        url,
                        &host,
                        fetch_kind,
                        attempt,
                        Some(status),
                        true,
                        None,
                        None,
                    );
                    return Ok(FetchBytesOutcome {
                        bytes: Vec::new(),
                        not_modified: true,
                    });
                }

                if !status.is_success() {
                    let retry_after = parse_retry_after(resp.headers());
                    if fetch_kind == "article"
                        && status == StatusCode::FORBIDDEN
                        && attempt < max_attempts
                    {
                        let has_cf_challenge = headers
                            .get("cf-mitigated")
                            .and_then(|v| v.to_str().ok())
                            .is_some();
                        let has_cf_cookie = headers
                            .get(SET_COOKIE)
                            .and_then(|v| v.to_str().ok())
                            .is_some_and(|v| v.contains("__cf_bm="));
                        if has_cf_challenge || has_cf_cookie {
                            let delay = Duration::from_secs(1);
                            log_request_attempt(
                                url,
                                &host,
                                fetch_kind,
                                attempt,
                                Some(status),
                                false,
                                Some(delay),
                                Some("cf_challenge_retry"),
                            );
                            sleep(delay).await;
                            continue;
                        }
                    }
                    if is_feed {
                        if let Some(cache) = cache.as_deref_mut() {
                            cache.last_status = Some(status.as_u16());
                            cache.consecutive_failures =
                                cache.consecutive_failures.saturating_add(1);
                        }
                    }
                    if should_retry_status(status) && attempt < max_attempts {
                        let base = compute_backoff(attempt - 1, http.config.backoff_max_secs);
                        let jitter_ms = rand::thread_rng().gen_range(0..=300);
                        let mut delay = base + Duration::from_millis(jitter_ms);
                        if let Some(ra) = retry_after {
                            if ra > delay {
                                delay = ra;
                            }
                        }
                        log_request_attempt(
                            url,
                            &host,
                            fetch_kind,
                            attempt,
                            Some(status),
                            false,
                            Some(delay),
                            None,
                        );
                        sleep(delay).await;
                        continue;
                    }
                    log_request_attempt(
                        url,
                        &host,
                        fetch_kind,
                        attempt,
                        Some(status),
                        false,
                        None,
                        Some("non-retriable status"),
                    );
                    if status.as_u16() == 403 {
                        log_debug(&format!(
                            "rss_request blocked_403 kind=\"{}\" url=\"{}\" host=\"{}\"",
                            fetch_kind, url, host
                        ));
                    }
                    let cache = cache
                        .as_deref()
                        .cloned()
                        .unwrap_or_else(RssFeedCache::default);
                    return Err(FeedFetchError::HttpStatus {
                        status: status.as_u16(),
                        kind: "http_error".to_string(),
                        cache,
                    });
                }

                let bytes = match resp.bytes().await {
                    Ok(b) => b.to_vec(),
                    Err(err) => {
                        let cache = cache
                            .as_deref()
                            .cloned()
                            .unwrap_or_else(RssFeedCache::default);
                        return Err(FeedFetchError::Network {
                            message: err.to_string(),
                            cache,
                        });
                    }
                };
                if is_resource_limit_body(&bytes) && attempt < max_attempts {
                    let delay = compute_backoff(attempt - 1, http.config.backoff_max_secs)
                        + Duration::from_millis(rand::thread_rng().gen_range(0..=300));
                    log_request_attempt(
                        url,
                        &host,
                        fetch_kind,
                        attempt,
                        Some(status),
                        false,
                        Some(delay),
                        Some("resource limit body"),
                    );
                    sleep(delay).await;
                    continue;
                }

                if let Some(cache) = cache.as_deref_mut() {
                    cache.consecutive_failures = 0;
                    cache.blocked_until_epoch_secs = None;
                    cache.last_error_kind = None;
                    cache.feed_url = Some(url.to_string());
                    if let Some(etag) = headers.get(ETAG).and_then(|v| v.to_str().ok()) {
                        cache.etag = Some(etag.to_string());
                    }
                    if let Some(modified) = headers.get(LAST_MODIFIED).and_then(|v| v.to_str().ok())
                    {
                        cache.last_modified = Some(modified.to_string());
                    }
                }

                log_request_attempt(
                    url,
                    &host,
                    fetch_kind,
                    attempt,
                    Some(status),
                    false,
                    None,
                    None,
                );
                return Ok(FetchBytesOutcome {
                    bytes,
                    not_modified: false,
                });
            }
            Err(err) => {
                let err_msg = format_error_chain(&err);
                if fetch_kind == "article" && err.is_timeout() {
                    let cache = cache
                        .as_deref()
                        .cloned()
                        .unwrap_or_else(RssFeedCache::default);
                    log_request_attempt(
                        url,
                        &host,
                        fetch_kind,
                        attempt,
                        err.status(),
                        false,
                        None,
                        Some(&err_msg),
                    );
                    return Err(FeedFetchError::Network {
                        message: err_msg,
                        cache,
                    });
                }
                if should_retry_error(&err) && attempt < max_attempts {
                    let base = compute_backoff(attempt - 1, http.config.backoff_max_secs);
                    let delay = base + Duration::from_millis(rand::thread_rng().gen_range(0..=300));
                    log_request_attempt(
                        url,
                        &host,
                        fetch_kind,
                        attempt,
                        err.status(),
                        false,
                        Some(delay),
                        Some(&err_msg),
                    );
                    sleep(delay).await;
                    continue;
                }
                if is_feed {
                    if let Some(cache) = cache.as_deref_mut() {
                        cache.consecutive_failures = cache.consecutive_failures.saturating_add(1);
                        cache.last_status = err.status().map(|s| s.as_u16());
                    }
                }
                log_request_attempt(
                    url,
                    &host,
                    fetch_kind,
                    attempt,
                    err.status(),
                    false,
                    None,
                    Some(&err_msg),
                );
                let cache = cache
                    .as_deref()
                    .cloned()
                    .unwrap_or_else(RssFeedCache::default);
                return Err(FeedFetchError::Network {
                    message: err_msg,
                    cache,
                });
            }
        }
    }

    let cache = cache
        .as_deref()
        .cloned()
        .unwrap_or_else(RssFeedCache::default);
    Err(FeedFetchError::Network {
        message: "Request retries exhausted".to_string(),
        cache,
    })
}

async fn fetch_site_extra_bytes(
    http: &RssHttp,
    url: &str,
    fetch_config: &RssFetchConfig,
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
    fetch_bytes_with_retries(http, url, false, "site", false, fetch_config, None)
        .await
        .ok()
        .map(|out| out.bytes)
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
pub async fn fetch_and_parse(
    url: &str,
    source_kind: RssSourceType,
    cache: RssFeedCache,
    fetch_config: RssFetchConfig,
    override_cooldown: bool,
) -> Result<RssFetchOutcome, FeedFetchError> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err(FeedFetchError::Network {
            message: "Empty URL".to_string(),
            cache,
        });
    }

    if matches!(source_kind, RssSourceType::Feed) && !override_cooldown {
        if let Some(until) = cache.blocked_until_epoch_secs {
            let now = now_unix();
            if now < until {
                let kind = cache
                    .last_error_kind
                    .clone()
                    .unwrap_or_else(|| "cooldown".to_string());
                if let Some(host) = host_from_url(&url) {
                    log_feed_cooldown_skip(&url, &host, until, &kind);
                }
                return Err(FeedFetchError::InCooldown { until, kind, cache });
            }
        }
    }

    let http = shared_http().map_err(|e| FeedFetchError::Network {
        message: e,
        cache: cache.clone(),
    })?;
    let mut cache = cache;

    if let Some(feed_url) = cache.feed_url.clone() {
        let feed_url = normalize_url(&feed_url);
        if !feed_url.is_empty() {
            let out = fetch_bytes_with_retries(
                http,
                &feed_url,
                true,
                "feed",
                override_cooldown,
                &fetch_config,
                Some(&mut cache),
            )
            .await;
            let out = match out {
                Ok(out) => out,
                Err(err) => return Err(err),
            };
            if out.not_modified {
                return Ok(RssFetchOutcome {
                    kind: RssSourceType::Feed,
                    title: String::new(),
                    items: Vec::new(),
                    cache,
                    not_modified: true,
                });
            }
            if let Some((title, items)) =
                parse_feed_bytes(out.bytes, &feed_url, fetch_config.max_excerpt_chars)
            {
                let items = dedup_items(items, fetch_config.max_items_per_feed);
                return Ok(RssFetchOutcome {
                    kind: RssSourceType::Feed,
                    title,
                    items,
                    cache,
                    not_modified: false,
                });
            }
        }
    }

    let primary_as_feed = matches!(source_kind, RssSourceType::Feed);
    let fetch_kind_main = if primary_as_feed { "feed" } else { "site" };
    let bytes_out = match fetch_bytes_with_retries(
        http,
        &url,
        primary_as_feed,
        fetch_kind_main,
        override_cooldown,
        &fetch_config,
        if primary_as_feed {
            Some(&mut cache)
        } else {
            None
        },
    )
    .await
    {
        Ok(out) => out,
        Err(e1) => {
            if url.starts_with("https://") {
                let http_url = url.replacen("https://", "http://", 1);
                let out = fetch_bytes_with_retries(
                    http,
                    &http_url,
                    primary_as_feed,
                    fetch_kind_main,
                    override_cooldown,
                    &fetch_config,
                    if primary_as_feed {
                        Some(&mut cache)
                    } else {
                        None
                    },
                )
                .await;
                match out {
                    Ok(out) => out,
                    Err(e2) => {
                        let cache = e2.cache_clone();
                        return Err(FeedFetchError::Network {
                            message: format!("{e1} | fallback-http failed: {e2}"),
                            cache,
                        });
                    }
                }
            } else {
                return Err(e1);
            }
        }
    };
    if bytes_out.not_modified && primary_as_feed {
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Feed,
            title: String::new(),
            items: Vec::new(),
            cache,
            not_modified: true,
        });
    }

    let bytes = bytes_out.bytes;
    let mut feed_url: Option<String> = None;
    let mut feed_title: Option<String> = None;
    let mut feed_items: Vec<RssItem> = Vec::new();
    // Try parsing as RSS/Atom feed
    if let Some((title, items)) =
        parse_feed_bytes(bytes.clone(), &url, fetch_config.max_excerpt_chars)
    {
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
            let out = fetch_bytes_with_retries(
                http,
                &candidate,
                true,
                "feed",
                override_cooldown,
                &fetch_config,
                Some(&mut cache),
            )
            .await;
            let Ok(out) = out else {
                continue;
            };
            if out.not_modified {
                return Ok(RssFetchOutcome {
                    kind: RssSourceType::Feed,
                    title: String::new(),
                    items: Vec::new(),
                    cache,
                    not_modified: true,
                });
            }
            if let Some((title, items)) =
                parse_feed_bytes(out.bytes, &candidate, fetch_config.max_excerpt_chars)
            {
                feed_url = Some(candidate);
                feed_title = Some(title);
                feed_items = items;
                break;
            }
        }
    }

    // If we have a feed, return quickly with all items (deduped).
    if feed_url.is_some() {
        let title = feed_title.unwrap_or_else(|| url.clone());
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Feed,
            title,
            items: dedup_items(feed_items, fetch_config.max_items_per_feed),
            cache,
            not_modified: false,
        });
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
            if let Some(hub_bytes) = fetch_site_extra_bytes(
                http,
                &hub,
                &fetch_config,
                &mut extra_requests,
                &mut burst_requests,
            )
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
                            http,
                            &sub,
                            &fetch_config,
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
                    http,
                    &purl,
                    &fetch_config,
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
        if !seen.insert(key.clone()) {
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
            guid: key,
            published: None,
        });
        if items.len() >= target_max {
            break;
        }
    }

    // If we found articles, treat as Site.
    if !feed_items.is_empty() {
        let mut merged = dedup_items(feed_items, fetch_config.max_items_per_feed);
        let mut dedup = HashSet::new();
        for item in &merged {
            dedup.insert(canonicalize_url(&item.link));
        }
        for item in items {
            let key = canonicalize_url(&item.link);
            if dedup.insert(key) {
                merged.push(item);
            }
            if merged.len() >= fetch_config.max_items_per_feed {
                break;
            }
        }
        let title = feed_title.unwrap_or_else(|| page_title.clone());
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Feed,
            title,
            items: merged,
            cache,
            not_modified: false,
        });
    }

    if !items.is_empty() {
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Site,
            title: page_title,
            items,
            cache,
            not_modified: false,
        });
    }

    // If no article links found, treat as Article (single page).
    // This matches your desired UX: pressing Enter imports the page.
    if is_article_meta {
        let items = vec![RssItem {
            title: page_title.clone(),
            link: url.clone(),
            description: String::new(),
            is_folder: false,
            guid: canonicalize_url(&url),
            published: None,
        }];
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Article,
            title: page_title,
            items,
            cache,
            not_modified: false,
        });
    }

    // Last resort: still allow importing the page.
    let items = vec![RssItem {
        title: page_title.clone(),
        link: url.clone(),
        description: String::new(),
        is_folder: false,
        guid: canonicalize_url(&url),
        published: None,
    }];
    Ok(RssFetchOutcome {
        kind: RssSourceType::Article,
        title: page_title,
        items,
        cache,
        not_modified: false,
    })
}

pub async fn fetch_podcast_feed(
    url: &str,
    cache: RssFeedCache,
    fetch_config: RssFetchConfig,
    override_cooldown: bool,
) -> Result<PodcastFetchOutcome, FeedFetchError> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err(FeedFetchError::Network {
            message: "Empty URL".to_string(),
            cache,
        });
    }

    if !override_cooldown {
        if let Some(until) = cache.blocked_until_epoch_secs {
            let now = now_unix();
            if now < until {
                let kind = cache
                    .last_error_kind
                    .clone()
                    .unwrap_or_else(|| "cooldown".to_string());
                if let Some(host) = host_from_url(&url) {
                    log_feed_cooldown_skip(&url, &host, until, &kind);
                }
                return Err(FeedFetchError::InCooldown { until, kind, cache });
            }
        }
    }

    let http = shared_http().map_err(|e| FeedFetchError::Network {
        message: e,
        cache: cache.clone(),
    })?;
    let mut cache = cache;

    let out = fetch_bytes_with_retries(
        http,
        &url,
        true,
        "feed",
        override_cooldown,
        &fetch_config,
        Some(&mut cache),
    )
    .await?;

    if out.not_modified {
        return Ok(PodcastFetchOutcome {
            title: String::new(),
            items: Vec::new(),
            cache,
            not_modified: true,
        });
    }

    if let Some((title, items)) =
        parse_podcast_feed_bytes(out.bytes, &url, fetch_config.max_excerpt_chars)
    {
        let items = dedup_podcast_items(items, fetch_config.max_items_per_feed);
        return Ok(PodcastFetchOutcome {
            title,
            items,
            cache,
            not_modified: false,
        });
    }

    Err(FeedFetchError::Network {
        message: "Unable to parse podcast feed".to_string(),
        cache,
    })
}

#[allow(dead_code)]
pub async fn fetch_itunes_search(
    url: &str,
    fetch_config: RssFetchConfig,
) -> Result<Vec<u8>, FeedFetchError> {
    let http = shared_http().map_err(|e| FeedFetchError::Network {
        message: e,
        cache: RssFeedCache::default(),
    })?;
    let out =
        fetch_bytes_with_retries(http, url, false, "itunes", false, &fetch_config, None).await?;
    Ok(out.bytes)
}

pub async fn fetch_url_bytes(
    url: &str,
    fetch_config: RssFetchConfig,
) -> Result<Vec<u8>, FeedFetchError> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err(FeedFetchError::Network {
            message: "Empty URL".to_string(),
            cache: RssFeedCache::default(),
        });
    }
    let http = shared_http().map_err(|e| FeedFetchError::Network {
        message: e,
        cache: RssFeedCache::default(),
    })?;
    let out =
        fetch_bytes_with_retries(http, &url, false, "generic", false, &fetch_config, None).await?;
    Ok(out.bytes)
}

fn is_probably_blocked_html(html: &str) -> bool {
    let lower = html.to_ascii_lowercase();
    lower.contains("just a moment")
        || lower.contains("cf-browser-verification")
        || lower.contains("cf-chl")
        || lower.contains("attention required")
        || lower.contains("you have a preview of this article while we are checking your access")
        || lower.contains("when we have confirmed access, the full article content will load")
}

pub async fn fetch_article_text(
    url: &str,
    fallback_title: &str,
    fallback_description: &str,
) -> Result<String, String> {
    let start_total = Instant::now();
    let url = normalize_url(url);
    if url.is_empty() {
        return Err("Empty URL".to_string());
    }
    let http = shared_http()?;
    let fetch_config = RssFetchConfig::default();
    let html = {
        let out =
            fetch_bytes_with_retries(http, &url, false, "article", false, &fetch_config, None)
                .await;
        match out {
            Ok(out) => {
                let html = String::from_utf8_lossy(&out.bytes).to_string();
                if is_probably_blocked_html(&html) {
                    log_debug(&format!(
                        "rss_article_fetch reqwest_blocked url=\"{}\"",
                        url
                    ));
                    None
                } else {
                    Some(html)
                }
            }
            Err(err) => {
                log_debug(&format!(
                    "rss_article_fetch reqwest_failed url=\"{}\" error=\"{}\"",
                    url, err
                ));
                None
            }
        }
    };
    let html = html.unwrap_or_default();
    let article = reader::reader_mode_extract(&html).unwrap_or(reader::ArticleContent {
        title: fallback_title.to_string(),
        content: fallback_description.to_string(),
        excerpt: String::new(),
    });
    log_debug(&format!(
        "rss_article_fetch_done ms={} url=\"{}\"",
        start_total.elapsed().as_millis(),
        url
    ));
    Ok(format!("{}\n\n{}", article.title, article.content))
}

pub fn config_from_settings(settings: &crate::settings::AppSettings) -> RssHttpConfig {
    RssHttpConfig {
        global_max_concurrency: settings.rss_global_max_concurrency,
        per_host_max_concurrency: settings.rss_per_host_max_concurrency,
        per_host_rps: settings.rss_per_host_rps,
        per_host_burst: settings.rss_per_host_burst,
        max_retries: settings.rss_max_retries,
        backoff_max_secs: settings.rss_backoff_max_secs,
    }
}

pub fn fetch_config_from_settings(settings: &crate::settings::AppSettings) -> RssFetchConfig {
    RssFetchConfig {
        max_items_per_feed: settings.rss_max_items_per_feed,
        max_excerpt_chars: settings.rss_max_excerpt_chars,
        cooldown_blocked_secs: settings.rss_cooldown_blocked_secs,
        cooldown_not_found_secs: settings.rss_cooldown_not_found_secs,
        cooldown_rate_limited_secs: settings.rss_cooldown_rate_limited_secs,
    }
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

    #[test]
    fn test_host_from_url() {
        assert_eq!(
            host_from_url("https://example.com/path"),
            Some("example.com".to_string())
        );
        assert_eq!(
            host_from_url("http://sub.domain.test/"),
            Some("sub.domain.test".to_string())
        );
        assert_eq!(host_from_url("not a url"), None);
    }

    #[test]
    fn test_should_retry_status() {
        assert!(should_retry_status(StatusCode::TOO_MANY_REQUESTS));
        assert!(should_retry_status(StatusCode::BAD_GATEWAY));
        assert!(should_retry_status(StatusCode::SERVICE_UNAVAILABLE));
        assert!(!should_retry_status(StatusCode::BAD_REQUEST));
        assert!(!should_retry_status(StatusCode::FORBIDDEN));
    }

    #[test]
    fn test_item_dedup_key_prefers_guid() {
        let item = RssItem {
            title: "Title".to_string(),
            link: "https://example.com/a".to_string(),
            description: String::new(),
            is_folder: false,
            guid: "guid-123".to_string(),
            published: None,
        };
        assert_eq!(item_dedup_key(&item), "guid:guid-123");
    }

    #[test]
    fn test_canonicalize_url_strips_tracking() {
        let url = "https://example.com/path/?utm_source=x&utm_medium=y&fbclid=abc#g";
        assert_eq!(canonicalize_url(url), "example.com/path");
    }

    #[test]
    fn test_canonicalize_url_keeps_functional_params() {
        let url = "https://example.com/article?id=123&article=abc&utm_campaign=x";
        assert_eq!(
            canonicalize_url(url),
            "example.com/article?id=123&article=abc"
        );
    }

    #[tokio::test]
    async fn test_feed_cooldown_short_circuit() {
        let now = now_unix();
        let mut cache = RssFeedCache::default();
        cache.blocked_until_epoch_secs = Some(now + 60);
        cache.last_error_kind = Some("blocked_403".to_string());
        let err = fetch_and_parse(
            "https://example.com/feed",
            RssSourceType::Feed,
            cache,
            RssFetchConfig::default(),
            false,
        )
        .await
        .unwrap_err();
        match err {
            FeedFetchError::InCooldown { kind, .. } => {
                assert_eq!(kind, "blocked_403");
            }
            _ => panic!("expected cooldown error"),
        }
    }

    #[test]
    fn test_feed_cooldown_403() {
        let mut cache = RssFeedCache::default();
        let config = RssFetchConfig::default();
        let now = 1000;
        let (until, kind, cooldown_secs) = apply_feed_cooldown(&mut cache, 403, None, &config, now);
        assert_eq!(kind, "blocked_403");
        assert_eq!(cooldown_secs, config.cooldown_blocked_secs);
        assert_eq!(until, now + config.cooldown_blocked_secs as i64);
        assert_eq!(cache.blocked_until_epoch_secs, Some(until));
        assert_eq!(cache.last_error_kind, Some("blocked_403".to_string()));
    }

    #[test]
    fn test_feed_cooldown_429_retry_after() {
        let mut cache = RssFeedCache::default();
        let config = RssFetchConfig::default();
        let now = 1000;
        let (until, kind, cooldown_secs) = apply_feed_cooldown(
            &mut cache,
            429,
            Some(Duration::from_secs(120)),
            &config,
            now,
        );
        assert_eq!(kind, "rate_limited_429");
        assert_eq!(cooldown_secs, 120);
        assert_eq!(until, now + 120);
    }

    #[test]
    fn test_feed_cooldown_429_default() {
        let mut cache = RssFeedCache::default();
        let config = RssFetchConfig::default();
        let now = 1000;
        let (until, kind, cooldown_secs) = apply_feed_cooldown(&mut cache, 429, None, &config, now);
        assert_eq!(kind, "rate_limited_429");
        assert_eq!(cooldown_secs, config.cooldown_rate_limited_secs);
        assert_eq!(until, now + config.cooldown_rate_limited_secs as i64);
    }

    #[tokio::test]
    #[ignore]
    async fn test_article_fetch_profiles_probe() {
        let feeds = std::fs::read_to_string("i18n/feed_en.txt")
            .unwrap_or_default()
            .lines()
            .filter_map(|line| {
                let mut parts = line.split('|').map(str::trim);
                let title = parts.next()?;
                let url = parts.next()?;
                Some((title.to_string(), url.to_string()))
            })
            .filter(|(title, _)| {
                title.contains("NYT")
                    || title.contains("New Scientist")
                    || title.contains("CNN")
                    || title.contains("NPR")
            })
            .collect::<Vec<_>>();
        assert!(!feeds.is_empty(), "no test feeds found");

        for (title, url) in feeds {
            let outcome = fetch_and_parse(
                &url,
                RssSourceType::Feed,
                RssFeedCache::default(),
                RssFetchConfig::default(),
                true,
            )
            .await
            .unwrap();
            let item = outcome
                .items
                .into_iter()
                .find(|item| !item.link.trim().is_empty())
                .expect("feed missing article links");
            let text = fetch_article_text(&item.link, &item.title, &item.description)
                .await
                .unwrap_or_default();
            let lower = text.to_ascii_lowercase();
            if title.contains("NYT") {
                assert!(
                    !lower.contains("preview of our"),
                    "nytimes preview detected: {} ({})",
                    item.title,
                    item.link
                );
            }
            assert!(
                text.len() > 200,
                "short article for {}: {} ({})",
                title,
                item.title,
                item.link
            );
        }
    }

    #[tokio::test]
    #[ignore]
    async fn test_article_fetch_nytimes_specific_url() {
        let url = "https://www.nytimes.com/2026/01/13/business/china-trade-surplus-exports.html";
        let text = fetch_article_text(url, "", "").await.unwrap_or_default();
        let lower = text.to_ascii_lowercase();
        assert!(
            !lower.contains("preview of this article while we are checking your access"),
            "nytimes preview detected for specific url"
        );
        assert!(text.len() > 200, "nytimes article too short");
    }
}
