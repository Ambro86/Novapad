use crate::log_debug;
use crate::podcast::chapters::Chapter;
use crate::tools::reader;
use feed_rs::parser;
use reqwest::{self, StatusCode, header};

use header::{ETAG, IF_MODIFIED_SINCE, IF_NONE_MATCH, LAST_MODIFIED, REFERER};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::error::Error;
use std::io::Cursor;
use std::sync::{Arc, OnceLock};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tokio::sync::{Mutex, OwnedSemaphorePermit, Semaphore};
use tokio::time::sleep;
use url::Url;

type HttpClient = reqwest::Client;
type HttpError = reqwest::Error;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum RssSourceType {
    Feed,
    Article,
    Site,
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
    #[serde(default)]
    pub last_seen_guid: Option<String>,
}

#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct RssItem {
    pub title: String,
    pub link: String,
    pub description: String,
    pub is_folder: bool,
    pub guid: String,
    pub published: Option<i64>,
}

#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct PodcastEpisode {
    pub title: String,
    pub link: String,
    pub description: String,
    pub guid: String,
    pub published: Option<i64>,
    pub enclosure_url: Option<String>,
    pub enclosure_type: Option<String>,
    pub chapters_url: Option<String>,
    pub chapters_type: Option<String>,
    pub podlove_chapters: Vec<Chapter>,
}

#[derive(Debug, Clone, Copy)]
#[allow(dead_code)]
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
#[allow(dead_code)]
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

#[allow(dead_code)]
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
#[allow(dead_code)]
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
    pub title: String,
    pub items: Vec<PodcastEpisode>,
    pub cache: RssFeedCache,
    pub not_modified: bool,
}

#[allow(dead_code)]
struct RssHttp {
    client: HttpClient,
    global_sem: Arc<Semaphore>,
    per_host_sem: Mutex<HashMap<String, Arc<Semaphore>>>,
    rate_state: Mutex<HashMap<String, HostRateState>>,
    config: RssHttpConfig,
}

impl RssHttp {
    fn new(config: RssHttpConfig) -> Result<Self, String> {
        let mut headers = header::HeaderMap::new();
        headers.insert(
            header::USER_AGENT,
            header::HeaderValue::from_static(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            ),
        );
        headers.insert(
            REFERER,
            header::HeaderValue::from_static("https://news.google.com/"),
        );

        let client = reqwest::Client::builder()
            .default_headers(headers)
            .cookie_store(true)
            .redirect(reqwest::redirect::Policy::limited(10))
            .gzip(true)
            .brotli(true)
            .connect_timeout(Duration::from_secs(10))
            .timeout(Duration::from_secs(30))
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
#[allow(dead_code)]
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
    let normalized = normalize_url(u);
    if let Ok(mut url) = Url::parse(&normalized) {
        url.set_fragment(None);
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
        let _ = url.set_port(None);
        let mut s = url.to_string();
        if let Some(rest) = s.strip_prefix("https://") {
            s = rest.to_string();
        } else if let Some(rest) = s.strip_prefix("http://") {
            s = rest.to_string();
        }
        while s.ends_with('/') && s.len() > 1 {
            s.pop();
        }
        return s;
    }
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

fn format_error_chain(e: &HttpError) -> String {
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

fn now_unix() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64
}

fn should_retry_status(status: StatusCode) -> bool {
    matches!(status.as_u16(), 429 | 502 | 503 | 504 | 508)
}

fn compute_backoff(attempt: usize, max_secs: u64) -> Duration {
    let secs = 1u64
        .checked_shl(attempt as u32)
        .unwrap_or(u64::MAX)
        .min(max_secs);
    Duration::from_secs(secs)
}

#[allow(dead_code)]
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
    log_debug(&format!(
        "rss_request kind=\"{fetch_kind}\" url=\"{url}\" host=\"{host}\" attempt={attempt} status={status_code} not_modified={not_modified} backoff_ms={backoff_ms} error=\"{}\"",
        err.unwrap_or("")
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
    String::new()
}

fn truncate_excerpt(input: &str, max_chars: usize) -> String {
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
    let mut seen = HashSet::new();
    let mut out = Vec::new();
    for item in items {
        let key = if !item.guid.trim().is_empty() {
            format!("guid:{}", item.guid.trim())
        } else {
            format!("link:{}", canonicalize_url(&item.link))
        };
        if seen.insert(key) {
            out.push(item);
            if out.len() >= max_items {
                break;
            }
        }
    }
    out
}

async fn fetch_bytes_with_retries(
    http: &RssHttp,
    url: &str,
    is_feed: bool,
    _fetch_kind: &str,
    _override_cooldown: bool,
    _fetch_config: &RssFetchConfig,
    mut cache: Option<&mut RssFeedCache>,
) -> Result<FetchBytesOutcome, FeedFetchError> {
    let host = host_from_url(url).unwrap_or_else(|| "unknown".to_string());
    let max_attempts = http.config.max_retries + 1;

    for attempt in 1..=max_attempts {
        let response = {
            let _permits = http.acquire_permits(&host).await.map_err(|e| {
                let cache = cache.as_deref().cloned().unwrap_or_default();
                FeedFetchError::Network { message: e, cache }
            })?;
            let mut req = http.client.get(url);
            if is_feed {
                if let Some(c) = cache.as_ref() {
                    if let Some(etag) = c.etag.as_deref() {
                        req = req.header(IF_NONE_MATCH, etag);
                    }
                    if let Some(m) = c.last_modified.as_deref() {
                        req = req.header(IF_MODIFIED_SINCE, m);
                    }
                }
            }
            req.send().await
        };

        match response {
            Ok(resp) => {
                let status = resp.status();
                let headers = resp.headers().clone();
                if let Some(c) = cache.as_deref_mut() {
                    c.last_fetch = Some(now_unix());
                    c.last_status = Some(status.as_u16());
                }
                if status == StatusCode::NOT_MODIFIED && is_feed {
                    return Ok(FetchBytesOutcome {
                        bytes: Vec::new(),
                        not_modified: true,
                    });
                }
                if !status.is_success() {
                    if should_retry_status(status) && attempt < max_attempts {
                        sleep(compute_backoff(attempt - 1, http.config.backoff_max_secs)).await;
                        continue;
                    }
                    let cache = cache.as_deref().cloned().unwrap_or_default();
                    return Err(FeedFetchError::HttpStatus {
                        status: status.as_u16(),
                        kind: "http_error".to_string(),
                        cache,
                    });
                }
                let bytes = resp
                    .bytes()
                    .await
                    .map_err(|e| {
                        let cache = cache.as_deref().cloned().unwrap_or_default();
                        FeedFetchError::Network {
                            message: e.to_string(),
                            cache,
                        }
                    })?
                    .to_vec();
                if let Some(c) = cache.as_deref_mut() {
                    c.consecutive_failures = 0;
                    if let Some(etag) = headers.get(ETAG).and_then(|v| v.to_str().ok()) {
                        c.etag = Some(etag.to_string());
                    }
                    if let Some(m) = headers.get(LAST_MODIFIED).and_then(|v| v.to_str().ok()) {
                        c.last_modified = Some(m.to_string());
                    }
                }
                return Ok(FetchBytesOutcome {
                    bytes,
                    not_modified: false,
                });
            }
            Err(err) => {
                if attempt < max_attempts {
                    sleep(compute_backoff(attempt - 1, http.config.backoff_max_secs)).await;
                    continue;
                }
                let cache = cache.as_deref().cloned().unwrap_or_default();
                return Err(FeedFetchError::Network {
                    message: format_error_chain(&err),
                    cache,
                });
            }
        }
    }
    let cache = cache.as_deref().cloned().unwrap_or_default();
    Err(FeedFetchError::Network {
        message: "Retries exhausted".to_string(),
        cache,
    })
}

struct FetchBytesOutcome {
    bytes: Vec<u8>,
    not_modified: bool,
}

pub async fn fetch_and_parse(
    url: &str,
    _source_kind: RssSourceType,
    cache: RssFeedCache,
    fetch_config: RssFetchConfig,
    override_cooldown: bool,
) -> Result<RssFetchOutcome, FeedFetchError> {
    let url = normalize_url(url);
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
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Feed,
            title: String::new(),
            items: Vec::new(),
            cache,
            not_modified: true,
        });
    }
    if let Some((title, items)) = parse_feed_bytes(out.bytes, &url, fetch_config.max_excerpt_chars)
    {
        return Ok(RssFetchOutcome {
            kind: RssSourceType::Feed,
            title,
            items: dedup_items(items, fetch_config.max_items_per_feed),
            cache,
            not_modified: false,
        });
    }
    Err(FeedFetchError::Network {
        message: "Parsing failed".to_string(),
        cache,
    })
}

pub async fn fetch_url_bytes(
    url: &str,
    fetch_config: RssFetchConfig,
) -> Result<Vec<u8>, FeedFetchError> {
    let http = shared_http().map_err(|e| FeedFetchError::Network {
        message: e,
        cache: RssFeedCache::default(),
    })?;
    let out = fetch_bytes_with_retries(
        http,
        &normalize_url(url),
        false,
        "generic",
        false,
        &fetch_config,
        None,
    )
    .await?;
    Ok(out.bytes)
}

pub async fn fetch_article_text(
    url: &str,
    fallback_title: &str,
    fallback_description: &str,
) -> Result<String, String> {
    let start_total = Instant::now();
    let url_str = normalize_url(url);
    if url_str.is_empty() {
        return Err("Empty URL".to_string());
    }

    log_debug(&format!(
        "rss_article_fetch starting via curl-impersonate url=\"{url_str}\""
    ));
    let url_for_curl = url_str.clone();
    let bytes_res = tokio::task::spawn_blocking(move || {
        crate::curl_client::CurlClient::fetch_url_impersonated(&url_for_curl)
            .map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?;

    let html = match bytes_res {
        Ok(bytes) => {
            let s = String::from_utf8_lossy(&bytes).to_string();
            // DEBUG: Salva l'HTML grezzo in un file vicino all'exe
            if let Ok(mut exe_path) = std::env::current_exe() {
                exe_path.set_file_name("debug_last_fetch.txt");
                let _ = std::fs::write(exe_path, &s);
            }
            s
        }
        Err(err) => {
            log_debug(&format!(
                "rss_article_fetch curl_failed url=\"{url_str}\" error=\"{err}\""
            ));
            return Err(err);
        }
    };

    let article = reader::reader_mode_extract(&html).unwrap_or(reader::ArticleContent {
        title: fallback_title.to_string(),
        content: fallback_description.to_string(),
        excerpt: String::new(),
    });
    log_debug(&format!(
        "rss_article_fetch_done ms={} url=\"{url_str}\"",
        start_total.elapsed().as_millis()
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

// Stubs for missing podcast/itunes functions to keep file structure consistent
pub async fn fetch_podcast_feed(
    url: &str,
    cache: RssFeedCache,
    cfg: RssFetchConfig,
    _: bool,
) -> Result<PodcastFetchOutcome, FeedFetchError> {
    let res = fetch_and_parse(url, RssSourceType::Feed, cache, cfg, true).await?;
    Ok(PodcastFetchOutcome {
        title: res.title,
        items: Vec::new(),
        cache: res.cache,
        not_modified: res.not_modified,
    })
}
