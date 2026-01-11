use crate::tools::reader;
use feed_rs::parser;
use rand::Rng;
use reqwest::StatusCode;
use reqwest::header::{ETAG, IF_MODIFIED_SINCE, IF_NONE_MATCH, LAST_MODIFIED, RETRY_AFTER};
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
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RssSource {
    pub title: String,
    pub url: String,
    pub kind: RssSourceType,
    #[serde(default)]
    pub user_title: bool,
    #[serde(default)]
    pub cache: RssFeedCache,
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

#[derive(Debug, Clone, Copy)]
pub struct RssFetchConfig {
    pub max_items_per_feed: usize,
    pub max_excerpt_chars: usize,
}

impl Default for RssFetchConfig {
    fn default() -> Self {
        Self {
            max_items_per_feed: 5000,
            max_excerpt_chars: 512,
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

struct RssHttp {
    client: reqwest::Client,
    global_sem: Arc<Semaphore>,
    per_host_sem: Mutex<HashMap<String, Arc<Semaphore>>>,
    rate_state: Mutex<HashMap<String, HostRateState>>,
    config: RssHttpConfig,
}

impl RssHttp {
    fn new(config: RssHttpConfig) -> Result<Self, String> {
        let user_agent = format!("Novapad/{} (RSS reader)", env!("CARGO_PKG_VERSION"));
        let client = reqwest::Client::builder()
            .user_agent(user_agent)
            .redirect(reqwest::redirect::Policy::limited(10))
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
    eprintln!(
        "rss_request kind=\"{}\" url=\"{}\" host=\"{}\" attempt={} status={} not_modified={} backoff_ms={} error=\"{}\"",
        fetch_kind, url, host, attempt, status_code, not_modified, backoff_ms, err_msg
    );
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
                .map(|t| t.content)
                .unwrap_or_else(|| "No Title".to_string());
            let link = entry
                .links
                .first()
                .map(|l| l.href.clone())
                .unwrap_or_default();
            let guid = if !entry.id.trim().is_empty() {
                entry.id.clone()
            } else if !link.trim().is_empty() {
                link.clone()
            } else {
                title.clone()
            };
            let published = entry.published.or(entry.updated).map(|d| d.timestamp());
            let description = entry.summary.map(|s| s.content).unwrap_or_default();
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

async fn fetch_bytes_with_retries(
    http: &RssHttp,
    url: &str,
    is_feed: bool,
    fetch_kind: &str,
    mut cache: Option<&mut RssFeedCache>,
) -> Result<FetchBytesOutcome, String> {
    let host = host_from_url(url).unwrap_or_else(|| "unknown".to_string());
    let max_attempts = http.config.max_retries + 1;

    for attempt in 1..=max_attempts {
        wait_for_rate_limit(http, &host).await;
        let response = {
            let _permits = http.acquire_permits(&host).await?;
            let mut req = http
                .client
                .get(url)
                .header(
                    "Accept",
                    "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                )
                .header("Accept-Language", "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7");

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

                if status == StatusCode::NOT_MODIFIED && is_feed {
                    if let Some(cache) = cache.as_deref_mut() {
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
                    return Err(format!("HTTP {status}"));
                }

                let bytes = resp.bytes().await.map_err(|e| e.to_string())?.to_vec();
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
                return Err(err_msg);
            }
        }
    }

    Err("Request retries exhausted".to_string())
}

async fn fetch_site_extra_bytes(
    http: &RssHttp,
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
    fetch_bytes_with_retries(http, url, false, "site", None)
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
) -> Result<RssFetchOutcome, String> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err("Empty URL".to_string());
    }

    let http = shared_http()?;
    let mut cache = cache;

    if let Some(feed_url) = cache.feed_url.clone() {
        let feed_url = normalize_url(&feed_url);
        if !feed_url.is_empty() {
            if let Ok(out) =
                fetch_bytes_with_retries(http, &feed_url, true, "feed", Some(&mut cache)).await
            {
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
    }

    let primary_as_feed = matches!(source_kind, RssSourceType::Feed);
    let fetch_kind_main = if primary_as_feed { "feed" } else { "site" };
    let bytes_out = match fetch_bytes_with_retries(
        http,
        &url,
        primary_as_feed,
        fetch_kind_main,
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
                fetch_bytes_with_retries(
                    http,
                    &http_url,
                    primary_as_feed,
                    fetch_kind_main,
                    if primary_as_feed {
                        Some(&mut cache)
                    } else {
                        None
                    },
                )
                .await
                .map_err(|e2| format!("{e1} | fallback-http failed: {e2}"))?
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
            let out =
                fetch_bytes_with_retries(http, &candidate, true, "feed", Some(&mut cache)).await;
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
            if let Some(hub_bytes) =
                fetch_site_extra_bytes(http, &hub, &mut extra_requests, &mut burst_requests).await
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
                let p_bytes =
                    fetch_site_extra_bytes(http, &purl, &mut extra_requests, &mut burst_requests)
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

pub async fn fetch_article_text(
    url: &str,
    fallback_title: &str,
    fallback_description: &str,
) -> Result<String, String> {
    let url = normalize_url(url);
    if url.is_empty() {
        return Err("Empty URL".to_string());
    }
    let http = shared_http()?;
    let out = fetch_bytes_with_retries(http, &url, false, "article", None).await?;
    let html = String::from_utf8_lossy(&out.bytes).to_string();
    let article = reader::reader_mode_extract(&html).unwrap_or(reader::ArticleContent {
        title: fallback_title.to_string(),
        content: fallback_description.to_string(),
        excerpt: String::new(),
    });
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
}
