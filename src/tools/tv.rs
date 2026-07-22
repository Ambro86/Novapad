use base64::Engine;
use chrono::{Local, NaiveDate};
use flate2::read::GzDecoder;
use reqwest::blocking::Client;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;
use std::fs;
use std::io::Read;
use std::sync::{Mutex, OnceLock};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use url::Url;

const TV_CHANNELS_URL: &str = "https://sonarpad.com/api/tv_channels_resolver.php?resolve=0";
const TV_API_CLIENT_TOKEN: &str = match option_env!("SONARPAD_ROUTE_CLIENT_TOKEN") {
    Some(token) => token,
    None => "",
};
const TV_API_TOKEN_HEADER: &str = "X-Sonarpad-TV-Token";
const ROUTE_API_TOKEN_HEADER: &str = "X-Sonarpad-Route-Token";
const DEFAULT_USER_AGENT: &str = "Sonarpad TV/1.0";
const IOS_SAFARI_PLAYBACK_USER_AGENT: &str = "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1";
const LA7_STREAM_URL: &str = "https://d1chghleocc9sm.cloudfront.net/v1/master/3722c60a815c199d9c0ef36c5b73da68a62b09d1/cc-evfku205gqrtf/Live.m3u8";
const LA7_CINEMA_DASH_URL: &str = "https://d15umi5iaezxgx.cloudfront.net/HBBTV/LA7D/DASH/Live.mpd";
const REGIONAL_CATEGORY_PREFIX: &str = "Regionali - ";
const TV_GUIDE_CHANNEL_PAYLOAD_JSON: &str = r#"{"payload_b64":"csAxIXZQMnhMMiawFTr6bjtEskCkzkNJJ+Zweyc6I0xoq5wAQq2me+nsGOl55vyuggHwBZyk/4KnTrP2iV7rNEEN7i90j4pqQXbXPAgPICMLN0By","algorithm":"gzip-xor-base64-v1"}"#;
const TV_GUIDE_TIMELINE_PAYLOAD_JSON: &str = r#"{"payload_b64":"csAxIXZQMnhMMuhZfR1S+OWXPRn4oJR5K4nkpYbgWGup/jgB+m6jPWForBe9oLtOwaBOreEeoqetOYbKLTxeLIC4fDkh4S9vy3U4I3E=","algorithm":"gzip-xor-base64-v1"}"#;
const TV_GUIDE_STATIC_KEY_PARTS: &[&[u8]] = &[b"sonar", b"pad-", b"SonarSecure-"];
const TV_GUIDE_FALLBACK_MAX_AGE_SECS: i64 = 6 * 60 * 60;
const TV_GUIDE_CACHE_TTL: Duration = Duration::from_secs(5 * 60);

static TV_CHANNEL_GUIDE_CACHE: OnceLock<Mutex<HashMap<String, CachedTvChannelGuide>>> =
    OnceLock::new();
static TV_TIMELINE_CACHE: OnceLock<Mutex<Option<CachedTvTimeline>>> = OnceLock::new();

#[derive(Clone, Debug, Serialize, Deserialize)]
pub(crate) struct TvChannel {
    pub(crate) name: String,
    pub(crate) url: String,
    #[serde(default)]
    pub(crate) dash_url: Option<String>,
    pub(crate) category: String,
    #[serde(default)]
    pub(crate) stream_resolver: Option<String>,
    #[serde(default)]
    pub(crate) resolver_endpoint: Option<String>,
    #[serde(default)]
    pub(crate) resolver_realm: Option<String>,
    #[serde(default)]
    pub(crate) resolver_channel_id: Option<String>,
    #[serde(default)]
    pub(crate) tvg_id: String,
    #[serde(default)]
    pub(crate) tvg_name: String,
    #[serde(default)]
    pub(crate) http_user_agent: String,
}

impl TvChannel {
    pub(crate) fn playback_user_agent(&self) -> &str {
        let value = self.http_user_agent.trim();
        if value.is_empty() {
            DEFAULT_USER_AGENT
        } else {
            value
        }
    }

    /// User-Agent da usare per la riproduzione vera e propria.
    /// Sonarpad mobile usa Safari iPhone quando il server non fornisce
    /// un User-Agent specifico per il canale; manteniamo lo stesso comportamento
    /// anche su Windows, senza cambiare le richieste dei resolver.
    pub(crate) fn media_playback_user_agent(&self) -> &str {
        let value = self.http_user_agent.trim();
        if value.is_empty() {
            IOS_SAFARI_PLAYBACK_USER_AGENT
        } else {
            value
        }
    }

    pub(crate) fn is_regional(&self) -> bool {
        self.category.starts_with(REGIONAL_CATEGORY_PREFIX)
    }

    pub(crate) fn regional_name(&self) -> Option<&str> {
        self.category
            .strip_prefix(REGIONAL_CATEGORY_PREFIX)
            .map(str::trim)
            .filter(|value| !value.is_empty())
    }
}

#[derive(Clone, Debug)]
pub(crate) struct TvChannelLoadResult {
    pub(crate) channels: Vec<TvChannel>,
    pub(crate) cache_warning: Option<String>,
}

#[derive(Clone, Debug)]
pub(crate) struct TvProgram {
    pub(crate) title: String,
    pub(crate) start_time: i64,
    pub(crate) end_time: i64,
}

#[derive(Clone)]
struct CachedTvChannelGuide {
    fetched_at: Instant,
    programs: Vec<TvProgram>,
}

#[derive(Clone)]
struct CachedTvTimeline {
    date: NaiveDate,
    fetched_at: Instant,
    programs_by_channel: HashMap<String, Vec<TvProgram>>,
    exact_channel_names: HashMap<String, Vec<String>>,
}

struct TvTimelineData {
    programs_by_channel: HashMap<String, Vec<TvProgram>>,
    exact_channel_names: HashMap<String, Vec<String>>,
}

#[derive(Debug, Deserialize)]
struct EncryptedTvGuidePayload {
    payload_b64: String,
    algorithm: String,
}

#[derive(Debug, Deserialize)]
struct TvServerResponse {
    #[serde(default)]
    channels: Vec<TvServerChannel>,
}

#[derive(Debug, Deserialize)]
struct TvServerChannel {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    url: Option<String>,
    #[serde(default)]
    dash_url: Option<String>,
    #[serde(default)]
    group_title: Option<String>,
    #[serde(default)]
    stream_resolver: Option<String>,
    #[serde(default)]
    resolver_endpoint: Option<String>,
    #[serde(default)]
    resolver_realm: Option<String>,
    #[serde(default)]
    resolver_channel_id: Option<String>,
    #[serde(default)]
    tvg_id: Option<String>,
    #[serde(default)]
    tvg_name: Option<String>,
    #[serde(default)]
    http_user_agent: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
struct TvChannelsCache {
    saved_at_unix: u64,
    channels: Vec<TvChannel>,
}

pub(crate) fn load_current_programs(
    channels: &[TvChannel],
) -> Result<HashMap<String, TvProgram>, String> {
    let now = Local::now();
    let timeline = load_timeline_programs_for_date(now.date_naive())?;
    let programs_by_channel = timeline.programs_by_channel;
    let now_seconds = now.timestamp();
    let mut current_programs = HashMap::new();
    for (channel, programs) in programs_by_channel {
        let current = current_program_at(&programs, now_seconds);
        if let Some(program) = current {
            current_programs.insert(channel, program);
        }
    }
    refresh_missing_or_expired_current_programs_from_channel_guides(
        channels,
        now.date_naive(),
        now_seconds,
        &timeline.exact_channel_names,
        &mut current_programs,
    );
    Ok(current_programs)
}

fn refresh_missing_or_expired_current_programs_from_channel_guides(
    channels: &[TvChannel],
    date: NaiveDate,
    now_seconds: i64,
    exact_channel_names: &HashMap<String, Vec<String>>,
    current_programs: &mut HashMap<String, TvProgram>,
) {
    // The timeline endpoint is intentionally compact and sometimes omits a
    // few channels even though their channel-specific guide is populated
    // (currently Boing and Italia 2 are examples).  Only enable the more
    // expensive per-channel fallback for provider groups that are actually
    // covered by the timeline.  This keeps regional/local channels, for which
    // the guide service has no data, from causing hundreds of empty requests.
    let mut category_coverage = HashMap::<&str, (usize, usize)>::new();
    for channel in channels {
        let coverage = category_coverage
            .entry(channel.category.as_str())
            .or_insert((0, 0));
        coverage.0 += 1;
        if current_program_for_channel(current_programs, channel).is_some() {
            coverage.1 += 1;
        }
    }

    let mut refreshed_keys = std::collections::HashSet::new();
    for channel in channels {
        let Some(&(category_channels, covered_channels)) =
            category_coverage.get(channel.category.as_str())
        else {
            continue;
        };
        if covered_channels == 0 || covered_channels.saturating_mul(2) < category_channels {
            continue;
        }

        let keys = guide_lookup_keys(channel);
        let Some(key) = keys
            .iter()
            .find(|key| current_programs.contains_key(*key))
            .cloned()
            .or_else(|| keys.into_iter().next())
        else {
            continue;
        };
        let is_expired = current_programs
            .get(&key)
            .is_some_and(|program| program.end_time <= now_seconds);
        let is_missing = !current_programs.contains_key(&key);
        if (!is_missing && !is_expired) || !refreshed_keys.insert(key.clone()) {
            continue;
        }

        let mut channel_variants = exact_channel_names.get(&key).cloned().unwrap_or_default();
        let requested_channel = guide_channel_name(channel).to_string();
        if !channel_variants
            .iter()
            .any(|variant| variant.eq_ignore_ascii_case(&requested_channel))
        {
            channel_variants.push(requested_channel.clone());
        }
        let dtt_channel = format!("{requested_channel} (DTT)");
        if !channel_variants
            .iter()
            .any(|variant| variant.eq_ignore_ascii_case(&dtt_channel))
        {
            channel_variants.push(dtt_channel);
        }
        if normalize_channel_name(&requested_channel) == "cine34"
            && !channel_variants
                .iter()
                .any(|variant| variant.eq_ignore_ascii_case("Cine 34"))
        {
            channel_variants.push("Cine 34".to_string());
        }

        let mut refreshed_program = None;
        for exact_channel in channel_variants {
            match load_channel_guide_for_exact_name(&exact_channel, date) {
                Ok(programs) => {
                    let Some(program) = current_program_at(&programs, now_seconds) else {
                        continue;
                    };
                    crate::log_debug(&format!(
                        "TV current programme refreshed from channel guide: channel={:?} exact={:?} missing={} program={:?}",
                        channel.name, exact_channel, is_missing, program.title
                    ));
                    refreshed_program = Some(program);
                    break;
                }
                Err(err) => crate::log_debug(&format!(
                    "TV current programme channel-guide variant failed: channel={:?} exact={:?} error={err}",
                    channel.name, exact_channel
                )),
            }
        }
        if let Some(program) = refreshed_program {
            current_programs.insert(key, program);
        }
    }
}

fn load_timeline_programs_for_date(date: NaiveDate) -> Result<TvTimelineData, String> {
    let cache = TV_TIMELINE_CACHE.get_or_init(|| Mutex::new(None));
    if let Ok(cache) = cache.lock()
        && let Some(cached) = cache.as_ref()
        && cached.date == date
        && cached.fetched_at.elapsed() < TV_GUIDE_CACHE_TTL
    {
        crate::log_debug(&format!("TV guide timeline cache hit: date={date}"));
        return Ok(TvTimelineData {
            programs_by_channel: cached.programs_by_channel.clone(),
            exact_channel_names: cached.exact_channel_names.clone(),
        });
    }

    let template = decode_tv_guide_timeline_url()?;
    let date_text = date.format("%Y-%m-%d").to_string();
    let url = template.replace("{date}", &date_text);

    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.guide_init", &[("error", &err.to_string())])
        })?;
    let response = client.get(url).send().map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.guide_download", &[("error", &err.to_string())])
    })?;
    let status = response.status();
    if !status.is_success() {
        let status_text = status.as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.guide_http",
            &[("status", &status_text)],
        ));
    }
    let root: Value = response.json().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.guide_invalid_response",
            &[("error", &err.to_string())],
        )
    })?;
    crate::log_debug(&format!("TV guide timeline API: date={date}"));
    let mut programs_by_channel = HashMap::<String, Vec<TvProgram>>::new();
    let exact_channel_names =
        collect_tv_guide_programs_from_timeline_root(&root, &mut programs_by_channel);
    for programs in programs_by_channel.values_mut() {
        programs.sort_by_key(|program| program.start_time);
        programs.dedup_by(|left, right| {
            left.start_time == right.start_time
                && left.end_time == right.end_time
                && left.title == right.title
        });
    }
    if let Ok(mut cache) = cache.lock() {
        *cache = Some(CachedTvTimeline {
            date,
            fetched_at: Instant::now(),
            programs_by_channel: programs_by_channel.clone(),
            exact_channel_names: exact_channel_names.clone(),
        });
    }
    Ok(TvTimelineData {
        programs_by_channel,
        exact_channel_names,
    })
}

pub(crate) fn load_channel_guide(
    channel: &TvChannel,
    date: NaiveDate,
) -> Result<Vec<TvProgram>, String> {
    let date_text = date.format("%Y-%m-%d").to_string();
    let requested_channel = guide_channel_name(channel);
    let requested_normalized = normalize_channel_name(requested_channel);
    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.guide_init", &[("error", &err.to_string())])
        })?;

    let exact_channel =
        resolve_exact_guide_channel_name(&client, &date_text, &requested_normalized)
            .unwrap_or_else(|| requested_channel.to_string());

    load_channel_guide_for_exact_name(&exact_channel, date)
}

fn load_channel_guide_for_exact_name(
    exact_channel: &str,
    date: NaiveDate,
) -> Result<Vec<TvProgram>, String> {
    let cache_key = format!(
        "{}\0{}",
        date.format("%Y-%m-%d"),
        exact_channel.trim().to_ascii_lowercase()
    );
    let cache = TV_CHANNEL_GUIDE_CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    if let Ok(cache) = cache.lock()
        && let Some(cached) = cache.get(&cache_key)
        && cached.fetched_at.elapsed() < TV_GUIDE_CACHE_TTL
    {
        crate::log_debug(&format!(
            "TV guide channel cache hit: exact={:?} date={} programs={}",
            exact_channel,
            date,
            cached.programs.len()
        ));
        return Ok(cached.programs.clone());
    }

    let date_text = date.format("%Y-%m-%d").to_string();
    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.guide_init", &[("error", &err.to_string())])
        })?;
    let template = decode_tv_guide_channel_url()?;
    let encoded_channel = encode_uri_component(exact_channel);
    let url = template
        .replace("{channel}", &encoded_channel)
        .replace("{date}", &date_text);
    let response = client.get(url).send().map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.guide_download", &[("error", &err.to_string())])
    })?;
    let status = response.status();
    if !status.is_success() {
        let status_text = status.as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.guide_http",
            &[("status", &status_text)],
        ));
    }

    let root: Value = response.json().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.guide_invalid_response",
            &[("error", &err.to_string())],
        )
    })?;
    let items = root
        .as_array()
        .ok_or_else(|| crate::i18n::tr_tv("tv.error.guide_missing_programs"))?;
    let mut programs = Vec::new();
    for item in items {
        let Some(object) = item.as_object() else {
            continue;
        };
        let title = object
            .get("title")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .trim();
        if title.is_empty() {
            continue;
        }
        let start_time = read_json_i64(object, "startTime", "start_time");
        let end_time = read_json_i64(object, "endTime", "end_time");
        if start_time <= 0 || end_time <= start_time {
            continue;
        }
        programs.push(TvProgram {
            title: title.to_string(),
            start_time,
            end_time,
        });
    }
    programs.sort_by_key(|program| program.start_time);
    programs.dedup_by(|left, right| {
        left.start_time == right.start_time
            && left.end_time == right.end_time
            && left.title == right.title
    });
    crate::log_debug(&format!(
        "TV guide channel API: exact={:?} date={} programs={}",
        exact_channel,
        date,
        programs.len()
    ));
    if let Ok(mut cache) = cache.lock() {
        cache.insert(
            cache_key,
            CachedTvChannelGuide {
                fetched_at: Instant::now(),
                programs: programs.clone(),
            },
        );
    }
    Ok(programs)
}

fn resolve_exact_guide_channel_name(
    client: &Client,
    date_text: &str,
    target_normalized: &str,
) -> Option<String> {
    if target_normalized.is_empty() {
        return None;
    }
    let template = decode_tv_guide_timeline_url().ok()?;
    let url = template.replace("{date}", date_text);
    let response = client.get(url).send().ok()?;
    if !response.status().is_success() {
        return None;
    }
    let root: Value = response.json().ok()?;
    let groups = root.as_array()?;
    for group in groups {
        let Some(items) = group.as_array() else {
            continue;
        };
        for item in items {
            let Some(object) = item.as_object() else {
                continue;
            };
            let channel_name = object
                .get("ch")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .trim();
            if normalize_channel_name(channel_name) == target_normalized {
                return Some(channel_name.to_string());
            }
        }
    }
    None
}

fn guide_channel_name(channel: &TvChannel) -> &str {
    let tvg_name = channel.tvg_name.trim();
    if tvg_name.is_empty() {
        channel.name.trim()
    } else {
        tvg_name
    }
}

fn encode_uri_component(value: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut encoded = String::with_capacity(value.len());
    for byte in value.bytes() {
        if byte.is_ascii_alphanumeric()
            || matches!(
                byte,
                b'-' | b'_' | b'.' | b'!' | b'~' | b'*' | b'\'' | b'(' | b')'
            )
        {
            encoded.push(char::from(byte));
        } else {
            encoded.push('%');
            encoded.push(char::from(HEX[usize::from(byte >> 4)]));
            encoded.push(char::from(HEX[usize::from(byte & 0x0f)]));
        }
    }
    encoded
}

fn collect_tv_guide_programs_from_timeline_root(
    root: &Value,
    programs_by_channel: &mut HashMap<String, Vec<TvProgram>>,
) -> HashMap<String, Vec<String>> {
    let Some(groups) = root.as_array() else {
        return HashMap::new();
    };
    // The timeline can contain multiple schedules whose names normalize to the
    // same Sonarpad channel (for example "Italia 1 (DTT)" and "Italia 1").
    // The channel-specific guide deliberately selects the first exact variant
    // returned by this timeline.  Keep that same variant here instead of
    // merging conflicting schedules and picking an arbitrary current program.
    let mut exact_channel_names = HashMap::<String, Vec<String>>::new();
    // Resolve the exact variant in a separate first pass, just like
    // `resolve_exact_guide_channel_name`.  Some timeline rows identify the
    // channel but do not yet contain a usable programme; they must still
    // determine which variant the channel-specific guide would select.
    for group in groups {
        let Some(items) = group.as_array() else {
            continue;
        };
        for item in items {
            let Some(object) = item.as_object() else {
                continue;
            };
            let guide_channel = object
                .get("ch")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .trim();
            let key = normalize_channel_name(guide_channel);
            if !key.is_empty() {
                let variants = exact_channel_names.entry(key).or_default();
                if !variants
                    .iter()
                    .any(|variant| variant.eq_ignore_ascii_case(guide_channel))
                {
                    variants.push(guide_channel.to_string());
                }
            }
        }
    }
    for group in groups {
        let Some(items) = group.as_array() else {
            continue;
        };
        for item in items {
            let Some(object) = item.as_object() else {
                continue;
            };
            let guide_channel = object
                .get("ch")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .trim();
            let title = object
                .get("title")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .trim();
            if guide_channel.is_empty() || title.is_empty() {
                continue;
            }
            let start_time = read_json_i64(object, "startTime", "start_time");
            let end_time = read_json_i64(object, "endTime", "end_time");
            if start_time <= 0 || end_time <= 0 {
                continue;
            }
            let key = normalize_channel_name(guide_channel);
            if key.is_empty() {
                continue;
            }
            if exact_channel_names
                .get(&key)
                .and_then(|variants| variants.first())
                .map(String::as_str)
                != Some(guide_channel)
            {
                continue;
            }
            programs_by_channel.entry(key).or_default().push(TvProgram {
                title: title.to_string(),
                start_time,
                end_time,
            });
        }
    }
    exact_channel_names
}

pub(crate) fn current_program_for_channel<'a>(
    current_programs: &'a HashMap<String, TvProgram>,
    channel: &TvChannel,
) -> Option<&'a TvProgram> {
    guide_lookup_keys(channel)
        .into_iter()
        .find_map(|key| current_programs.get(&key))
}

fn guide_lookup_keys(channel: &TvChannel) -> Vec<String> {
    let mut keys = Vec::new();
    for value in [&channel.name, &channel.tvg_name, &channel.tvg_id] {
        push_normalized_guide_key(&mut keys, value);
    }
    if channel.tvg_id.to_ascii_lowercase().ends_with(".it") {
        let without_suffix = &channel.tvg_id[..channel.tvg_id.len().saturating_sub(3)];
        push_normalized_guide_key(&mut keys, without_suffix);
    }
    keys
}

fn push_normalized_guide_key(keys: &mut Vec<String>, value: &str) {
    let normalized = normalize_channel_name(value);
    if !normalized.is_empty() && !keys.contains(&normalized) {
        keys.push(normalized);
    }
}

fn latest_started_program(programs: &[TvProgram], now_seconds: i64) -> Option<TvProgram> {
    let latest = programs
        .iter()
        .filter(|program| program.start_time <= now_seconds)
        .max_by_key(|program| program.start_time)?;
    (now_seconds - latest.start_time <= TV_GUIDE_FALLBACK_MAX_AGE_SECS).then(|| latest.clone())
}

fn current_program_at(programs: &[TvProgram], now_seconds: i64) -> Option<TvProgram> {
    // EPG feeds occasionally contain overlapping entries.  Once a newer
    // programme has started it must supersede an older entry whose end time
    // incorrectly extends beyond the real changeover.
    programs
        .iter()
        .filter(|program| program.start_time <= now_seconds && program.end_time > now_seconds)
        .max_by_key(|program| program.start_time)
        .cloned()
        .or_else(|| latest_started_program(programs, now_seconds))
}

fn read_json_i64(object: &serde_json::Map<String, Value>, camel_key: &str, snake_key: &str) -> i64 {
    let Some(value) = object.get(camel_key).or_else(|| object.get(snake_key)) else {
        return 0;
    };
    value
        .as_i64()
        .or_else(|| value.as_u64().and_then(|number| i64::try_from(number).ok()))
        .or_else(|| value.as_str().and_then(|text| text.parse::<i64>().ok()))
        .unwrap_or(0)
}

fn decode_tv_guide_timeline_url() -> Result<String, String> {
    decode_tv_guide_payload(TV_GUIDE_TIMELINE_PAYLOAD_JSON)
}

fn decode_tv_guide_channel_url() -> Result<String, String> {
    decode_tv_guide_payload(TV_GUIDE_CHANNEL_PAYLOAD_JSON)
}

fn decode_tv_guide_payload(payload_json: &str) -> Result<String, String> {
    let secret = crate::settings::load_saved_rai_luce_code()
        .ok_or_else(|| crate::i18n::tr_tv("tv.error.guide_code_missing"))?;
    let secret = secret.trim();
    if secret.is_empty() {
        return Err(crate::i18n::tr_tv("tv.error.guide_code_invalid"));
    }
    let payload: EncryptedTvGuidePayload = serde_json::from_str(payload_json).map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.guide_payload_invalid",
            &[("error", &err.to_string())],
        )
    })?;
    if payload.algorithm != "gzip-xor-base64-v1" {
        return Err(crate::i18n::tr_tv_f(
            "tv.error.guide_algorithm_unsupported",
            &[("algorithm", &payload.algorithm)],
        ));
    }

    let encrypted = base64::engine::general_purpose::STANDARD
        .decode(payload.payload_b64)
        .map_err(|err| {
            crate::i18n::tr_tv_f(
                "tv.error.guide_base64_invalid",
                &[("error", &err.to_string())],
            )
        })?;
    let mut key = secret.as_bytes().to_vec();
    for part in TV_GUIDE_STATIC_KEY_PARTS {
        key.extend_from_slice(part);
    }
    if key.is_empty() {
        return Err(crate::i18n::tr_tv("tv.error.guide_key_invalid"));
    }
    let decrypted = encrypted
        .iter()
        .enumerate()
        .map(|(index, byte)| byte ^ key[index % key.len()])
        .collect::<Vec<_>>();
    let mut decoder = GzDecoder::new(decrypted.as_slice());
    let mut decoded = String::new();
    decoder.read_to_string(&mut decoded).map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.guide_gzip_invalid",
            &[("error", &err.to_string())],
        )
    })?;
    if decoded.trim().is_empty() {
        return Err(crate::i18n::tr_tv("tv.error.guide_url_empty"));
    }
    Ok(decoded)
}

pub(crate) fn normalize_channel_name(name: &str) -> String {
    let mut value = name.trim().to_ascii_lowercase();
    if let Some(rest) = value.strip_prefix('[')
        && let Some(end) = rest.find(']')
        && rest[..end].chars().all(|ch| ch.is_ascii_digit())
    {
        value = rest[end + 1..].trim().to_string();
    }
    value = value
        .replace("(dtt)", "")
        .replace(" dtt", "")
        .replace(" hd", "")
        .replace("twenty seven", "27")
        .replace("twentyseven", "27");
    let mut normalized = value
        .chars()
        .filter(|ch| ch.is_ascii_alphanumeric())
        .collect::<String>();
    if normalized.ends_with("hd") {
        normalized.truncate(normalized.len() - 2);
    }
    match normalized.as_str() {
        "la7dtt" => "la7".to_string(),
        "mediaset20" | "20mediaset" => "20".to_string(),
        "mediaset27" | "27mediaset" => "27".to_string(),
        "retequattro" | "rete4mediaset" | "mediasetrete4" => "rete4".to_string(),
        "canale5mediaset" | "mediasetcanale5" => "canale5".to_string(),
        "italia1mediaset" | "mediasetitalia1" => "italia1".to_string(),
        "italia2mediaset" | "mediasetitalia2" => "italia2".to_string(),
        "focustv" | "tvfocus" | "mediasetfocus" | "focusmediaset" => "focus".to_string(),
        "sportitalialive24" => "sportitalia".to_string(),
        "virginradio" => "virginradiotv".to_string(),
        _ if normalized.contains("rete4") || normalized.contains("retequattro") => {
            "rete4".to_string()
        }
        _ => normalized,
    }
}

pub(crate) fn load_channels_with_cache() -> Result<TvChannelLoadResult, String> {
    match load_channels_from_server() {
        Ok(channels) if !channels.is_empty() => {
            if let Err(err) = write_channels_cache(&channels) {
                crate::log_debug(&format!("TV channel cache write failed: {err}"));
            }
            Ok(TvChannelLoadResult {
                channels,
                cache_warning: None,
            })
        }
        Ok(_) => load_channels_from_cache(&crate::i18n::tr_tv("tv.warning.server_empty_use_cache")),
        Err(err) => {
            crate::log_debug(&format!("TV channel server request failed: {err}"));
            load_channels_from_cache(&crate::i18n::tr_tv(
                "tv.warning.server_unreachable_use_cache",
            ))
            .map_err(|_| err)
        }
    }
}

fn load_channels_from_server() -> Result<Vec<TvChannel>, String> {
    if TV_API_CLIENT_TOKEN.trim().is_empty() {
        return Err(crate::i18n::tr_tv("tv.error.token_missing"));
    }

    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.connection_init", &[("error", &err.to_string())])
        })?;
    let response = client
        .get(TV_CHANNELS_URL)
        .header(TV_API_TOKEN_HEADER, TV_API_CLIENT_TOKEN)
        .header(ROUTE_API_TOKEN_HEADER, TV_API_CLIENT_TOKEN)
        .send()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.list_download", &[("error", &err.to_string())])
        })?;

    let status = response.status();
    if status.as_u16() == 401 || status.as_u16() == 403 {
        return Err(crate::i18n::tr_tv("tv.error.client_unauthorized"));
    }
    if !status.is_success() {
        let status_text = status.as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.list_http",
            &[("status", &status_text)],
        ));
    }

    let payload: TvServerResponse = response.json().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.list_invalid_response",
            &[("error", &err.to_string())],
        )
    })?;
    let mut channels = Vec::new();
    for raw in payload.channels {
        if let Some(channel) = normalize_server_channel(raw) {
            channels.push(channel);
        }
    }
    Ok(channels)
}

fn normalize_server_channel(raw: TvServerChannel) -> Option<TvChannel> {
    let name = strip_numeric_prefix(raw.name.as_deref().unwrap_or_default().trim());
    let mut url = raw.url.as_deref().unwrap_or_default().trim().to_string();
    if name.eq_ignore_ascii_case("La7") {
        url = LA7_STREAM_URL.to_string();
    } else if name.eq_ignore_ascii_case("La7 Cinema") || name.eq_ignore_ascii_case("La7D") {
        url = LA7_CINEMA_DASH_URL.to_string();
    }
    if name.is_empty() || url.is_empty() {
        return None;
    }

    let group_title = raw.group_title.as_deref().unwrap_or_default().trim();
    let category = if group_title.is_empty() {
        crate::i18n::tr_tv("tv.category.other")
    } else {
        group_title.to_string()
    };

    Some(TvChannel {
        name,
        url,
        dash_url: trimmed_option(raw.dash_url),
        category,
        stream_resolver: trimmed_option(raw.stream_resolver),
        resolver_endpoint: trimmed_option(raw.resolver_endpoint),
        resolver_realm: trimmed_option(raw.resolver_realm),
        resolver_channel_id: trimmed_option(raw.resolver_channel_id),
        tvg_id: raw.tvg_id.as_deref().unwrap_or_default().trim().to_string(),
        tvg_name: raw
            .tvg_name
            .as_deref()
            .unwrap_or_default()
            .trim()
            .to_string(),
        http_user_agent: raw
            .http_user_agent
            .as_deref()
            .unwrap_or_default()
            .trim()
            .to_string(),
    })
}

fn trimmed_option(value: Option<String>) -> Option<String> {
    value
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
}

fn strip_numeric_prefix(value: &str) -> String {
    let trimmed = value.trim();
    let Some(rest) = trimmed.strip_prefix('[') else {
        return trimmed.to_string();
    };
    let Some(end) = rest.find(']') else {
        return trimmed.to_string();
    };
    if rest[..end].chars().all(|ch| ch.is_ascii_digit()) {
        rest[end + 1..].trim().to_string()
    } else {
        trimmed.to_string()
    }
}

fn channels_cache_path() -> std::path::PathBuf {
    crate::settings::settings_dir().join("tv_channels_cache.json")
}

fn write_channels_cache(channels: &[TvChannel]) -> Result<(), String> {
    let path = channels_cache_path();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.cache_directory", &[("error", &err.to_string())])
        })?;
    }
    let saved_at_unix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let cache = TvChannelsCache {
        saved_at_unix,
        channels: channels.to_vec(),
    };
    let bytes = serde_json::to_vec(&cache).map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.cache_prepare", &[("error", &err.to_string())])
    })?;
    fs::write(path, bytes)
        .map_err(|err| crate::i18n::tr_tv_f("tv.error.cache_save", &[("error", &err.to_string())]))
}

fn load_channels_from_cache(warning: &str) -> Result<TvChannelLoadResult, String> {
    let path = channels_cache_path();
    let bytes = fs::read(path).map_err(|_| crate::i18n::tr_tv("tv.error.cache_missing"))?;
    let cache: TvChannelsCache = serde_json::from_slice(&bytes).map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.cache_invalid", &[("error", &err.to_string())])
    })?;
    if cache.channels.is_empty() {
        return Err(crate::i18n::tr_tv("tv.error.cache_empty"));
    }
    let _cache_saved_at_unix = cache.saved_at_unix;
    Ok(TvChannelLoadResult {
        channels: cache.channels,
        cache_warning: Some(warning.to_string()),
    })
}

pub(crate) fn resolve_stream_url(channel: &TvChannel) -> Result<String, String> {
    let mut resolver = channel.stream_resolver.as_deref();
    let mut resolver_channel_id = channel.resolver_channel_id.as_deref();
    let normalized_name = normalize_channel_name(&channel.name);

    if let Some(discovery_id) = discovery_channel_id(&normalized_name) {
        resolver = Some("aurora_channel");
        if resolver_channel_id.is_none() {
            resolver_channel_id = Some(discovery_id);
        }
    }

    if resolver == Some("aurora_channel")
        && let Some(channel_id) = resolver_channel_id
    {
        match resolve_aurora_channel(channel, channel_id) {
            Ok(url) => return Ok(url),
            Err(err) => crate::log_debug(&format!(
                "TV Aurora resolver failed for {}: {}",
                channel.name, err
            )),
        }
    }

    if channel.url.contains("/relinker/relinkerServlet") {
        return resolve_rai_relinker(channel);
    }

    if let Some(dash_url) = channel
        .dash_url
        .as_deref()
        .map(str::trim)
        .filter(|url| !url.is_empty())
    {
        crate::log_debug(&format!(
            "TV direct stream: using DASH alternative for {}",
            channel.name
        ));
        return Ok(dash_url.to_string());
    }

    Ok(channel.url.trim().to_string())
}

fn discovery_channel_id(normalized_name: &str) -> Option<&'static str> {
    match normalized_name {
        "realtime" => Some("2"),
        "nove" | "la9" | "9" => Some("3"),
        "dmax" => Some("4"),
        "foodnetwork" => Some("6"),
        "motortrend" => Some("11"),
        "discoverychannel" => Some("12"),
        "hgtv" => Some("13"),
        "k2" => Some("24"),
        "frisbee" => Some("26"),
        "giallo" | "giallotv" => Some("27"),
        _ => None,
    }
}

fn resolve_aurora_channel(channel: &TvChannel, channel_id: &str) -> Result<String, String> {
    let endpoint = channel
        .resolver_endpoint
        .as_deref()
        .unwrap_or("https://public.aurora.enhanced.live")
        .trim_end_matches('/');
    let realm = channel.resolver_realm.as_deref().unwrap_or("it");
    let client = Client::builder()
        .user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")
        .timeout(Duration::from_secs(15))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.aurora_init", &[("error", &err.to_string())])
        })?;
    let mut token_url = Url::parse(&format!("{endpoint}/token")).map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.aurora_token_url", &[("error", &err.to_string())])
    })?;
    token_url.query_pairs_mut().append_pair("realm", realm);
    let referer = channel.url.trim();

    let token_response = add_aurora_headers(client.get(token_url), realm, referer)
        .send()
        .map_err(|err| {
            crate::i18n::tr_tv_f(
                "tv.error.aurora_token_request",
                &[("error", &err.to_string())],
            )
        })?;
    if !token_response.status().is_success() {
        let status_text = token_response.status().as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.aurora_token_http",
            &[("status", &status_text)],
        ));
    }
    let token_json: Value = token_response.json().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.aurora_token_response",
            &[("error", &err.to_string())],
        )
    })?;
    let token = token_json
        .pointer("/data/attributes/token")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| crate::i18n::tr_tv("tv.error.aurora_token_missing"))?;

    let playback_url = format!("{endpoint}/playback/v3/channelPlaybackInfo");
    let body = serde_json::json!({
        "channelId": channel_id,
        "deviceInfo": {
            "adBlocker": false,
            "drmSupported": true,
            "hdrCapabilities": ["SDR"],
            "hwDecodingCapabilities": [],
            "soundCapabilities": ["STEREO"]
        },
        "wisteriaProperties": {
            "device": {
                "browser": {"name": "chrome", "version": "136"},
                "type": "desktop"
            },
            "platform": "desktop"
        }
    });
    let playback_response = add_aurora_headers(client.post(playback_url), realm, referer)
        .bearer_auth(token)
        .json(&body)
        .send()
        .map_err(|err| {
            crate::i18n::tr_tv_f(
                "tv.error.aurora_playback_request",
                &[("error", &err.to_string())],
            )
        })?;
    if !playback_response.status().is_success() {
        let status_text = playback_response.status().as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.aurora_playback_http",
            &[("status", &status_text)],
        ));
    }
    let playback_json: Value = playback_response.json().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.aurora_playback_response",
            &[("error", &err.to_string())],
        )
    })?;
    find_stream_url(&playback_json, ".m3u8")
        .ok_or_else(|| crate::i18n::tr_tv("tv.error.aurora_hls_missing"))
}

fn add_aurora_headers(
    request: reqwest::blocking::RequestBuilder,
    realm: &str,
    referer: &str,
) -> reqwest::blocking::RequestBuilder {
    request
        .header("Accept", "application/json,text/plain,*/*")
        .header("Content-Type", "application/json")
        .header("Origin", "https://nove.tv")
        .header("Referer", referer)
        .header("X-disco-client", "WEB:UNKNOWN:wbdatv:2.1.9")
        .header("X-disco-params", format!("realm={realm}"))
        .header(
            "X-Device-Info",
            "STONEJS/1 (Unknown/Unknown; Windows/10; Unknown)",
        )
}

fn find_stream_url(value: &Value, extension: &str) -> Option<String> {
    match value {
        Value::String(text) => {
            let trimmed = text.trim();
            (trimmed.starts_with("http://") || trimmed.starts_with("https://"))
                .then(|| trimmed.to_string())
                .filter(|url| url.to_ascii_lowercase().contains(extension))
        }
        Value::Array(values) => values
            .iter()
            .find_map(|value| find_stream_url(value, extension)),
        Value::Object(values) => values
            .values()
            .find_map(|value| find_stream_url(value, extension)),
        _ => None,
    }
}

fn resolve_rai_relinker(channel: &TvChannel) -> Result<String, String> {
    let mut url = Url::parse(channel.url.trim()).map_err(|err| {
        crate::i18n::tr_tv_f("tv.error.rai_relinker_url", &[("error", &err.to_string())])
    })?;
    let mut query = url
        .query_pairs()
        .filter(|(key, _)| key != "output")
        .map(|(key, value)| (key.into_owned(), value.into_owned()))
        .collect::<Vec<_>>();
    query.push(("output".to_string(), "54".to_string()));
    {
        let mut query_pairs = url.query_pairs_mut();
        query_pairs.clear();
        for (key, value) in query {
            query_pairs.append_pair(&key, &value);
        }
    }

    let client = Client::builder()
        .user_agent(channel.playback_user_agent())
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| {
            crate::i18n::tr_tv_f("tv.error.rai_relinker_init", &[("error", &err.to_string())])
        })?;
    let response = client
        .get(url)
        .header("Origin", "https://www.raiplay.it")
        .header("Referer", "https://www.raiplay.it/")
        .send()
        .map_err(|err| {
            crate::i18n::tr_tv_f(
                "tv.error.rai_stream_resolve",
                &[("error", &err.to_string())],
            )
        })?;
    if !response.status().is_success() {
        let status_text = response.status().as_u16().to_string();
        return Err(crate::i18n::tr_tv_f(
            "tv.error.rai_relinker_http",
            &[("status", &status_text)],
        ));
    }
    let final_url = response.url().to_string();
    let body = response.text().map_err(|err| {
        crate::i18n::tr_tv_f(
            "tv.error.rai_relinker_response",
            &[("error", &err.to_string())],
        )
    })?;
    let trimmed = body.trim();
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        return Ok(trimmed.to_string());
    }
    if trimmed.starts_with("#EXTM3U") {
        return Ok(final_url);
    }
    extract_xml_url(trimmed, true)
        .or_else(|| extract_xml_url(trimmed, false))
        .ok_or_else(|| crate::i18n::tr_tv("tv.error.rai_stream_missing"))
}

fn extract_xml_url(xml: &str, require_content_type: bool) -> Option<String> {
    let mut offset = 0usize;
    while let Some(relative_start) = xml[offset..].find("<url") {
        let start = offset + relative_start;
        let tag_end = xml[start..].find('>')? + start;
        let tag = &xml[start..=tag_end];
        if require_content_type && !tag.contains("type=\"content\"") {
            offset = tag_end.saturating_add(1);
            continue;
        }
        let value_start = tag_end + 1;
        let value_end = xml[value_start..].find("</url>")? + value_start;
        let value = xml[value_start..value_end].trim();
        if !value.is_empty() {
            return Some(value.to_string());
        }
        offset = value_end.saturating_add("</url>".len());
    }
    None
}

pub(crate) fn is_rai_audio_description_channel(channel: &TvChannel) -> bool {
    channel.category.trim().eq_ignore_ascii_case("rai")
        && channel.name.trim().to_ascii_lowercase().starts_with("rai")
}

pub(crate) fn matches_search(channel: &TvChannel, query: &str) -> bool {
    let query = normalize_search_text(query);
    !query.is_empty()
        && [
            channel.name.as_str(),
            channel.category.as_str(),
            channel.tvg_name.as_str(),
            channel.tvg_id.as_str(),
        ]
        .iter()
        .any(|value| normalize_search_text(value).contains(&query))
}

fn normalize_search_text(value: &str) -> String {
    value
        .to_lowercase()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_numeric_prefix_only_when_numeric() {
        assert_eq!(strip_numeric_prefix("[12] Rai 1"), "Rai 1");
        assert_eq!(strip_numeric_prefix("[HD] Rai 1"), "[HD] Rai 1");
    }

    #[test]
    fn normalizes_common_channel_aliases() {
        assert_eq!(normalize_channel_name("[20] Rete Quattro HD"), "rete4");
        assert_eq!(normalize_channel_name("Twenty Seven"), "27");
        assert_eq!(normalize_channel_name("Focus Tv"), "focus");
    }

    #[test]
    fn extracts_relinker_xml_content_url() {
        let xml = r#"<root><url type="content">https://example.test/live.m3u8</url></root>"#;
        assert_eq!(
            extract_xml_url(xml, true).as_deref(),
            Some("https://example.test/live.m3u8")
        );
    }

    #[test]
    fn all_rai_category_channels_try_audiodescription() {
        let channel = TvChannel {
            name: "Rai 4".to_string(),
            url: "https://example.test/direct/master.m3u8".to_string(),
            dash_url: None,
            category: "Rai".to_string(),
            stream_resolver: None,
            resolver_endpoint: None,
            resolver_realm: None,
            resolver_channel_id: None,
            tvg_id: String::new(),
            tvg_name: String::new(),
            http_user_agent: String::new(),
        };

        assert!(is_rai_audio_description_channel(&channel));
    }

    #[test]
    fn direct_channel_prefers_server_dash_url() {
        let raw: TvServerChannel = serde_json::from_value(serde_json::json!({
            "name": "[6] Italia 1",
            "url": "https://live02-seg.example.test/live/ch-i1/index.m3u8",
            "dash_url": "  https://live03-col.example.test/live/ch-i1/manifest.mpd  ",
            "group_title": "Mediaset"
        }))
        .expect("server channel should deserialize");
        let channel = normalize_server_channel(raw).expect("channel should normalize");

        assert_eq!(channel.name, "Italia 1");
        assert_eq!(
            channel.dash_url.as_deref(),
            Some("https://live03-col.example.test/live/ch-i1/manifest.mpd")
        );
        assert_eq!(
            resolve_stream_url(&channel).expect("direct stream should resolve"),
            "https://live03-col.example.test/live/ch-i1/manifest.mpd"
        );
    }

    #[test]
    fn timeline_parser_ignores_nested_stale_programs() {
        let root = serde_json::json!([
            [
                {
                    "ch": "Italia 2",
                    "title": "Occhi di gatto",
                    "startTime": 100,
                    "endTime": 200,
                    "metadata": {
                        "ch": "Italia 2",
                        "title": "Che campioni Holly e Benji!",
                        "startTime": 50,
                        "endTime": 250
                    }
                }
            ]
        ]);
        let mut programs = HashMap::new();

        collect_tv_guide_programs_from_timeline_root(&root, &mut programs);

        let italy_two = programs
            .get("italia2")
            .expect("Italia 2 should be collected from the main timeline row");
        assert_eq!(italy_two.len(), 1);
        assert_eq!(italy_two[0].title, "Occhi di gatto");
    }

    #[test]
    fn timeline_parser_does_not_merge_conflicting_channel_variants() {
        let root = serde_json::json!([
            [
                {
                    "ch": "Italia 1 (DTT)",
                    "title": "Transatlantici",
                    "startTime": 100,
                    "endTime": 200
                },
                {
                    "ch": "Italia 1",
                    "title": "Ramses",
                    "startTime": 100,
                    "endTime": 200
                },
                {
                    "ch": "Italia 1 (DTT)",
                    "title": "Programma successivo",
                    "startTime": 200,
                    "endTime": 300
                }
            ]
        ]);
        let mut programs = HashMap::new();

        collect_tv_guide_programs_from_timeline_root(&root, &mut programs);

        let italy_one = programs
            .get("italia1")
            .expect("Italia 1 should use the first exact timeline variant");
        assert_eq!(italy_one.len(), 2);
        assert_eq!(italy_one[0].title, "Transatlantici");
        assert_eq!(italy_one[1].title, "Programma successivo");
    }

    #[test]
    fn current_program_prefers_latest_start_when_guide_entries_overlap() {
        let programs = vec![
            TvProgram {
                title: "Chicago Med".to_string(),
                start_time: 3 * 60,
                end_time: 4 * 60,
            },
            TvProgram {
                title: "Show reel".to_string(),
                start_time: 3 * 60 + 40,
                end_time: 4 * 60 + 15,
            },
        ];

        let current = current_program_at(&programs, 3 * 60 + 50)
            .expect("an overlapping current programme should be selected");

        assert_eq!(current.title, "Show reel");
    }

    #[test]
    fn timeline_variant_selection_matches_guide_even_if_first_row_is_incomplete() {
        let root = serde_json::json!([
            [
                {
                    "ch": "20"
                },
                {
                    "ch": "20 Mediaset",
                    "title": "Chicago Med",
                    "startTime": 180,
                    "endTime": 240
                },
                {
                    "ch": "20",
                    "title": "Show reel",
                    "startTime": 220,
                    "endTime": 255
                }
            ]
        ]);
        let mut programs = HashMap::new();

        collect_tv_guide_programs_from_timeline_root(&root, &mut programs);

        let mediaset_twenty = programs
            .get("20")
            .expect("the exact variant selected by the guide should be retained");
        assert_eq!(mediaset_twenty.len(), 1);
        assert_eq!(mediaset_twenty[0].title, "Show reel");
    }
}
