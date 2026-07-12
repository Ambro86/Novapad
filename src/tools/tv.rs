use base64::Engine;
use chrono::{Local, NaiveDate};
use flate2::read::GzDecoder;
use reqwest::blocking::Client;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;
use std::fs;
use std::io::Read;
use std::time::{Duration, SystemTime, UNIX_EPOCH};
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

#[derive(Clone, Debug, Serialize, Deserialize)]
pub(crate) struct TvChannel {
    pub(crate) name: String,
    pub(crate) url: String,
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

pub(crate) fn load_current_programs() -> Result<HashMap<String, TvProgram>, String> {
    let now = Local::now();
    let programs_by_channel = load_programs_for_date(now.date_naive())?;
    let now_seconds = now.timestamp();
    let mut current_programs = HashMap::new();
    for (channel, programs) in programs_by_channel {
        let current = programs
            .iter()
            .find(|program| program.start_time <= now_seconds && program.end_time > now_seconds)
            .cloned()
            .or_else(|| latest_started_program(&programs, now_seconds));
        if let Some(program) = current {
            current_programs.insert(channel, program);
        }
    }
    Ok(current_programs)
}

pub(crate) fn load_programs_for_date(
    date: NaiveDate,
) -> Result<HashMap<String, Vec<TvProgram>>, String> {
    let template = decode_tv_guide_timeline_url()?;
    let date_text = date.format("%Y-%m-%d").to_string();
    let url = template.replace("{date}", &date_text);

    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| format!("Impossibile inizializzare la guida TV: {err}"))?;
    let response = client
        .get(url)
        .send()
        .map_err(|err| format!("Impossibile scaricare la guida TV: {err}"))?;
    let status = response.status();
    if !status.is_success() {
        return Err(format!(
            "Impossibile scaricare la guida TV: errore HTTP {}.",
            status.as_u16()
        ));
    }
    let root: Value = response
        .json()
        .map_err(|err| format!("Risposta della guida TV non valida: {err}"))?;
    let mut programs_by_channel = HashMap::<String, Vec<TvProgram>>::new();
    collect_tv_guide_programs(&root, &mut programs_by_channel);
    for programs in programs_by_channel.values_mut() {
        programs.sort_by_key(|program| program.start_time);
        programs.dedup_by(|left, right| {
            left.start_time == right.start_time
                && left.end_time == right.end_time
                && left.title == right.title
        });
    }
    Ok(programs_by_channel)
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
        .map_err(|err| format!("Impossibile inizializzare la guida TV: {err}"))?;

    let exact_channel =
        resolve_exact_guide_channel_name(&client, &date_text, &requested_normalized)
            .unwrap_or_else(|| requested_channel.to_string());

    let template = decode_tv_guide_channel_url()?;
    let encoded_channel = encode_uri_component(&exact_channel);
    let url = template
        .replace("{channel}", &encoded_channel)
        .replace("{date}", &date_text);
    let response = client
        .get(url)
        .send()
        .map_err(|err| format!("Impossibile scaricare la guida TV: {err}"))?;
    let status = response.status();
    if !status.is_success() {
        return Err(format!(
            "Impossibile scaricare la guida TV: errore HTTP {}.",
            status.as_u16()
        ));
    }

    let root: Value = response
        .json()
        .map_err(|err| format!("Risposta della guida TV non valida: {err}"))?;
    let items = root.as_array().ok_or_else(|| {
        "Risposta della guida TV non valida: elenco programmi assente.".to_string()
    })?;
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
        "TV guide channel API: requested={:?} exact={:?} date={} programs={}",
        requested_channel,
        exact_channel,
        date,
        programs.len()
    ));
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

fn collect_tv_guide_programs(
    value: &Value,
    programs_by_channel: &mut HashMap<String, Vec<TvProgram>>,
) {
    match value {
        Value::Array(items) => {
            for item in items {
                collect_tv_guide_programs(item, programs_by_channel);
            }
        }
        Value::Object(object) => {
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
            let start_time = read_json_i64(object, "startTime", "start_time");
            let end_time = read_json_i64(object, "endTime", "end_time");
            if !guide_channel.is_empty()
                && !title.is_empty()
                && start_time > 0
                && end_time > start_time
            {
                let key = normalize_channel_name(guide_channel);
                if !key.is_empty() {
                    programs_by_channel.entry(key).or_default().push(TvProgram {
                        title: title.to_string(),
                        start_time,
                        end_time,
                    });
                }
            }
            for nested in object.values() {
                collect_tv_guide_programs(nested, programs_by_channel);
            }
        }
        _ => {}
    }
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
        .ok_or_else(|| "Codice RaiLuce non disponibile per la guida TV.".to_string())?;
    let secret = secret.trim();
    if secret.is_empty() {
        return Err("Codice RaiLuce non valido per la guida TV.".to_string());
    }
    let payload: EncryptedTvGuidePayload = serde_json::from_str(payload_json)
        .map_err(|err| format!("Payload della guida TV non valido: {err}"))?;
    if payload.algorithm != "gzip-xor-base64-v1" {
        return Err(format!(
            "Algoritmo della guida TV non supportato: {}",
            payload.algorithm
        ));
    }

    let encrypted = base64::engine::general_purpose::STANDARD
        .decode(payload.payload_b64)
        .map_err(|err| format!("Payload base64 della guida TV non valido: {err}"))?;
    let mut key = secret.as_bytes().to_vec();
    for part in TV_GUIDE_STATIC_KEY_PARTS {
        key.extend_from_slice(part);
    }
    if key.is_empty() {
        return Err("Chiave della guida TV non valida.".to_string());
    }
    let decrypted = encrypted
        .iter()
        .enumerate()
        .map(|(index, byte)| byte ^ key[index % key.len()])
        .collect::<Vec<_>>();
    let mut decoder = GzDecoder::new(decrypted.as_slice());
    let mut decoded = String::new();
    decoder
        .read_to_string(&mut decoded)
        .map_err(|err| format!("Payload gzip della guida TV non valido: {err}"))?;
    if decoded.trim().is_empty() {
        return Err("URL della guida TV vuoto.".to_string());
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
        Ok(_) => load_channels_from_cache(
            "Il server non ha restituito canali TV. Uso l'ultima lista salvata.",
        ),
        Err(err) => {
            crate::log_debug(&format!("TV channel server request failed: {err}"));
            load_channels_from_cache(
                "Connessione assente o server TV non raggiungibile. Uso l'ultima lista salvata.",
            )
            .map_err(|_| err)
        }
    }
}

fn load_channels_from_server() -> Result<Vec<TvChannel>, String> {
    if TV_API_CLIENT_TOKEN.trim().is_empty() {
        return Err(
            "Token tecnico TV non configurato nella build di Sonarpad. Configura SONARPAD_ROUTE_CLIENT_TOKEN."
                .to_string(),
        );
    }

    let client = Client::builder()
        .user_agent(DEFAULT_USER_AGENT)
        .timeout(Duration::from_secs(10))
        .build()
        .map_err(|err| format!("Impossibile inizializzare il collegamento TV: {err}"))?;
    let response = client
        .get(TV_CHANNELS_URL)
        .header(TV_API_TOKEN_HEADER, TV_API_CLIENT_TOKEN)
        .header(ROUTE_API_TOKEN_HEADER, TV_API_CLIENT_TOKEN)
        .send()
        .map_err(|err| format!("Impossibile scaricare la lista TV: {err}"))?;

    let status = response.status();
    if status.as_u16() == 401 || status.as_u16() == 403 {
        return Err("Client TV non autorizzato dal server.".to_string());
    }
    if !status.is_success() {
        return Err(format!(
            "Impossibile scaricare la lista TV: errore HTTP {}.",
            status.as_u16()
        ));
    }

    let payload: TvServerResponse = response
        .json()
        .map_err(|err| format!("Risposta della lista TV non valida: {err}"))?;
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
        "Altri".to_string()
    } else {
        group_title.to_string()
    };

    Some(TvChannel {
        name,
        url,
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
        fs::create_dir_all(parent)
            .map_err(|err| format!("Impossibile creare la cartella cache TV: {err}"))?;
    }
    let saved_at_unix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let cache = TvChannelsCache {
        saved_at_unix,
        channels: channels.to_vec(),
    };
    let bytes = serde_json::to_vec(&cache)
        .map_err(|err| format!("Impossibile preparare la cache TV: {err}"))?;
    fs::write(path, bytes).map_err(|err| format!("Impossibile salvare la cache TV: {err}"))
}

fn load_channels_from_cache(warning: &str) -> Result<TvChannelLoadResult, String> {
    let path = channels_cache_path();
    let bytes = fs::read(path).map_err(|_| "Nessuna lista TV salvata disponibile.".to_string())?;
    let cache: TvChannelsCache = serde_json::from_slice(&bytes)
        .map_err(|err| format!("La lista TV salvata non è valida: {err}"))?;
    if cache.channels.is_empty() {
        return Err("La lista TV salvata è vuota.".to_string());
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
        .map_err(|err| format!("Impossibile inizializzare il resolver Aurora: {err}"))?;
    let mut token_url = Url::parse(&format!("{endpoint}/token"))
        .map_err(|err| format!("URL token Aurora non valido: {err}"))?;
    token_url.query_pairs_mut().append_pair("realm", realm);
    let referer = channel.url.trim();

    let token_response = add_aurora_headers(client.get(token_url), realm, referer)
        .send()
        .map_err(|err| format!("Richiesta token Aurora non riuscita: {err}"))?;
    if !token_response.status().is_success() {
        return Err(format!(
            "Richiesta token Aurora: errore HTTP {}.",
            token_response.status().as_u16()
        ));
    }
    let token_json: Value = token_response
        .json()
        .map_err(|err| format!("Risposta token Aurora non valida: {err}"))?;
    let token = token_json
        .pointer("/data/attributes/token")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| "Token Aurora non disponibile.".to_string())?;

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
        .map_err(|err| format!("Richiesta riproduzione Aurora non riuscita: {err}"))?;
    if !playback_response.status().is_success() {
        return Err(format!(
            "Richiesta riproduzione Aurora: errore HTTP {}.",
            playback_response.status().as_u16()
        ));
    }
    let playback_json: Value = playback_response
        .json()
        .map_err(|err| format!("Risposta riproduzione Aurora non valida: {err}"))?;
    find_stream_url(&playback_json, ".m3u8")
        .ok_or_else(|| "URL HLS non trovato nella risposta Aurora.".to_string())
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
    let mut url = Url::parse(channel.url.trim())
        .map_err(|err| format!("URL relinker RAI non valido: {err}"))?;
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
        .map_err(|err| format!("Impossibile inizializzare il relinker RAI: {err}"))?;
    let response = client
        .get(url)
        .header("Origin", "https://www.raiplay.it")
        .header("Referer", "https://www.raiplay.it/")
        .send()
        .map_err(|err| format!("Impossibile risolvere lo stream RAI: {err}"))?;
    if !response.status().is_success() {
        return Err(format!(
            "Il relinker RAI ha restituito l'errore HTTP {}.",
            response.status().as_u16()
        ));
    }
    let final_url = response.url().to_string();
    let body = response
        .text()
        .map_err(|err| format!("Risposta del relinker RAI non leggibile: {err}"))?;
    let trimmed = body.trim();
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        return Ok(trimmed.to_string());
    }
    if trimmed.starts_with("#EXTM3U") {
        return Ok(final_url);
    }
    extract_xml_url(trimmed, true)
        .or_else(|| extract_xml_url(trimmed, false))
        .ok_or_else(|| "Stream TV non trovato nella risposta del relinker RAI.".to_string())
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
    channel.name.trim().to_ascii_lowercase().starts_with("rai")
        && channel.url.contains("mediapolis.rai.it/relinker/")
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
    }

    #[test]
    fn extracts_relinker_xml_content_url() {
        let xml = r#"<root><url type="content">https://example.test/live.m3u8</url></root>"#;
        assert_eq!(
            extract_xml_url(xml, true).as_deref(),
            Some("https://example.test/live.m3u8")
        );
    }
}
