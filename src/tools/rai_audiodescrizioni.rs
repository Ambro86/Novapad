use std::cmp::Ordering;
use std::collections::BTreeMap;
use std::io::Read;
use std::path::PathBuf;

use base64::Engine;
use chrono::{TimeZone, Utc};
use flate2::read::GzDecoder;
use serde::{Deserialize, Serialize};

const RAI_AUDIODESCRIZIONI_URL_KEY_A: &[u8] = b"rai-";
const RAI_AUDIODESCRIZIONI_URL_KEY_B: &[u8] = b"audio";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64: &str = "GhUdXRJPS0YdExZHSggBDBwNBxIMXwIaCh0KHBVHTg4YSygCEBMGFVdaNwYBExMZTAVYMAYAHhJGXwQTF0YHFwANXk4YBQABXQYMQwQHBR0KFk4FWAIQSQUGARVHSA8WSgMcHQ8=";
const RAI_AUDIODESCRIZIONI_CATALOGUE_URL_B64: &str = "GhUdXRJPS0YdExZHSggBDBwNBxIMXwIaCh0KHBVHTg4YSygCEBMGFVdaNwYBExMZTAVYMAYAHhJGXwQTF0YHFwANXk4YBQABXQYMQwQHBR0KFk4FWAIQSQoOBgAFQgYAAUcKHAJHRxIaCg==";
const LUCE_PAYLOAD_KEY_ENV: &str = "LUCE_ENCRYPTION_KEY";
const LUCE_KEY_FILE_NAME: &str = "luce.key";
const LUCE_PAYLOAD_STATIC_KEY_PARTS: &[&[u8]] = &[b"sonar", b"pad-", b"SonarSecure-"];

#[derive(Clone, Debug, Serialize, Deserialize)]
pub(crate) struct CatalogItem {
    #[serde(rename = "setId")]
    pub(crate) set_id: String,
    #[serde(rename = "setName")]
    pub(crate) set_name: String,
    #[serde(rename = "itemId")]
    pub(crate) item_id: String,
    #[serde(default)]
    pub(crate) title: String,
    #[serde(default)]
    pub(crate) date: String,
    #[serde(default)]
    #[serde(rename = "isoDate")]
    pub(crate) iso_date: Option<String>,
    #[serde(default)]
    pub(crate) description: String,
    #[serde(default)]
    pub(crate) url: String,
    #[serde(default)]
    pub(crate) image: String,
    #[serde(default)]
    #[serde(rename = "imageTimestamp")]
    pub(crate) image_timestamp: Option<i64>,
    #[serde(default)]
    #[serde(rename = "audioUrl")]
    pub(crate) audio_url: String,
    #[serde(default)]
    #[serde(rename = "genDate")]
    pub(crate) gen_date: Option<String>,
    #[serde(default)]
    #[serde(rename = "sourceOrder")]
    pub(crate) source_order: i64,
    #[serde(default)]
    #[serde(rename = "firstSeenAt")]
    pub(crate) first_seen_at: Option<String>,
    #[serde(default)]
    #[serde(rename = "lastSeenAt")]
    pub(crate) last_seen_at: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub(crate) struct Catalog {
    pub(crate) source: String,
    #[serde(rename = "generatedAt")]
    pub(crate) generated_at: String,
    #[serde(rename = "totalItems")]
    pub(crate) total_items: usize,
    pub(crate) items: Vec<CatalogItem>,
}

#[derive(Clone, Debug)]
pub(crate) struct CatalogGroup {
    pub(crate) title: String,
    pub(crate) items: Vec<CatalogItem>,
}

#[derive(Debug, Deserialize)]
struct ExternalItem {
    title: String,
    #[serde(default)]
    #[serde(rename = "partOf")]
    part_of: String,
    #[serde(default)]
    added: Option<i64>,
    #[serde(default)]
    url: String,
}

#[derive(Debug, Deserialize)]
struct ExternalCatalogueGroup {
    title: String,
    #[serde(default)]
    data: Vec<ExternalCatalogueItem>,
}

#[derive(Debug, Deserialize)]
struct ExternalCatalogueItem {
    title: String,
    #[serde(default)]
    url: String,
}

#[derive(Debug, Deserialize)]
struct EncryptedPayload {
    algorithm: String,
    #[serde(rename = "payload_b64")]
    payload_b64: String,
}

pub(crate) fn load_catalog() -> Result<Catalog, String> {
    fetch_catalog()
}

pub(crate) fn load_grouped_catalog() -> Result<Vec<CatalogGroup>, String> {
    fetch_grouped_catalog()
}

pub(crate) fn is_luce_key_missing_error(err: &str) -> bool {
    err.starts_with("Chiave Luce mancante:")
}

pub(crate) fn resolve_audio_url(audio_url: &str) -> Result<String, String> {
    let audio_url = audio_url.trim();
    if audio_url.is_empty() {
        return Err("L'audiodescrizione selezionata non ha un URL audio disponibile.".to_string());
    }
    Ok(audio_url.to_string())
}

fn fetch_catalog() -> Result<Catalog, String> {
    let source_url =
        decode_obfuscated_url(RAI_AUDIODESCRIZIONI_LIST_URL_B64, &obfuscated_url_key())?;
    let raw = fetch_and_decode_luce_payload(&source_url)?;
    let entries: Vec<ExternalItem> =
        serde_json::from_str(&raw).map_err(|err| format!("Catalogo Rai non valido: {err}"))?;
    let now = Utc::now().to_rfc3339();
    let mut items_with_added = Vec::new();

    for (item_index, item) in entries.into_iter().enumerate() {
        let title = normalize_item_title(&item.title);
        let audio_url = item.url.trim().to_string();
        if title.is_empty() || audio_url.is_empty() {
            continue;
        }

        let set_name = item.part_of.trim().to_string();
        let set_id = slugify(&set_name);
        let source_order = item_index as i64;
        let (date, iso_date, gen_date) = format_added_date(item.added);
        items_with_added.push((
            item.added.unwrap_or(i64::MIN),
            CatalogItem {
                set_id: set_id.clone(),
                set_name,
                item_id: format!("{set_id}|{}|{audio_url}", slugify(&title)),
                title,
                date,
                iso_date,
                description: String::new(),
                url: String::new(),
                image: String::new(),
                image_timestamp: None,
                audio_url,
                gen_date,
                source_order,
                first_seen_at: Some(now.clone()),
                last_seen_at: Some(now.clone()),
            },
        ));
    }

    items_with_added.sort_by(|(left_added, left_item), (right_added, right_item)| {
        right_added
            .cmp(left_added)
            .then_with(|| left_item.source_order.cmp(&right_item.source_order))
    });
    let items = items_with_added
        .into_iter()
        .map(|(_, item)| item)
        .collect::<Vec<_>>();

    Ok(Catalog {
        source: source_url,
        generated_at: now,
        total_items: items.len(),
        items,
    })
}

fn fetch_grouped_catalog() -> Result<Vec<CatalogGroup>, String> {
    let source_url = decode_obfuscated_url(
        RAI_AUDIODESCRIZIONI_CATALOGUE_URL_B64,
        &obfuscated_url_key(),
    )?;
    let raw = fetch_and_decode_luce_payload(&source_url)?;
    let groups: Vec<ExternalCatalogueGroup> = serde_json::from_str(&raw)
        .map_err(|err| format!("Catalogo Rai completo non valido: {err}"))?;
    let mut parsed_groups = Vec::new();

    for group in groups {
        let title = group.title.trim().to_string();
        if title.is_empty() {
            continue;
        }
        let set_id = slugify(&title);
        let mut items = Vec::new();
        for (item_index, item) in group.data.into_iter().enumerate() {
            let item_title = normalize_item_title(&item.title);
            let audio_url = item.url.trim().to_string();
            if item_title.is_empty() || audio_url.is_empty() {
                continue;
            }
            items.push(CatalogItem {
                set_id: set_id.clone(),
                set_name: title.clone(),
                item_id: format!("{set_id}|{}|{audio_url}", slugify(&item_title)),
                title: item_title,
                date: String::new(),
                iso_date: None,
                description: String::new(),
                url: String::new(),
                image: String::new(),
                image_timestamp: None,
                audio_url,
                gen_date: None,
                source_order: item_index as i64,
                first_seen_at: None,
                last_seen_at: None,
            });
        }

        if !items.is_empty() {
            parsed_groups.push(CatalogGroup { title, items });
        }
    }

    Ok(normalize_grouped_catalog(parsed_groups))
}

fn fetch_text_blocking(url: &str) -> Result<String, String> {
    let bytes = crate::curl_client::CurlClient::fetch_url_impersonated(url)
        .map_err(|err| format!("Impossibile scaricare il catalogo Rai: {err}"))?;
    String::from_utf8(bytes)
        .map_err(|err| format!("Catalogo Rai non decodificabile come UTF-8: {err}"))
}

fn fetch_and_decode_luce_payload(url: &str) -> Result<String, String> {
    let raw = fetch_text_blocking(url)?;
    let payload: EncryptedPayload =
        serde_json::from_str(&raw).map_err(|err| format!("Payload Luce non valido: {err}"))?;
    if payload.algorithm != "gzip-xor-base64-v1" {
        return Err(format!(
            "Algoritmo payload Luce non supportato: {}",
            payload.algorithm
        ));
    }

    let secret_key = resolve_luce_secret_key()?;
    let encrypted = base64::engine::general_purpose::STANDARD
        .decode(payload.payload_b64)
        .map_err(|err| format!("Payload Luce base64 non valido: {err}"))?;
    let decrypted = xor_with_luce_key(&encrypted, &secret_key, LUCE_PAYLOAD_STATIC_KEY_PARTS)?;
    let mut decoder = GzDecoder::new(decrypted.as_slice());
    let mut decoded = String::new();
    decoder
        .read_to_string(&mut decoded)
        .map_err(|err| format!("Payload Luce gzip non valido: {err}"))?;
    Ok(decoded)
}

fn xor_with_luce_key(
    payload: &[u8],
    secret_key: &str,
    static_key_parts: &[&[u8]],
) -> Result<Vec<u8>, String> {
    let mut key = Vec::new();
    for part in static_key_parts {
        key.extend_from_slice(part);
    }
    key.extend_from_slice(secret_key.as_bytes());
    if key.is_empty() {
        return Err("Chiave payload Luce non valida.".to_string());
    }
    Ok(payload
        .iter()
        .enumerate()
        .map(|(index, byte)| byte ^ key[index % key.len()])
        .collect())
}

fn resolve_luce_secret_key() -> Result<String, String> {
    if let Ok(secret_key) = std::env::var(LUCE_PAYLOAD_KEY_ENV) {
        let trimmed = secret_key.trim();
        if !trimmed.is_empty() {
            return Ok(trimmed.to_string());
        }
    }

    for path in luce_key_candidate_paths() {
        let Ok(contents) = std::fs::read_to_string(&path) else {
            continue;
        };
        let trimmed = contents.trim();
        if !trimmed.is_empty() {
            return Ok(trimmed.to_string());
        }
    }

    if let Some(secret_key) = crate::settings::load_saved_rai_luce_code() {
        let trimmed = secret_key.trim();
        if !trimmed.is_empty() {
            return Ok(trimmed.to_string());
        }
    }

    Err(format!(
        "Chiave Luce mancante: imposta {LUCE_PAYLOAD_KEY_ENV}, crea {LUCE_KEY_FILE_NAME} accanto all'eseguibile o nella cartella Sonarpad, oppure inserisci il codice nelle impostazioni RSS/Podcast."
    ))
}

fn luce_key_candidate_paths() -> Vec<PathBuf> {
    let mut paths = Vec::new();

    if let Ok(exe_path) = std::env::current_exe()
        && let Some(parent) = exe_path.parent()
    {
        paths.push(parent.join(LUCE_KEY_FILE_NAME));
    }

    paths.push(crate::settings::settings_dir().join(LUCE_KEY_FILE_NAME));
    paths
}

fn slugify(input: &str) -> String {
    let mut slug = String::with_capacity(input.len());
    let mut last_was_sep = false;

    for ch in input.chars() {
        if ch.is_ascii_alphanumeric() {
            slug.push(ch.to_ascii_lowercase());
            last_was_sep = false;
        } else if !last_was_sep {
            slug.push('-');
            last_was_sep = true;
        }
    }

    slug.trim_matches('-').to_string()
}

fn decode_obfuscated_url(encoded: &str, key: &[u8]) -> Result<String, String> {
    if key.is_empty() {
        return Err("Chiave URL Rai non valida.".to_string());
    }

    let bytes = base64::engine::general_purpose::STANDARD
        .decode(encoded)
        .map_err(|err| format!("URL Rai offuscato non valido: {err}"))?;
    let decoded: Vec<u8> = bytes
        .into_iter()
        .enumerate()
        .map(|(index, byte)| byte ^ key[index % key.len()])
        .collect();
    String::from_utf8(decoded).map_err(|err| format!("URL Rai decodificato non valido: {err}"))
}

fn obfuscated_url_key() -> Vec<u8> {
    [
        RAI_AUDIODESCRIZIONI_URL_KEY_A,
        RAI_AUDIODESCRIZIONI_URL_KEY_B,
    ]
    .concat()
}

fn format_added_date(added: Option<i64>) -> (String, Option<String>, Option<String>) {
    let Some(timestamp) = added else {
        return (String::new(), None, None);
    };
    let Some(datetime) = Utc.timestamp_opt(timestamp, 0).single() else {
        return (String::new(), None, None);
    };
    (
        datetime.format("%d/%m/%Y").to_string(),
        Some(datetime.format("%Y-%m-%d").to_string()),
        Some(datetime.format("%d/%m/%Y %H:%M:%S").to_string()),
    )
}

fn normalize_grouped_catalog(groups: Vec<CatalogGroup>) -> Vec<CatalogGroup> {
    let mut merged = BTreeMap::<String, Vec<CatalogItem>>::new();

    for group in groups {
        let normalized_title = normalize_group_title(&group.title);
        merged
            .entry(normalized_title)
            .or_default()
            .extend(group.items.into_iter());
    }

    let mut normalized_groups = merged
        .into_iter()
        .map(|(title, mut items)| {
            items.sort_by(|left, right| {
                compare_natural_labels(&left.title, &right.title)
                    .then_with(|| left.source_order.cmp(&right.source_order))
            });
            items.dedup_by(|left, right| dedupe_key(&left.title) == dedupe_key(&right.title));
            CatalogGroup { title, items }
        })
        .collect::<Vec<_>>();

    normalized_groups
        .sort_by(|left, right| sortable_label(&left.title).cmp(&sortable_label(&right.title)));
    normalized_groups
}

fn normalize_group_title(title: &str) -> String {
    let trimmed = title.trim();
    let lower = trimmed.to_lowercase();
    if (trimmed.starts_with("Film (") && trimmed.ends_with(')'))
        || lower == "film - audiodescrizioni"
    {
        "Film".to_string()
    } else if lower == "miniserie tv - audiodescrizioni" {
        "Miniserie Tv".to_string()
    } else {
        trimmed.to_string()
    }
}

fn sortable_label(input: &str) -> String {
    input.trim().to_lowercase()
}

fn normalize_item_title(input: &str) -> String {
    let trimmed = input.trim();
    let lower = trimmed.to_lowercase();
    if let Some(prefix) = lower.strip_suffix(" - audiodescrizione") {
        let prefix_len = prefix.len();
        trimmed[..prefix_len].trim().to_string()
    } else {
        trimmed.to_string()
    }
}

fn dedupe_key(input: &str) -> String {
    sortable_label(&normalize_item_title(input))
}

fn compare_natural_labels(left: &str, right: &str) -> Ordering {
    let left_chars = left.trim().to_lowercase().chars().collect::<Vec<_>>();
    let right_chars = right.trim().to_lowercase().chars().collect::<Vec<_>>();
    let mut left_index = 0usize;
    let mut right_index = 0usize;

    while left_index < left_chars.len() && right_index < right_chars.len() {
        let left_char = left_chars[left_index];
        let right_char = right_chars[right_index];
        if left_char.is_ascii_digit() && right_char.is_ascii_digit() {
            let left_start = left_index;
            while left_index < left_chars.len() && left_chars[left_index].is_ascii_digit() {
                left_index += 1;
            }
            let right_start = right_index;
            while right_index < right_chars.len() && right_chars[right_index].is_ascii_digit() {
                right_index += 1;
            }

            let left_number = left_chars[left_start..left_index]
                .iter()
                .collect::<String>()
                .parse::<u64>()
                .unwrap_or(0);
            let right_number = right_chars[right_start..right_index]
                .iter()
                .collect::<String>()
                .parse::<u64>()
                .unwrap_or(0);
            match left_number.cmp(&right_number) {
                Ordering::Equal => continue,
                non_equal => return non_equal,
            }
        }

        match left_char.cmp(&right_char) {
            Ordering::Equal => {
                left_index += 1;
                right_index += 1;
            }
            non_equal => return non_equal,
        }
    }

    left_chars.len().cmp(&right_chars.len())
}
