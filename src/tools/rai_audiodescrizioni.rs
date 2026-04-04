use base64::Engine;
use chrono::{TimeZone, Utc};
use flate2::read::GzDecoder;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::BTreeMap;
use std::io::{Read, Write};
use std::path::PathBuf;

const RAI_AUDIODESCRIZIONI_URL_KEY_A: &[u8] = b"rai-";
const RAI_AUDIODESCRIZIONI_URL_KEY_B: &[u8] = b"audio";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64: &str = "GhUdXRJPS0YdExZHSggBDBwNBxIMXwIaCh0KHBVHTg4YSygCEBMGFVdaNwYBExMZTAVYMAYAHhJGXwQTF0YHFwANXk4YBQABXQYMQwQHBR0KFk4FWAIQSQUGARVHSA8WSgMcHQ8=";
const RAI_AUDIODESCRIZIONI_CATALOGUE_URL_B64: &str = "GhUdXRJPS0YdExZHSggBDBwNBxIMXwIaCh0KHBVHTg4YSygCEBMGFVdaNwYBExMZTAVYMAYAHhJGXwQTF0YHFwANXk4YBQABXQYMQwQHBR0KFk4FWAIQSQoOBgAFQgYAAUcKHAJHRxIaCg==";
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

pub(crate) fn grouped_catalog_dump_path() -> PathBuf {
    crate::settings::settings_dir().join("rai_grouped_catalog_dump.txt")
}

pub(crate) fn write_grouped_catalog_dump(groups: &[CatalogGroup]) -> Result<PathBuf, String> {
    let path = grouped_catalog_dump_path();
    let mut file = std::fs::File::create(&path)
        .map_err(|err| format!("Impossibile creare dump catalogo Rai: {err}"))?;
    writeln!(
        file,
        "Dump catalogo completo audiodescrizioni Rai - {}",
        chrono::Local::now().format("%Y-%m-%d %H:%M:%S")
    )
    .map_err(|err| format!("Impossibile scrivere dump catalogo Rai: {err}"))?;
    for group in groups {
        writeln!(file).map_err(|err| format!("Impossibile scrivere dump catalogo Rai: {err}"))?;
        writeln!(file, "[{}]", group.title)
            .map_err(|err| format!("Impossibile scrivere dump catalogo Rai: {err}"))?;
        for (index, item) in group.items.iter().enumerate() {
            writeln!(file, "{}. {}", index + 1, item.title.trim())
                .map_err(|err| format!("Impossibile scrivere dump catalogo Rai: {err}"))?;
        }
    }
    Ok(path)
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

pub(crate) fn resolve_audio_url_for_clipboard(audio_url: &str) -> Result<String, String> {
    let audio_url = resolve_audio_url(audio_url)?;
    if !audio_url.contains("/relinker/relinkerServlet") {
        return Ok(audio_url);
    }

    crate::curl_client::CurlClient::resolve_final_url_iphone_impersonated(&audio_url)
        .map_err(|err| format!("Impossibile risolvere l'URL audio Rai: {err}"))
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
    if let Some(secret_key) = crate::settings::load_saved_rai_luce_code() {
        let trimmed = secret_key.trim();
        if !trimmed.is_empty() {
            return Ok(trimmed.to_string());
        }
    }

    Err("Chiave Luce mancante: inserisci il codice nelle impostazioni RSS/Podcast.".to_string())
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
                compare_natural_labels(
                    &normalize_sort_key(&left.title),
                    &normalize_sort_key(&right.title),
                )
                .then_with(|| compare_natural_labels(&left.title, &right.title))
                .then_with(|| left.source_order.cmp(&right.source_order))
            });
            items.dedup_by(|left, right| dedupe_key(&left.title) == dedupe_key(&right.title));
            CatalogGroup { title, items }
        })
        .collect::<Vec<_>>();

    normalized_groups.sort_by(|left, right| {
        compare_natural_labels(&left.title, &right.title)
            .then_with(|| sortable_label(&left.title).cmp(&sortable_label(&right.title)))
    });
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

fn normalize_sort_key(input: &str) -> String {
    input.split_whitespace().collect::<Vec<_>>().join(" ")
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
    sortable_label(&normalize_sort_key(&normalize_item_title(input)))
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

#[cfg(test)]
mod tests {
    use super::{CatalogGroup, CatalogItem, normalize_grouped_catalog};

    fn item(title: &str, source_order: i64) -> CatalogItem {
        CatalogItem {
            set_id: String::new(),
            set_name: String::new(),
            item_id: format!("item-{source_order}"),
            title: title.to_string(),
            date: String::new(),
            iso_date: None,
            description: String::new(),
            url: String::new(),
            image: String::new(),
            image_timestamp: None,
            audio_url: String::new(),
            gen_date: None,
            source_order,
            first_seen_at: None,
            last_seen_at: None,
        }
    }

    #[test]
    fn grouped_catalog_sorts_series_titles_naturally() {
        let groups = vec![
            CatalogGroup {
                title: "Un Medico in Famiglia 10".to_string(),
                items: vec![item("Un medico in famiglia 10 - Puntata 1", 0)],
            },
            CatalogGroup {
                title: "Un medico in famiglia 6".to_string(),
                items: vec![item("Un medico in famiglia 6 - Puntata 1", 1)],
            },
            CatalogGroup {
                title: "Un medico in famiglia 5".to_string(),
                items: vec![item("Un medico in famiglia 5 - Puntata 1", 2)],
            },
        ];

        let normalized = normalize_grouped_catalog(groups);
        let titles = normalized
            .iter()
            .map(|group| group.title.as_str())
            .collect::<Vec<_>>();

        assert_eq!(
            titles,
            vec![
                "Un medico in famiglia 5",
                "Un medico in famiglia 6",
                "Un Medico in Famiglia 10",
            ]
        );
    }

    #[test]
    fn grouped_catalog_sorts_items_with_double_spaces_naturally() {
        let groups = vec![CatalogGroup {
            title: "Questo nostro amore".to_string(),
            items: vec![
                item("Questo nostro amore - Puntata  4", 0),
                item("Questo nostro amore - Puntata 1", 1),
                item("Questo nostro amore - Puntata 2", 2),
            ],
        }];

        let normalized = normalize_grouped_catalog(groups);
        let titles = normalized[0]
            .items
            .iter()
            .map(|item| item.title.as_str())
            .collect::<Vec<_>>();

        assert_eq!(
            titles,
            vec![
                "Questo nostro amore - Puntata 1",
                "Questo nostro amore - Puntata 2",
                "Questo nostro amore - Puntata  4",
            ]
        );
    }

    #[test]
    #[ignore]
    fn debug_dump_current_grouped_catalog_to_file() {
        let groups = super::load_grouped_catalog().expect("expected Rai grouped catalog");
        let path =
            super::write_grouped_catalog_dump(&groups).expect("expected Rai grouped catalog dump");
        println!("{}", path.display());
    }
}
