use base64::Engine;
use chrono::{TimeZone, Utc};
use serde::{Deserialize, Serialize};

const RAI_AUDIODESCRIZIONI_LIST_URL_KEY_A: &[u8] = b"rai-";
const RAI_AUDIODESCRIZIONI_LIST_URL_KEY_B: &[u8] = b"audio";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64_A: &str = "GhUdXRJPS0YLQ1JcVw0dVAAM";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64_B: &str = "GFYERk8WCAYaFgcbQg8BSgcK";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64_C: &str = "Bk42SQMqSwYIFQg2RA8qEB9A";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64_D: &str = "AgAKRgASAQ1AExQNRA5aCAAc";
const RAI_AUDIODESCRIZIONI_LIST_URL_B64_E: &str = "Bk8DXg4b";

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

pub(crate) fn load_catalog() -> Result<Catalog, String> {
    fetch_catalog()
}

pub(crate) fn resolve_audio_url(audio_url: &str) -> Result<String, String> {
    let audio_url = audio_url.trim();
    if audio_url.is_empty() {
        return Err("L'audiodescrizione selezionata non ha un URL audio disponibile.".to_string());
    }
    Ok(audio_url.to_string())
}

fn fetch_catalog() -> Result<Catalog, String> {
    let source_url = decode_obfuscated_url(&obfuscated_list_url_b64(), &obfuscated_list_url_key())?;
    let raw = fetch_text_blocking(&source_url)?;
    let entries: Vec<ExternalItem> =
        serde_json::from_str(&raw).map_err(|err| format!("Catalogo Rai non valido: {err}"))?;
    let now = Utc::now().to_rfc3339();
    let mut items = Vec::new();

    for (item_index, item) in entries.into_iter().enumerate() {
        let title = item.title.trim().to_string();
        let audio_url = item.url.trim().to_string();
        if title.is_empty() || audio_url.is_empty() {
            continue;
        }

        let set_name = item.part_of.trim().to_string();
        let set_id = slugify(&set_name);
        let source_order = item_index as i64;
        let (date, iso_date, gen_date) = format_added_date(item.added);
        items.push(CatalogItem {
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
        });
    }

    Ok(Catalog {
        source: source_url,
        generated_at: now,
        total_items: items.len(),
        items,
    })
}

fn fetch_text_blocking(url: &str) -> Result<String, String> {
    let bytes = crate::curl_client::CurlClient::fetch_url_impersonated(url)
        .map_err(|err| format!("Impossibile scaricare il catalogo Rai: {err}"))?;
    String::from_utf8(bytes)
        .map_err(|err| format!("Catalogo Rai non decodificabile come UTF-8: {err}"))
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

fn obfuscated_list_url_b64() -> String {
    [
        RAI_AUDIODESCRIZIONI_LIST_URL_B64_A,
        RAI_AUDIODESCRIZIONI_LIST_URL_B64_B,
        RAI_AUDIODESCRIZIONI_LIST_URL_B64_C,
        RAI_AUDIODESCRIZIONI_LIST_URL_B64_D,
        RAI_AUDIODESCRIZIONI_LIST_URL_B64_E,
    ]
    .concat()
}

fn obfuscated_list_url_key() -> Vec<u8> {
    [
        RAI_AUDIODESCRIZIONI_LIST_URL_KEY_A,
        RAI_AUDIODESCRIZIONI_LIST_URL_KEY_B,
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
