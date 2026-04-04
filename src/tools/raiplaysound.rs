use base64::Engine;
use serde_json::Value;
use std::collections::HashSet;

const RAIPLAYSOUND_BASE_URL_B64: &str = "BT9NUQVqHVc7T1RfJlUTPlshCyReBjdSSi9E";
const RAIPLAYSOUND_GENRES_URL_B64: &str = "BT9NUQVqHVc7T1RfJlUTPlshCyReBjdSSi9ERy1WGycIPFwmQi0H";
const RAIPLAYSOUND_SEARCH_URL_B64: &str = "BT9NUQVqHVc7T1RfJlUTPlshCyReBjdSSi9ERytHGi8bIRsvHjAIG2AzGT0fKFEMBTVADiVbRl41RBNhQXFdOkIWOEQHLg==";
const RAIPLAYSOUND_SUGGESTION_URL_B64: &str = "BT9NUQVqHVc7T1RfJlUTPlshCyReBjdSSi9ERytHGi8bIRsvHjAIG2AzGT0fKFEMBTVADiVbRl41RBNhQXJdJF4GN1JLNUUPLVYGNhM6HA==";
const RAIPLAYSOUND_SEARCH_SOURCE_PREFIX: &str = "raiplaysound-search:";
const RAIPLAYSOUND_SEARCH_TEMPLATE_IN: &str = "650d4cc74d28b941fec3218c";
const RAIPLAYSOUND_SEARCH_TEMPLATE_OUT: &str = "6516d22540da6c377b151643";
const RAIPLAYSOUND_SEARCH_PAGE_SIZE: u32 = 12;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum BrowseItemKind {
    Page,
    Audio,
}

#[derive(Clone, Debug)]
pub(crate) struct BrowseItem {
    pub(crate) id: String,
    pub(crate) title: String,
    pub(crate) description: Option<String>,
    pub(crate) path_id: Option<String>,
    pub(crate) audio_url: Option<String>,
    pub(crate) kind: BrowseItemKind,
}

#[derive(Clone, Debug)]
pub(crate) struct BrowsePage {
    pub(crate) source: String,
    pub(crate) title: String,
    pub(crate) items: Vec<BrowseItem>,
}

pub(crate) fn load_root_page() -> Result<BrowsePage, String> {
    load_page_from_url(&raiplaysound_genres_url()?)
}

pub(crate) fn load_page(path_or_url: &str) -> Result<BrowsePage, String> {
    if let Some(query) = path_or_url.strip_prefix(RAIPLAYSOUND_SEARCH_SOURCE_PREFIX) {
        return search(query);
    }
    load_page_from_url(&absolute_url(path_or_url)?)
}

pub(crate) fn search(query: &str) -> Result<BrowsePage, String> {
    let trimmed_query = query.trim();
    if trimmed_query.is_empty() {
        return Err("Inserisci un testo da cercare in RaiPlay Sound.".to_string());
    }
    let effective_query =
        refine_search_query(trimmed_query).unwrap_or_else(|| trimmed_query.to_string());

    let body = serde_json::json!({
        "templateIn": RAIPLAYSOUND_SEARCH_TEMPLATE_IN,
        "templateOut": RAIPLAYSOUND_SEARCH_TEMPLATE_OUT,
        "params": {
            "from": 0,
            "size": RAIPLAYSOUND_SEARCH_PAGE_SIZE,
            "param": effective_query,
            "sort": "relevance",
        }
    });
    let body_text = serde_json::to_string(&body)
        .map_err(|err| format!("Richiesta ricerca RaiPlay Sound non valida: {err}"))?;
    let bytes = crate::curl_client::CurlClient::post_form_impersonated(
        &raiplaysound_search_url()?,
        &body_text,
        &["Content-Type: application/json"],
    )
    .map_err(|err| format!("Impossibile cercare in RaiPlay Sound: {err}"))?;
    let root: Value = serde_json::from_slice(&bytes)
        .map_err(|err| format!("Risposta ricerca RaiPlay Sound non valida: {err}"))?;

    let mut items = Vec::new();
    let mut seen = HashSet::new();
    if let Some(cards) = root
        .get("aggs")
        .and_then(|aggs| aggs.get("podcast"))
        .and_then(|podcast| podcast.get("cards"))
        .and_then(Value::as_array)
    {
        collect_cards(cards, None, false, &mut seen, &mut items)?;
    }
    if let Some(cards) = root
        .get("aggs")
        .and_then(|aggs| aggs.get("audio"))
        .and_then(|audio| audio.get("cards"))
        .and_then(Value::as_array)
    {
        collect_cards(cards, None, false, &mut seen, &mut items)?;
    }

    Ok(BrowsePage {
        source: format!("{RAIPLAYSOUND_SEARCH_SOURCE_PREFIX}{trimmed_query}"),
        title: format!("Risultati per {trimmed_query}"),
        items,
    })
}

fn refine_search_query(query: &str) -> Option<String> {
    let normalized_query = normalize_search_text(query);
    if normalized_query.is_empty() {
        return None;
    }

    let body = serde_json::json!({ "text": query });
    let body_text = match serde_json::to_string(&body) {
        Ok(value) => value,
        Err(err) => {
            crate::log_debug(&format!(
                "RaiPlay Sound suggestion request serialization failed: {err}"
            ));
            return None;
        }
    };
    let suggestion_url = match raiplaysound_suggestion_url() {
        Ok(value) => value,
        Err(err) => {
            crate::log_debug(&format!(
                "RaiPlay Sound suggestion URL decode failed: {err}"
            ));
            return None;
        }
    };
    let bytes = match crate::curl_client::CurlClient::post_form_impersonated(
        &suggestion_url,
        &body_text,
        &["Content-Type: application/json"],
    ) {
        Ok(value) => value,
        Err(err) => {
            crate::log_debug(&format!("RaiPlay Sound suggestion request failed: {err}"));
            return None;
        }
    };
    let root: Value = match serde_json::from_slice(&bytes) {
        Ok(value) => value,
        Err(err) => {
            crate::log_debug(&format!(
                "RaiPlay Sound suggestion JSON decode failed: {err}"
            ));
            return None;
        }
    };
    let suggestion = root
        .get("suggestions")
        .and_then(Value::as_array)
        .and_then(|items| items.first())
        .and_then(|item| string_field(item, "text"))?;
    let normalized_suggestion = normalize_search_text(&suggestion);

    if normalized_suggestion == normalized_query
        || normalized_suggestion.starts_with(&normalized_query)
    {
        return Some(suggestion);
    }
    None
}

fn load_page_from_url(url: &str) -> Result<BrowsePage, String> {
    let root = fetch_json(url)?;
    let is_root_page = url.eq_ignore_ascii_case(&raiplaysound_genres_url()?);
    let title = string_field(&root, "title")
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "RaiPlay Sound".to_string());
    let mut items = Vec::new();
    let mut seen = HashSet::new();

    if let Some(cards) = root
        .get("block")
        .and_then(|block| block.get("cards"))
        .and_then(Value::as_array)
    {
        collect_cards(cards, None, is_root_page, &mut seen, &mut items)?;
    }

    if let Some(blocks) = root.get("blocks").and_then(Value::as_array) {
        for block in blocks {
            let section = string_field(block, "title").filter(|value| !value.is_empty());
            if let Some(cards) = block.get("cards").and_then(Value::as_array) {
                collect_cards(
                    cards,
                    section.as_deref(),
                    is_root_page,
                    &mut seen,
                    &mut items,
                )?;
            }
        }
    }

    Ok(BrowsePage {
        source: url.to_string(),
        title,
        items,
    })
}

fn collect_cards(
    cards: &[Value],
    section: Option<&str>,
    is_root_page: bool,
    seen: &mut HashSet<String>,
    items: &mut Vec<BrowseItem>,
) -> Result<(), String> {
    for card in cards {
        if let Some(item) = parse_card(card, section, is_root_page)?
            && seen.insert(item.id.clone())
        {
            items.push(item);
        }
    }
    Ok(())
}

fn parse_card(
    card: &Value,
    _section: Option<&str>,
    is_root_page: bool,
) -> Result<Option<BrowseItem>, String> {
    let path_id = string_field(card, "path_id").or_else(|| string_field(card, "pathId"));
    let title = preferred_title(card);
    if is_root_page && should_hide_root_item(&title) {
        return Ok(None);
    }
    let description = preferred_description(card);
    let audio_url = card
        .get("downloadable_audio")
        .and_then(|audio| string_field(audio, "url"))
        .or_else(|| {
            card.get("downlodable_audio")
                .and_then(|audio| string_field(audio, "url"))
        })
        .or_else(|| {
            card.get("audio")
                .and_then(|audio| string_field(audio, "url"))
        });

    let kind = if audio_url.is_some() {
        BrowseItemKind::Audio
    } else if path_id.is_some() {
        BrowseItemKind::Page
    } else {
        return Ok(None);
    };

    let id = match kind {
        BrowseItemKind::Audio => {
            format!(
                "audio|{}|{}",
                audio_url.clone().unwrap_or_default(),
                path_id.clone().unwrap_or_default()
            )
        }
        BrowseItemKind::Page => format!("page|{}", path_id.clone().unwrap_or_default()),
    };

    Ok(Some(BrowseItem {
        id,
        title,
        description,
        path_id: path_id.map(|value| absolute_url(&value)).transpose()?,
        audio_url,
        kind,
    }))
}

fn preferred_title(card: &Value) -> String {
    for key in [
        "titolo",
        "toptitle",
        "episode_title",
        "title",
        "label",
        "programma",
    ] {
        if let Some(value) = string_field(card, key).filter(|value| !value.is_empty()) {
            return value;
        }
    }
    "Elemento RaiPlay Sound".to_string()
}

fn preferred_description(card: &Value) -> Option<String> {
    for key in [
        "sommario",
        "subtitle",
        "description",
        "vanity",
        "friendlyType",
    ] {
        if let Some(value) = string_field(card, key).filter(|value| !value.is_empty()) {
            return Some(value);
        }
    }
    None
}

fn fetch_json(url: &str) -> Result<Value, String> {
    let bytes = crate::curl_client::CurlClient::fetch_url_impersonated(url)
        .map_err(|err| format!("Impossibile caricare i dati di RaiPlay Sound: {err}"))?;
    serde_json::from_slice(&bytes)
        .map_err(|err| format!("Risposta JSON RaiPlay Sound non valida: {err}"))
}

fn string_field(value: &Value, key: &str) -> Option<String> {
    value
        .get(key)
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(ToOwned::to_owned)
}

fn absolute_url(path_or_url: &str) -> Result<String, String> {
    let trimmed = path_or_url.trim();
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        Ok(trimmed.to_string())
    } else {
        Ok(format!("{}{}", raiplaysound_base_url()?, trimmed))
    }
}

fn normalize_search_text(text: &str) -> String {
    text.split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

fn should_hide_root_item(title: &str) -> bool {
    matches!(
        title.trim(),
        "Audiodescrizioni-fiction" | "Audiodescrizioni_film"
    )
}

fn raiplaysound_base_url() -> Result<String, String> {
    decode_raiplaysound_url(RAIPLAYSOUND_BASE_URL_B64)
}

fn raiplaysound_genres_url() -> Result<String, String> {
    decode_raiplaysound_url(RAIPLAYSOUND_GENRES_URL_B64)
}

fn raiplaysound_search_url() -> Result<String, String> {
    decode_raiplaysound_url(RAIPLAYSOUND_SEARCH_URL_B64)
}

fn raiplaysound_suggestion_url() -> Result<String, String> {
    decode_raiplaysound_url(RAIPLAYSOUND_SUGGESTION_URL_B64)
}

fn decode_raiplaysound_url(encoded: &str) -> Result<String, String> {
    let key = raiplaysound_obfuscated_url_key()?;
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(encoded)
        .map_err(|err| format!("URL RaiPlay Sound offuscato non valido: {err}"))?;
    let decoded: Vec<u8> = bytes
        .into_iter()
        .enumerate()
        .map(|(index, byte)| byte ^ key[index % key.len()])
        .collect();
    String::from_utf8(decoded)
        .map_err(|err| format!("URL RaiPlay Sound decodificato non valido: {err}"))
}

fn raiplaysound_obfuscated_url_key() -> Result<Vec<u8>, String> {
    Ok(resolve_raiplaysound_secret_key()?.into_bytes())
}

fn resolve_raiplaysound_secret_key() -> Result<String, String> {
    if let Some(secret_key) = crate::settings::load_saved_rai_luce_code() {
        let trimmed = secret_key.trim();
        if !trimmed.is_empty() {
            return Ok(trimmed.to_string());
        }
    }
    Err("Chiave Luce mancante: inserisci il codice nelle impostazioni RSS/Podcast.".to_string())
}
