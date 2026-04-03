use serde_json::Value;
use std::collections::HashSet;

const RAIPLAYSOUND_BASE_URL: &str = "https://www.raiplaysound.it";
const RAIPLAYSOUND_GENRES_URL: &str = "https://www.raiplaysound.it/generi.json";

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum BrowseItemKind {
    Page,
    Audio,
}

#[derive(Clone, Debug)]
pub(crate) struct BrowseItem {
    pub(crate) id: String,
    pub(crate) label: String,
    pub(crate) title: String,
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
    load_page_from_url(RAIPLAYSOUND_GENRES_URL)
}

pub(crate) fn load_page(path_or_url: &str) -> Result<BrowsePage, String> {
    load_page_from_url(&absolute_url(path_or_url))
}

fn load_page_from_url(url: &str) -> Result<BrowsePage, String> {
    let root = fetch_json(url)?;
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
        collect_cards(cards, None, &mut seen, &mut items);
    }

    if let Some(blocks) = root.get("blocks").and_then(Value::as_array) {
        for block in blocks {
            let section = string_field(block, "title").filter(|value| !value.is_empty());
            if let Some(cards) = block.get("cards").and_then(Value::as_array) {
                collect_cards(cards, section.as_deref(), &mut seen, &mut items);
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
    seen: &mut HashSet<String>,
    items: &mut Vec<BrowseItem>,
) {
    for card in cards {
        if let Some(item) = parse_card(card, section)
            && seen.insert(item.id.clone())
        {
            items.push(item);
        }
    }
}

fn parse_card(card: &Value, section: Option<&str>) -> Option<BrowseItem> {
    let path_id = string_field(card, "path_id").or_else(|| string_field(card, "pathId"));
    let title = preferred_title(card);
    let description = preferred_description(card);
    let audio_url = card
        .get("downloadable_audio")
        .and_then(|audio| string_field(audio, "url"))
        .or_else(|| {
            card.get("audio")
                .and_then(|audio| string_field(audio, "url"))
        });

    let kind = if audio_url.is_some() {
        BrowseItemKind::Audio
    } else if path_id.is_some() {
        BrowseItemKind::Page
    } else {
        return None;
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

    Some(BrowseItem {
        id,
        label: build_label(section, &title, description.as_deref()),
        title,
        path_id: path_id.map(|value| absolute_url(&value)),
        audio_url,
        kind,
    })
}

fn preferred_title(card: &Value) -> String {
    for key in ["toptitle", "episode_title", "title", "label"] {
        if let Some(value) = string_field(card, key).filter(|value| !value.is_empty()) {
            return value;
        }
    }
    "Elemento RaiPlay Sound".to_string()
}

fn preferred_description(card: &Value) -> Option<String> {
    for key in ["subtitle", "description", "vanity"] {
        if let Some(value) = string_field(card, key).filter(|value| !value.is_empty()) {
            return Some(value);
        }
    }
    None
}

fn build_label(_section: Option<&str>, title: &str, description: Option<&str>) -> String {
    let mut parts = Vec::new();
    if !title.trim().is_empty() {
        parts.push(title.trim().to_string());
    }
    if let Some(description) = description.filter(|value| !value.trim().is_empty()) {
        parts.push(description.trim().to_string());
    }
    if parts.is_empty() {
        "Elemento RaiPlay Sound".to_string()
    } else {
        parts.join(" - ")
    }
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

fn absolute_url(path_or_url: &str) -> String {
    let trimmed = path_or_url.trim();
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        trimmed.to_string()
    } else {
        format!("{RAIPLAYSOUND_BASE_URL}{trimmed}")
    }
}
