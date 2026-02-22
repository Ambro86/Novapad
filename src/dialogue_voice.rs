use crate::settings::TtsEngine;
use sha2::Digest;
use std::path::{Path, PathBuf};

#[derive(Clone, PartialEq, Eq)]
pub struct DialogueVoiceConfig {
    pub engine: TtsEngine,
    pub voice: String,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub opening_quote: String,
    pub closing_quote: String,
    pub allow_multiline: bool,
}

impl DialogueVoiceConfig {
    fn legacy_sidecar_path_for(path: &Path) -> PathBuf {
        let file_name = path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("document");
        let mut sidecar = path.to_path_buf();
        sidecar.set_file_name(format!("{file_name}.dialogue.ini"));
        sidecar
    }

    pub fn sidecar_path_for(path: &Path) -> PathBuf {
        let mut hasher = sha2::Sha256::new();
        hasher.update(path.to_string_lossy().as_bytes());
        let hash = hex::encode(hasher.finalize());

        let file_stem = path
            .file_stem()
            .and_then(|s| s.to_str())
            .filter(|s| !s.trim().is_empty())
            .unwrap_or("document");
        let safe_stem: String = file_stem
            .chars()
            .map(|c| {
                if c.is_ascii_alphanumeric() || c == '-' || c == '_' {
                    c
                } else {
                    '_'
                }
            })
            .collect();
        let file_name = format!("{safe_stem}.{}.dialogue.ini", &hash[..16]);
        let mut sidecar = crate::settings::settings_dir().join("dialogs");
        sidecar.push(file_name);
        sidecar
    }
}

pub fn parse_engine_input(input: &str) -> Option<TtsEngine> {
    match input.trim().to_ascii_lowercase().as_str() {
        "edge" | "microsoft" | "microsoft voices" => Some(TtsEngine::Edge),
        "sapi5" | "sapi 5" => Some(TtsEngine::Sapi5),
        "sapi4" | "sapi 4" => Some(TtsEngine::Sapi4),
        _ => None,
    }
}

pub fn engine_to_key(engine: TtsEngine) -> &'static str {
    match engine {
        TtsEngine::Edge => "edge",
        TtsEngine::Sapi5 => "sapi5",
        TtsEngine::Sapi4 => "sapi4",
    }
}

pub fn save_dialogue_voice_config(path: &Path, cfg: &DialogueVoiceConfig) -> Result<(), String> {
    let sidecar = DialogueVoiceConfig::sidecar_path_for(path);
    if let Some(parent) = sidecar.parent() {
        std::fs::create_dir_all(parent).map_err(|e| {
            format!(
                "Failed to create dialogue config folder {:?}: {}",
                parent, e
            )
        })?;
    }
    let body = format!(
        "engine={}\nvoice={}\nrate={}\npitch={}\nvolume={}\nopen_quote={}\nclose_quote={}\nallow_multiline={}\n",
        engine_to_key(cfg.engine),
        cfg.voice,
        cfg.rate,
        cfg.pitch,
        cfg.volume,
        cfg.opening_quote,
        cfg.closing_quote,
        if cfg.allow_multiline { "true" } else { "false" }
    );
    std::fs::write(&sidecar, body)
        .map_err(|e| format!("Failed to save dialogue sidecar {:?}: {}", sidecar, e))
}

pub fn load_dialogue_voice_config(path: &Path) -> Option<DialogueVoiceConfig> {
    let sidecar = DialogueVoiceConfig::sidecar_path_for(path);
    let text = std::fs::read_to_string(&sidecar).ok().or_else(|| {
        std::fs::read_to_string(DialogueVoiceConfig::legacy_sidecar_path_for(path)).ok()
    })?;
    let mut engine = None;
    let mut voice = String::new();
    let mut rate = 0;
    let mut pitch = 0;
    let mut volume = 100;
    let mut opening_quote = "\"".to_string();
    let mut closing_quote = "\"".to_string();
    let mut allow_multiline = false;

    for raw in text.lines() {
        let line = raw.trim();
        if line.is_empty() || line.starts_with('#') || line.starts_with(';') {
            continue;
        }
        let Some((k, v)) = line.split_once('=') else {
            continue;
        };
        let key = k.trim().to_ascii_lowercase();
        let val = v.trim();
        match key.as_str() {
            "engine" => engine = parse_engine_input(val),
            "voice" => voice = val.to_string(),
            "rate" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    rate = parsed;
                }
            }
            "pitch" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    pitch = parsed;
                }
            }
            "volume" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    volume = parsed;
                }
            }
            "open_quote" => opening_quote = val.to_string(),
            "close_quote" => closing_quote = val.to_string(),
            "allow_multiline" => {
                allow_multiline = matches!(val.to_ascii_lowercase().as_str(), "true" | "1" | "yes")
            }
            _ => {}
        }
    }

    let engine = engine?;
    if voice.trim().is_empty() || opening_quote.is_empty() || closing_quote.is_empty() {
        return None;
    }
    Some(DialogueVoiceConfig {
        engine,
        voice,
        rate,
        pitch,
        volume,
        opening_quote,
        closing_quote,
        allow_multiline,
    })
}

fn xml_escape_attr(value: &str) -> String {
    value.replace('&', "&amp;").replace('"', "&quot;")
}

fn find_closing(
    text: &str,
    search_from: usize,
    closing_quote: &str,
    allow_multiline: bool,
) -> Option<usize> {
    if allow_multiline {
        return text[search_from..]
            .find(closing_quote)
            .map(|p| search_from + p);
    }
    let line_end = text[search_from..]
        .find('\n')
        .map(|p| search_from + p)
        .unwrap_or(text.len());
    text[search_from..line_end]
        .find(closing_quote)
        .map(|p| search_from + p)
}

pub fn apply_dialogue_tags(text: &str, cfg: &DialogueVoiceConfig) -> String {
    if text.is_empty() || cfg.voice.trim().is_empty() {
        return text.to_string();
    }
    if cfg.opening_quote.is_empty() || cfg.closing_quote.is_empty() {
        return text.to_string();
    }
    if text.to_ascii_lowercase().contains("<voice") {
        return text.to_string();
    }

    let open_tag = format!(
        "<voice engine=\"{}\" voice=\"{}\" rate=\"{}\" pitch=\"{}\" volume=\"{}\">",
        engine_to_key(cfg.engine),
        xml_escape_attr(&cfg.voice),
        cfg.rate,
        cfg.pitch,
        cfg.volume
    );
    let close_tag = "</voice>";
    let mut out = String::with_capacity(text.len() + 128);
    let mut cursor = 0usize;
    let mut replaced_any = false;

    while cursor < text.len() {
        let Some(open_rel) = text[cursor..].find(&cfg.opening_quote) else {
            out.push_str(&text[cursor..]);
            break;
        };
        let open_pos = cursor + open_rel;
        out.push_str(&text[cursor..open_pos]);
        let search_from = open_pos + cfg.opening_quote.len();
        let Some(close_pos) =
            find_closing(text, search_from, &cfg.closing_quote, cfg.allow_multiline)
        else {
            out.push_str(&text[open_pos..]);
            break;
        };
        let end = close_pos + cfg.closing_quote.len();
        out.push_str(&open_tag);
        out.push_str(&text[open_pos..end]);
        out.push_str(close_tag);
        cursor = end;
        replaced_any = true;
    }

    if replaced_any { out } else { text.to_string() }
}

pub fn apply_dialogue_tags_from_sidecar(text: &str, doc_path: Option<&Path>) -> String {
    let Some(path) = doc_path else {
        return text.to_string();
    };
    let Some(cfg) = load_dialogue_voice_config(path) else {
        return text.to_string();
    };
    apply_dialogue_tags(text, &cfg)
}

#[cfg(test)]
mod tests {
    use super::{DialogueVoiceConfig, apply_dialogue_tags};
    use crate::settings::TtsEngine;

    fn cfg(opening: &str, closing: &str, allow_multiline: bool) -> DialogueVoiceConfig {
        DialogueVoiceConfig {
            engine: TtsEngine::Edge,
            voice: "it-IT-ElsaNeural".to_string(),
            rate: 0,
            pitch: 0,
            volume: 100,
            opening_quote: opening.to_string(),
            closing_quote: closing.to_string(),
            allow_multiline,
        }
    }

    #[test]
    fn dialogue_tags_wrap_quoted_text_same_line() {
        let text = r#"Lui disse: "ciao"."#;
        let out = apply_dialogue_tags(text, &cfg("\"", "\"", false));
        assert!(out.contains("<voice"));
        assert!(out.contains("\"ciao\""));
    }

    #[test]
    fn dialogue_tags_do_not_cross_line_when_disabled() {
        let text = "\"ciao\nmondo\"";
        let out = apply_dialogue_tags(text, &cfg("\"", "\"", false));
        assert_eq!(out, text);
    }
}
