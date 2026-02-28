use crate::settings::{AppSettings, TtsEngine};
use std::path::PathBuf;

#[derive(Clone, PartialEq, Eq)]
pub struct DialogueVoiceConfig {
    pub engine: TtsEngine,
    pub voice: String,
    pub use_secondary_voice: bool,
    pub secondary_voice: String,
    pub secondary_engine: TtsEngine,
    pub secondary_rate: i32,
    pub secondary_pitch: i32,
    pub secondary_volume: i32,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub opening_quote: String,
    pub closing_quote: String,
    pub allow_multiline: bool,
}

impl DialogueVoiceConfig {
    pub fn config_path() -> PathBuf {
        crate::settings::settings_dir().join("dialogue.ini")
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

pub fn save_dialogue_voice_config(cfg: &DialogueVoiceConfig) -> Result<(), String> {
    let config_path = DialogueVoiceConfig::config_path();
    if let Some(parent) = config_path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| {
            format!(
                "Failed to create dialogue config folder {:?}: {}",
                parent, e
            )
        })?;
    }
    let body = format!(
        "engine={}\nvoice={}\nuse_secondary_voice={}\nsecondary_voice={}\nsecondary_engine={}\nsecondary_rate={}\nsecondary_pitch={}\nsecondary_volume={}\nrate={}\npitch={}\nvolume={}\nopen_quote={}\nclose_quote={}\nallow_multiline={}\n",
        engine_to_key(cfg.engine),
        cfg.voice,
        if cfg.use_secondary_voice {
            "true"
        } else {
            "false"
        },
        cfg.secondary_voice,
        engine_to_key(cfg.secondary_engine),
        cfg.secondary_rate,
        cfg.secondary_pitch,
        cfg.secondary_volume,
        cfg.rate,
        cfg.pitch,
        cfg.volume,
        cfg.opening_quote,
        cfg.closing_quote,
        if cfg.allow_multiline { "true" } else { "false" }
    );
    std::fs::write(&config_path, body)
        .map_err(|e| format!("Failed to save dialogue config {:?}: {}", config_path, e))
}

pub fn load_dialogue_voice_config() -> Option<DialogueVoiceConfig> {
    let config_path = DialogueVoiceConfig::config_path();
    let text = std::fs::read_to_string(config_path).ok()?;
    let mut engine = None;
    let mut voice = String::new();
    let mut use_secondary_voice = false;
    let mut secondary_voice = String::new();
    let mut secondary_engine = None;
    let mut secondary_rate = 0;
    let mut secondary_pitch = 0;
    let mut secondary_volume = 100;
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
            "use_secondary_voice" => {
                use_secondary_voice =
                    matches!(val.to_ascii_lowercase().as_str(), "true" | "1" | "yes")
            }
            "secondary_voice" => secondary_voice = val.to_string(),
            "secondary_engine" => secondary_engine = parse_engine_input(val),
            "secondary_rate" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    secondary_rate = parsed;
                }
            }
            "secondary_pitch" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    secondary_pitch = parsed;
                }
            }
            "secondary_volume" => {
                if let Ok(parsed) = val.parse::<i32>() {
                    secondary_volume = parsed;
                }
            }
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
        use_secondary_voice,
        secondary_voice,
        secondary_engine: secondary_engine.unwrap_or(engine),
        secondary_rate,
        secondary_pitch,
        secondary_volume,
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

    let build_open_tag =
        |engine: TtsEngine, voice_name: &str, rate: i32, pitch: i32, volume: i32| {
            format!(
                "<voice engine=\"{}\" voice=\"{}\" rate=\"{}\" pitch=\"{}\" volume=\"{}\">",
                engine_to_key(engine),
                xml_escape_attr(voice_name),
                rate,
                pitch,
                volume
            )
        };
    let close_tag = "</voice>";
    let use_secondary = cfg.use_secondary_voice && !cfg.secondary_voice.trim().is_empty();
    let secondary_voice = cfg.secondary_voice.trim();
    let mut out = String::with_capacity(text.len() + 128);
    let mut cursor = 0usize;
    let mut replaced_any = false;
    let mut use_secondary_next = false;

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
        let (engine, selected_voice, rate, pitch, volume) = if use_secondary && use_secondary_next {
            (
                cfg.secondary_engine,
                secondary_voice,
                cfg.secondary_rate,
                cfg.secondary_pitch,
                cfg.secondary_volume,
            )
        } else {
            (
                cfg.engine,
                cfg.voice.as_str(),
                cfg.rate,
                cfg.pitch,
                cfg.volume,
            )
        };
        let open_tag = build_open_tag(engine, selected_voice, rate, pitch, volume);
        out.push_str(&open_tag);
        out.push_str(&text[open_pos..end]);
        out.push_str(close_tag);
        if use_secondary {
            use_secondary_next = !use_secondary_next;
        }
        cursor = end;
        replaced_any = true;
    }

    if replaced_any { out } else { text.to_string() }
}

pub fn apply_dialogue_tags_from_settings(text: &str, settings: &AppSettings) -> String {
    if !settings.use_dialogue_voice {
        return text.to_string();
    }
    let fallback = DialogueVoiceConfig {
        engine: settings.dialogue_tts_engine,
        voice: settings.dialogue_voice.clone(),
        use_secondary_voice: settings.dialogue_use_secondary_voice,
        secondary_voice: settings.dialogue_secondary_voice.clone(),
        secondary_engine: settings.dialogue_secondary_tts_engine,
        secondary_rate: settings.dialogue_secondary_voice_rate,
        secondary_pitch: settings.dialogue_secondary_voice_pitch,
        secondary_volume: settings.dialogue_secondary_voice_volume,
        rate: settings.dialogue_voice_rate,
        pitch: settings.dialogue_voice_pitch,
        volume: settings.dialogue_voice_volume,
        opening_quote: settings.dialogue_opening_quote.clone(),
        closing_quote: settings.dialogue_closing_quote.clone(),
        allow_multiline: settings.dialogue_allow_multiline,
    };
    let cfg = load_dialogue_voice_config().unwrap_or(fallback);
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
            use_secondary_voice: false,
            secondary_voice: String::new(),
            secondary_engine: TtsEngine::Edge,
            secondary_rate: 0,
            secondary_pitch: 0,
            secondary_volume: 100,
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

    #[test]
    fn dialogue_tags_alternate_primary_and_secondary_voice() {
        let text = "\"ciao\" e \"come va\"";
        let mut settings = cfg("\"", "\"", false);
        settings.use_secondary_voice = true;
        settings.secondary_voice = "it-IT-DiegoNeural".to_string();
        settings.secondary_rate = -10;
        settings.secondary_pitch = 6;
        settings.secondary_volume = 130;
        let out = apply_dialogue_tags(text, &settings);
        assert!(out.contains("it-IT-ElsaNeural"));
        assert!(out.contains("it-IT-DiegoNeural"));
        assert!(
            out.contains("voice=\"it-IT-DiegoNeural\" rate=\"-10\" pitch=\"6\" volume=\"130\"")
        );
    }
}
