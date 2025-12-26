use serde::{Deserialize, Serialize};
use std::path::PathBuf;

pub const TRUSTED_CLIENT_TOKEN: &str = "6A5AA1D4EAFF4E9FB37E23D68491D6F4";
pub const VOICE_LIST_URL: &str = "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/voices/list";

#[derive(Clone, Serialize, Deserialize)]
pub struct VoiceInfo {
    pub short_name: String,
    pub locale: String,
    pub is_multilingual: bool,
}

#[derive(Clone, Serialize, Deserialize)]
pub struct AudiobookResult {
    pub success: bool,
    pub message: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum TextEncoding {
    Utf8,
    Utf16Le,
    Utf16Be,
    Windows1252,
}

impl Default for TextEncoding {
    fn default() -> Self {
        TextEncoding::Utf8
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Text(TextEncoding),
    Docx,
    Doc,
    Pdf,
    Spreadsheet,
    Epub,
    Audiobook,
}

impl Default for FileFormat {
    fn default() -> Self {
        FileFormat::Text(TextEncoding::Utf8)
    }
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum OpenBehavior {
    #[serde(rename = "new_tab")]
    NewTab,
    #[serde(rename = "new_window")]
    NewWindow,
}

impl Default for OpenBehavior {
    fn default() -> Self {
        OpenBehavior::NewTab
    }
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum Language {
    #[serde(rename = "it")]
    Italian,
    #[serde(rename = "en")]
    English,
}

impl Default for Language {
    fn default() -> Self {
        Language::Italian
    }
}

#[derive(Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AppSettings {
    pub open_behavior: OpenBehavior,
    pub language: Language,
    pub tts_voice: String,
    pub tts_only_multilingual: bool,
    pub split_on_newline: bool,
    pub word_wrap: bool,
    pub move_cursor_during_reading: bool,
    pub audiobook_skip_seconds: u32,
}

impl Default for AppSettings {
    fn default() -> Self {
        AppSettings {
            open_behavior: OpenBehavior::NewTab,
            language: Language::Italian,
            tts_voice: "it-IT-IsabellaNeural".to_string(),
            tts_only_multilingual: false,
            split_on_newline: true,
            word_wrap: true,
            move_cursor_during_reading: false,
            audiobook_skip_seconds: 60,
        }
    }
}

fn settings_store_path() -> Option<PathBuf> {
    let base = std::env::var_os("APPDATA")?;
    let mut path = PathBuf::from(base);
    path.push("Novapad");
    path.push("settings.json");
    Some(path)
}

pub fn load_settings() -> AppSettings {
    let Some(path) = settings_store_path() else {
        return AppSettings::default();
    };
    let data = std::fs::read_to_string(path).ok();
    let Some(data) = data else {
        return AppSettings::default();
    };
    serde_json::from_str(&data).unwrap_or_default()
}

pub fn save_settings(settings: AppSettings) {
    let Some(path) = settings_store_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(&settings) {
        let _ = std::fs::write(path, json);
    }
}
