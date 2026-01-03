use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use windows::Win32::Globalization::GetUserDefaultLocaleName;

pub const TRUSTED_CLIENT_TOKEN: &str = "6A5AA1D4EAFF4E9FB37E23D68491D6F4";
pub const VOICE_LIST_URL: &str =
    "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/voices/list";

#[derive(Clone, Serialize, Deserialize)]
pub struct VoiceInfo {
    pub short_name: String,
    pub locale: String,
    pub is_multilingual: bool,
}

#[derive(Clone, Serialize, Deserialize)]
pub struct FavoriteVoice {
    pub engine: TtsEngine,
    pub short_name: String,
}

#[derive(Clone, Serialize, Deserialize)]
pub struct DictionaryEntry {
    pub original: String,
    pub replacement: String,
}

#[derive(Clone, Serialize, Deserialize)]
pub struct AudiobookResult {
    pub success: bool,
    pub message: String,
}

#[derive(Clone, Copy, PartialEq, Eq, Default)]
pub enum TextEncoding {
    #[default]
    Utf8,
    Utf16Le,
    Utf16Be,
    Windows1252,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Text(TextEncoding),
    Docx,
    Doc,
    Pdf,
    Spreadsheet,
    Epub,
    Html,
    Audiobook,
}

impl Default for FileFormat {
    fn default() -> Self {
        FileFormat::Text(TextEncoding::Utf8)
    }
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum OpenBehavior {
    #[serde(rename = "new_tab")]
    #[default]
    NewTab,
    #[serde(rename = "new_window")]
    NewWindow,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum Language {
    #[serde(rename = "it")]
    #[default]
    Italian,
    #[serde(rename = "en")]
    English,
    #[serde(rename = "es")]
    Spanish,
    #[serde(rename = "pt")]
    Portuguese,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum TtsEngine {
    #[serde(rename = "edge")]
    #[default]
    Edge,
    #[serde(rename = "sapi5")]
    Sapi5,
}

#[derive(Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AppSettings {
    pub open_behavior: OpenBehavior,
    pub language: Language,
    pub settings_in_current_dir: bool,
    pub tts_engine: TtsEngine,
    pub tts_voice: String,
    pub tts_only_multilingual: bool,
    pub split_on_newline: bool,
    pub word_wrap: bool,
    pub wrap_width: u32,
    pub quote_prefix: String,
    pub move_cursor_during_reading: bool,
    pub audiobook_skip_seconds: u32,
    pub audiobook_split: u32,
    pub audiobook_split_by_text: bool,
    pub audiobook_split_text: String,
    pub audiobook_split_text_requires_newline: bool,
    pub youtube_include_timestamps: bool,
    pub last_seen_changelog_version: String,
    pub favorite_voices: Vec<FavoriteVoice>,
    pub dictionary: Vec<DictionaryEntry>,
    pub text_color: u32,
    pub text_size: i32,
    pub tts_rate: i32,
    pub tts_pitch: i32,
    pub tts_volume: i32,
    pub show_voice_panel: bool,
    pub show_favorite_panel: bool,
    pub check_updates_on_startup: bool,
    pub prompt_program: String,
    pub prompt_auto_scroll: bool,
    pub prompt_strip_ansi: bool,
    pub prompt_beep_on_idle: bool,
    pub prompt_prevent_sleep: bool,
}

impl Default for AppSettings {
    fn default() -> Self {
        AppSettings {
            open_behavior: OpenBehavior::NewTab,
            language: Language::Italian,
            settings_in_current_dir: false,
            tts_engine: TtsEngine::Edge,
            tts_voice: "it-IT-IsabellaNeural".to_string(),
            tts_only_multilingual: false,
            split_on_newline: false,
            word_wrap: true,
            wrap_width: 80,
            quote_prefix: "> ".to_string(),
            move_cursor_during_reading: false,
            audiobook_skip_seconds: 60,
            audiobook_split: 0,
            audiobook_split_by_text: false,
            audiobook_split_text: String::new(),
            audiobook_split_text_requires_newline: true,
            youtube_include_timestamps: true,
            last_seen_changelog_version: String::new(),
            favorite_voices: Vec::new(),
            dictionary: Vec::new(),
            text_color: 0x000000,
            text_size: 12,
            tts_rate: 0,
            tts_pitch: 0,
            tts_volume: 100,
            show_voice_panel: false,
            show_favorite_panel: false,
            check_updates_on_startup: true,
            prompt_program: "cmd.exe".to_string(),
            prompt_auto_scroll: true,
            prompt_strip_ansi: true,
            prompt_beep_on_idle: true,
            prompt_prevent_sleep: true,
        }
    }
}

fn settings_store_path_appdata() -> Option<PathBuf> {
    let base = std::env::var_os("APPDATA")?;
    let mut path = PathBuf::from(base);
    path.push("Novapad");
    path.push("settings.json");
    Some(path)
}

fn settings_store_path_current_dir() -> Option<PathBuf> {
    let mut path = std::env::current_dir().ok()?;
    path.push("settings.json");
    Some(path)
}

fn system_language() -> Language {
    let mut buffer = [0u16; 85];
    let len = unsafe { GetUserDefaultLocaleName(&mut buffer) };
    if len > 0 {
        let locale = String::from_utf16_lossy(&buffer[..(len as usize).saturating_sub(1)]);
        let lower = locale.to_lowercase();
        if lower.starts_with("it") {
            return Language::Italian;
        }
        if lower.starts_with("es") {
            return Language::Spanish;
        }
        if lower.starts_with("pt") {
            return Language::Portuguese;
        }
        return Language::English;
    }
    Language::Italian
}

pub fn load_settings() -> AppSettings {
    let default_settings = AppSettings {
        language: system_language(),
        ..Default::default()
    };
    let appdata_path = settings_store_path_appdata();
    let current_path = settings_store_path_current_dir();

    if let Some(path) = appdata_path.as_ref().filter(|path| path.exists()) {
        let settings = std::fs::read_to_string(path)
            .ok()
            .and_then(|data| serde_json::from_str(&data).ok())
            .unwrap_or_else(|| default_settings.clone());

        if settings.settings_in_current_dir {
            if let Some(current) = current_path.as_ref().filter(|path| path.exists()) {
                if let Ok(data) = std::fs::read_to_string(current) {
                    if let Ok(portable) = serde_json::from_str(&data) {
                        return portable;
                    }
                }
            }
        }
        return settings;
    }

    if let Some(path) = current_path.as_ref().filter(|path| path.exists()) {
        if let Ok(data) = std::fs::read_to_string(path) {
            if let Ok(settings) = serde_json::from_str(&data) {
                return settings;
            }
        }
    }

    default_settings
}

pub fn save_settings(settings: AppSettings) {
    let appdata_path = settings_store_path_appdata();
    let current_path = settings_store_path_current_dir();
    let path = if settings.settings_in_current_dir {
        current_path.clone()
    } else {
        appdata_path.clone()
    };
    let Some(path) = path else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let mut wrote = false;
    if let Ok(json) = serde_json::to_string_pretty(&settings) {
        if std::fs::write(&path, json).is_ok() {
            wrote = true;
        }
    }
    if wrote {
        if let (Some(appdata_path), Some(current_path)) = (appdata_path, current_path) {
            if appdata_path != current_path {
                let stale_path = if settings.settings_in_current_dir {
                    appdata_path
                } else {
                    current_path
                };
                if stale_path.exists() {
                    let _ = std::fs::remove_file(stale_path);
                }
            }
        }
    }
}

pub fn untitled_base(language: Language) -> String {
    crate::i18n::tr(language, "app.untitled_base")
}

pub fn untitled_title(language: Language, count: usize) -> String {
    format!("{} {}", untitled_base(language), count)
}

pub fn recent_missing_message(language: Language) -> String {
    crate::i18n::tr(language, "app.recent_missing")
}

pub fn confirm_save_message(language: Language, title: &str) -> String {
    crate::i18n::tr_f(language, "app.confirm_save", &[("title", title)])
}

pub fn confirm_title(language: Language) -> String {
    crate::i18n::tr(language, "app.confirm_title")
}

pub fn error_title(language: Language) -> String {
    crate::i18n::tr(language, "app.error_title")
}

pub fn tts_no_text_message(language: Language) -> String {
    crate::i18n::tr(language, "app.tts_no_text")
}

pub fn audiobook_done_title(language: Language) -> String {
    crate::i18n::tr(language, "app.audiobook_done_title")
}

pub fn info_title(language: Language) -> String {
    crate::i18n::tr(language, "app.info_title")
}

pub fn pdf_loaded_message(language: Language) -> String {
    crate::i18n::tr(language, "app.pdf_loaded")
}

pub fn text_not_found_message(language: Language) -> String {
    crate::i18n::tr(language, "app.text_not_found")
}

pub fn find_title(language: Language) -> String {
    crate::i18n::tr(language, "app.find_title")
}

pub fn error_open_file_message(language: Language, _err: impl std::fmt::Display) -> String {
    crate::i18n::tr_f(
        language,
        "app.error_open_file",
        &[("err", &format!("{_err}"))],
    )
}

pub fn error_save_file_message(language: Language, _err: impl std::fmt::Display) -> String {
    crate::i18n::tr_f(
        language,
        "app.error_save_file",
        &[("err", &format!("{_err}"))],
    )
}
