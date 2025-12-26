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

pub fn untitled_base(language: Language) -> &'static str {
    match language {
        Language::Italian => "Senza titolo",
        Language::English => "Untitled",
    }
}

pub fn untitled_title(language: Language, count: usize) -> String {
    format!("{} {}", untitled_base(language), count)
}

pub fn recent_missing_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Il file recente non esiste piu'.",
        Language::English => "The recent file no longer exists.",
    }
}

pub fn confirm_save_message(language: Language, title: &str) -> String {
    match language {
        Language::Italian => format!("Il documento \"{}\" e' modificato. Salvare?", title),
        Language::English => format!("The document \"{}\" has been modified. Save?", title),
    }
}

pub fn confirm_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Conferma",
        Language::English => "Confirm",
    }
}

pub fn error_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Errore",
        Language::English => "Error",
    }
}

pub fn tts_no_text_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Non c'e' testo da leggere.",
        Language::English => "There is no text to read.",
    }
}

pub fn audiobook_done_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Audiolibro",
        Language::English => "Audiobook",
    }
}

pub fn info_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Info",
        Language::English => "Info",
    }
}

pub fn pdf_loaded_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "PDF caricato.",
        Language::English => "PDF loaded.",
    }
}

pub fn text_not_found_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Testo non trovato.",
        Language::English => "Text not found.",
    }
}

pub fn find_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Trova",
        Language::English => "Find",
    }
}

pub fn error_open_file_message(language: Language, _err: impl std::fmt::Display) -> String {
    match language {
        Language::Italian => format!("Errore apertura file: {_err}"),
        Language::English => format!("Error opening file: {_err}"),
    }
}

pub fn error_save_file_message(language: Language, _err: impl std::fmt::Display) -> String {
    match language {
        Language::Italian => format!("Errore salvataggio file: {_err}"),
        Language::English => format!("Error saving file: {_err}"),
    }
}