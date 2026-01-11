use crate::tools::rss::RssSource;
use serde::{Deserialize, Serialize};
use std::ffi::OsStr;
use std::os::windows::prelude::*;
use std::path::PathBuf;
use std::path::{Component, Prefix};
use windows::Win32::Foundation::HANDLE;
use windows::Win32::Globalization::GetUserDefaultLocaleName;
use windows::Win32::Storage::FileSystem::GetDriveTypeW;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::Shell::{FOLDERID_Documents, SHGetKnownFolderPath};

pub const DRIVE_REMOVABLE: u32 = 2;

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

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default, Debug)]
pub enum TextEncoding {
    #[serde(rename = "ansi")]
    Ansi,
    #[serde(rename = "utf8")]
    #[default]
    Utf8,
    #[serde(rename = "utf8bom")]
    Utf8Bom,
    #[serde(rename = "utf16le")]
    Utf16Le,
    #[serde(rename = "utf16be")]
    Utf16Be,
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
    Ppt,
    Pptx,
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
pub enum ModifiedMarkerPosition {
    #[serde(rename = "end")]
    #[default]
    End,
    #[serde(rename = "beginning")]
    Beginning,
    #[serde(other)]
    Unknown,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum TtsEngine {
    #[serde(rename = "edge")]
    #[default]
    Edge,
    #[serde(rename = "sapi5")]
    Sapi5,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum PodcastFormat {
    #[serde(rename = "mp3")]
    #[default]
    Mp3,
    #[serde(rename = "wav")]
    Wav,
}

pub const PODCAST_DEVICE_DEFAULT: &str = "default";

#[derive(Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AppSettings {
    pub open_behavior: OpenBehavior,
    pub language: Language,
    pub modified_marker_position: ModifiedMarkerPosition,
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
    pub podcast_include_microphone: bool,
    pub podcast_microphone_device_id: String,
    pub podcast_microphone_gain: f32,
    pub podcast_include_system_audio: bool,
    pub podcast_system_device_id: String,
    pub podcast_system_gain: f32,
    pub podcast_output_format: PodcastFormat,
    pub podcast_mp3_bitrate: u32,
    pub podcast_save_folder: String,
    pub podcast_include_video: bool,
    pub podcast_monitor_id: String,
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
    pub prompt_announce_lines: bool,
    #[serde(default)]
    pub rss_sources: Vec<RssSource>,
    #[serde(default)]
    pub rss_removed_default_en: Vec<String>,
    #[serde(default)]
    pub rss_default_en_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_it: Vec<String>,
    #[serde(default)]
    pub rss_default_it_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_es: Vec<String>,
    #[serde(default)]
    pub rss_default_es_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_pt: Vec<String>,
    #[serde(default)]
    pub rss_default_pt_keys: Vec<String>,
    #[serde(default)]
    pub rss_global_max_concurrency: usize,
    #[serde(default)]
    pub rss_per_host_max_concurrency: usize,
    #[serde(default)]
    pub rss_per_host_rps: u32,
    #[serde(default)]
    pub rss_per_host_burst: u32,
    #[serde(default)]
    pub rss_max_retries: usize,
    #[serde(default)]
    pub rss_backoff_max_secs: u64,
    #[serde(default)]
    pub rss_initial_page_size: usize,
    #[serde(default)]
    pub rss_next_page_size: usize,
    #[serde(default)]
    pub rss_max_items_per_feed: usize,
    #[serde(default)]
    pub rss_max_excerpt_chars: usize,
}

impl Default for AppSettings {
    fn default() -> Self {
        AppSettings {
            open_behavior: OpenBehavior::NewTab,
            language: Language::Italian,
            modified_marker_position: ModifiedMarkerPosition::End,
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
            podcast_include_microphone: true,
            podcast_microphone_device_id: PODCAST_DEVICE_DEFAULT.to_string(),
            podcast_microphone_gain: 1.5,
            podcast_include_system_audio: true,
            podcast_system_device_id: PODCAST_DEVICE_DEFAULT.to_string(),
            podcast_system_gain: 1.0,
            podcast_output_format: PodcastFormat::Mp3,
            podcast_mp3_bitrate: 128,
            podcast_save_folder: default_podcast_save_folder(),
            podcast_include_video: false,
            podcast_monitor_id: String::new(),
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
            prompt_announce_lines: true,
            rss_sources: Vec::new(),
            rss_removed_default_en: Vec::new(),
            rss_default_en_keys: Vec::new(),
            rss_removed_default_it: Vec::new(),
            rss_default_it_keys: Vec::new(),
            rss_removed_default_es: Vec::new(),
            rss_default_es_keys: Vec::new(),
            rss_removed_default_pt: Vec::new(),
            rss_default_pt_keys: Vec::new(),
            rss_global_max_concurrency: 8,
            rss_per_host_max_concurrency: 2,
            rss_per_host_rps: 1,
            rss_per_host_burst: 2,
            rss_max_retries: 4,
            rss_backoff_max_secs: 120,
            rss_initial_page_size: 100,
            rss_next_page_size: 100,
            rss_max_items_per_feed: 5000,
            rss_max_excerpt_chars: 512,
        }
    }
}

fn wide(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(Some(0)).collect()
}

fn is_portable_folder(exe_dir: &std::path::Path) -> bool {
    exe_dir
        .file_name()
        .and_then(|n| n.to_str())
        .map(|n| n.eq_ignore_ascii_case("novapad portable"))
        .unwrap_or(false)
}

fn exe_drive_type(exe: &std::path::Path) -> Option<u32> {
    match exe.components().next()? {
        Component::Prefix(p) => match p.kind() {
            Prefix::Disk(letter) | Prefix::VerbatimDisk(letter) => {
                let root = format!("{}:\\", letter as char);
                Some(unsafe { GetDriveTypeW(windows::core::PCWSTR(wide(&root).as_ptr())) })
            }
            _ => None,
        },
        _ => None,
    }
}

fn dir_is_writable(dir: &std::path::Path) -> bool {
    if std::fs::create_dir_all(dir).is_err() {
        return false;
    }
    let probe = dir.join(format!(".probe_{}", std::process::id()));
    match std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&probe)
    {
        Ok(_) => {
            let _ = std::fs::remove_file(&probe);
            true
        }
        Err(_) => false,
    }
}

fn resolve_settings_dir() -> PathBuf {
    let exe_path = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("."));
    let exe_dir = exe_path
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .to_path_buf();

    // Portable: <exe_dir>\config\settings.json
    let portable_dir = exe_dir.join("config");

    // Non-portable: %APPDATA%\Novapad\settings.json
    let appdata_dir = std::env::var_os("APPDATA")
        .map(PathBuf::from)
        .map(|p| p.join("Novapad"))
        .unwrap_or_else(|| portable_dir.clone());

    // 1) "novapad portable" -> portable forzato
    let preferred_dir = if is_portable_folder(&exe_dir) {
        portable_dir.clone()
    }
    // 2) drive removibile -> portable
    else if matches!(exe_drive_type(&exe_path), Some(t) if t == DRIVE_REMOVABLE) {
        portable_dir.clone()
    }
    // 3) default -> AppData\Novapad
    else {
        appdata_dir
    };

    let settings_dir = if dir_is_writable(&preferred_dir) {
        preferred_dir
    } else {
        let _ = std::fs::create_dir_all(&portable_dir);
        portable_dir
    };

    settings_dir
}

pub fn settings_dir() -> PathBuf {
    resolve_settings_dir()
}

fn get_settings_path() -> PathBuf {
    resolve_settings_dir().join("settings.json")
}

#[allow(dead_code)]
const PORTABLE_MODE: bool = cfg!(feature = "portable");

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

pub fn default_podcast_save_folder() -> String {
    let mut base = known_folder_path(&FOLDERID_Documents).unwrap_or_else(|| {
        std::env::var_os("USERPROFILE")
            .map(PathBuf::from)
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
            .join("Documents")
    });
    base.push("Novapad Recordings");
    base.to_string_lossy().to_string()
}

fn known_folder_path(folder: &windows::core::GUID) -> Option<PathBuf> {
    unsafe {
        let raw = SHGetKnownFolderPath(
            folder,
            windows::Win32::UI::Shell::KNOWN_FOLDER_FLAG(0),
            HANDLE(0),
        )
        .ok()?;
        if raw.is_null() {
            return None;
        }
        let path = crate::accessibility::from_wide(raw.0);
        CoTaskMemFree(Some(raw.0 as *const _));
        if path.is_empty() {
            None
        } else {
            Some(PathBuf::from(path))
        }
    }
}

pub fn load_settings() -> AppSettings {
    let default_settings = AppSettings {
        language: system_language(),
        ..Default::default()
    };

    let path = get_settings_path();
    if path.exists() {
        if let Ok(data) = std::fs::read_to_string(&path) {
            if let Ok(settings) = serde_json::from_str(&data) {
                return normalize_settings(settings);
            }
        }
    }

    normalize_settings(default_settings)
}

fn normalize_settings(mut settings: AppSettings) -> AppSettings {
    if settings.podcast_save_folder.trim().is_empty() {
        settings.podcast_save_folder = default_podcast_save_folder();
    }
    if settings.podcast_mp3_bitrate == 0 {
        settings.podcast_mp3_bitrate = 128;
    }
    if settings.modified_marker_position == ModifiedMarkerPosition::Unknown {
        settings.modified_marker_position = ModifiedMarkerPosition::End;
    }
    if settings.rss_global_max_concurrency == 0 {
        settings.rss_global_max_concurrency = 8;
    }
    if settings.rss_per_host_max_concurrency == 0 {
        settings.rss_per_host_max_concurrency = 2;
    }
    if settings.rss_per_host_rps == 0 {
        settings.rss_per_host_rps = 1;
    }
    if settings.rss_per_host_burst == 0 {
        settings.rss_per_host_burst = 2;
    }
    if settings.rss_max_retries == 0 {
        settings.rss_max_retries = 4;
    }
    if settings.rss_backoff_max_secs == 0 {
        settings.rss_backoff_max_secs = 120;
    }
    if settings.rss_initial_page_size == 0 {
        settings.rss_initial_page_size = 100;
    }
    if settings.rss_next_page_size == 0 {
        settings.rss_next_page_size = 100;
    }
    if settings.rss_max_items_per_feed == 0 {
        settings.rss_max_items_per_feed = 5000;
    }
    if settings.rss_max_excerpt_chars == 0 {
        settings.rss_max_excerpt_chars = 512;
    }
    settings
}

pub fn save_settings(settings: AppSettings) {
    let path = get_settings_path();
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(&settings) {
        let _ = std::fs::write(path, json);
    }
}

pub fn save_settings_with_default_copy(settings: AppSettings, _keep_default_copy: bool) {
    save_settings(settings);
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

pub fn confirm_save_message(language: Language, title: &str) -> String {
    crate::i18n::tr_f(language, "app.confirm_save", &[("title", title)])
}

pub fn untitled_base(language: Language) -> String {
    crate::i18n::tr(language, "app.untitled_base")
}

pub fn untitled_title(language: Language, number: usize) -> String {
    let base = untitled_base(language);
    if number == 0 {
        base
    } else {
        format!("{} {}", base, number)
    }
}

pub fn recent_missing_message(language: Language) -> String {
    crate::i18n::tr(language, "app.recent_missing")
}
