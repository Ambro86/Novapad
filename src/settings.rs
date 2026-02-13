use crate::accessibility::to_wide;
use crate::tools::rss::RssSource;
use serde::{Deserialize, Serialize};
#[cfg(not(feature = "standalone"))]
use std::ffi::OsStr;
#[cfg(not(feature = "standalone"))]
use std::os::windows::prelude::*;
use std::path::PathBuf;
#[cfg(not(feature = "standalone"))]
use std::path::{Component, Prefix};
use windows::Win32::Foundation::{ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, HANDLE, HLOCAL, LocalFree};
use windows::Win32::Globalization::GetUserDefaultLocaleName;
use windows::Win32::Security::Cryptography::{
    CRYPT_INTEGER_BLOB, CRYPTPROTECT_UI_FORBIDDEN, CryptProtectData, CryptUnprotectData,
};
#[cfg(not(feature = "standalone"))]
use windows::Win32::Storage::FileSystem::GetDriveTypeW;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Registry::{
    HKEY_CURRENT_USER, KEY_SET_VALUE, REG_OPTION_NON_VOLATILE, REG_SZ, RegCloseKey,
    RegCreateKeyExW, RegDeleteTreeW, RegSetValueExW,
};
use windows::Win32::UI::Shell::{FOLDERID_Documents, SHGetKnownFolderPath};
use windows::core::PCWSTR;

#[cfg(not(feature = "standalone"))]
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
    #[serde(default)]
    pub use_custom_voice: bool,
    #[serde(default)]
    pub custom_voice_engine: Option<TtsEngine>,
    #[serde(default)]
    pub custom_voice: Option<String>,
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
    Odp,
    Odt,
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
    #[serde(rename = "sv")]
    Swedish,
    #[serde(rename = "vi")]
    Vietnamese,
    #[serde(rename = "cs")]
    Czech,
    #[serde(rename = "pl")]
    Polish,
    #[serde(rename = "fr")]
    French,
    #[serde(rename = "sr")]
    Serbian,
    #[serde(rename = "uk")]
    Ukrainian,
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
    #[serde(rename = "sapi4")]
    Sapi4,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum SpellcheckLanguageMode {
    #[serde(rename = "follow")]
    #[default]
    FollowEditorLanguage,
    #[serde(rename = "fixed")]
    FixedLanguage,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum IndentationMode {
    #[serde(rename = "default")]
    #[default]
    Default,
    #[serde(rename = "tabs")]
    Tabs,
    #[serde(rename = "spaces")]
    Spaces,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum PodcastFormat {
    #[serde(rename = "mp3")]
    #[default]
    Mp3,
    #[serde(rename = "wav")]
    Wav,
}

#[derive(Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Default)]
pub enum SubtitleReadMode {
    #[serde(rename = "off")]
    #[default]
    Off,
    #[serde(rename = "nvda")]
    Nvda,
    #[serde(rename = "user")]
    User,
    #[serde(rename = "sapi5")]
    Sapi5,
    #[serde(rename = "sapi4")]
    Sapi4,
    #[serde(rename = "edge")]
    Edge,
    #[serde(rename = "record")]
    Record,
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
    #[serde(default)]
    pub use_dialogue_voice: bool,
    #[serde(default)]
    pub dialogue_voice: String,
    #[serde(default = "default_dialogue_voice_rate")]
    pub dialogue_voice_rate: i32,
    #[serde(default = "default_dialogue_voice_pitch")]
    pub dialogue_voice_pitch: i32,
    #[serde(default = "default_dialogue_voice_volume")]
    pub dialogue_voice_volume: i32,
    #[serde(default = "default_dialogue_tts_engine")]
    pub dialogue_tts_engine: TtsEngine,
    pub tts_only_multilingual: bool,
    pub tts_manual_tuning: bool,
    pub split_on_newline: bool,
    pub word_wrap: bool,
    pub wrap_width: u32,
    pub smart_quotes: bool,
    #[serde(default)]
    pub strip_markdown_keep_bullets: bool,
    pub quote_prefix: String,
    #[serde(default)]
    pub indentation_mode: IndentationMode,
    #[serde(default = "default_indent_tab_width")]
    pub indent_tab_width: u32,
    #[serde(default = "default_indent_space_width")]
    pub indent_space_width: u32,
    pub move_cursor_during_reading: bool,
    pub audiobook_skip_seconds: u32,
    pub audiobook_playback_speed: f32,
    pub audiobook_playback_pitch: f32,
    pub audiobook_playback_volume: f32,
    pub audiobook_m4b_bitrate: u32,
    pub audiobook_split: u32,
    pub audiobook_split_by_text: bool,
    pub audiobook_split_text: String,
    pub audiobook_split_text_requires_newline: bool,
    #[serde(default)]
    pub audiobook_split_by_epub_chapter: bool,
    #[serde(default)]
    pub audiobook_split_by_time: bool,
    #[serde(default = "default_audiobook_split_minutes")]
    pub audiobook_split_minutes: u32,
    #[serde(default = "default_audiobook_split_start_number")]
    pub audiobook_split_start_number: u32,
    #[serde(default)]
    pub subtitle_read_mode: SubtitleReadMode,
    #[serde(default)]
    pub subtitle_offset_ms: i32,
    #[serde(default)]
    pub subtitle_mix_export_on_play: bool,
    #[serde(default)]
    pub subtitle_mix_ducking: bool,
    pub podcast_include_microphone: bool,
    pub podcast_microphone_device_id: String,
    pub podcast_microphone_gain: f32,
    pub podcast_include_system_audio: bool,
    pub podcast_system_device_id: String,
    pub podcast_system_gain: f32,
    pub podcast_output_format: PodcastFormat,
    pub podcast_mp3_bitrate: u32,
    pub podcast_save_folder: String,
    #[serde(default = "default_audiobook_save_folder")]
    pub audiobook_save_folder: String,
    pub podcast_include_video: bool,
    pub podcast_monitor_id: String,
    pub podcast_cache_limit_mb: u32,
    pub podcast_index_api_key: String,
    pub podcast_index_api_secret: String,
    pub youtube_include_timestamps: bool,
    pub last_seen_changelog_version: String,
    pub favorite_voices: Vec<FavoriteVoice>,
    pub dictionary: Vec<DictionaryEntry>,
    pub dictionary_translation_language: String,
    #[serde(default)]
    pub dictionary_search_history: Vec<String>,
    pub wikipedia_language: String,
    pub text_color: u32,
    pub text_size: i32,
    pub tts_rate: i32,
    pub tts_pitch: i32,
    pub tts_volume: i32,
    #[serde(default)]
    pub editor_font_face: String,
    pub show_voice_panel: bool,
    pub show_favorite_panel: bool,
    pub check_updates_on_startup: bool,
    pub prompt_program: String,
    pub prompt_auto_scroll: bool,
    pub prompt_strip_ansi: bool,
    pub prompt_beep_on_idle: bool,
    pub prompt_prevent_sleep: bool,
    pub prompt_announce_lines: bool,
    pub interpreter_path: String,
    pub context_menu_open_with: bool,
    #[serde(default = "default_true")]
    pub confirm_delete_rss_podcast: bool,
    #[serde(default = "default_true")]
    pub announce_unread_rss_podcast_items: bool,
    pub spellcheck_enabled: bool,
    pub spellcheck_language_mode: SpellcheckLanguageMode,
    pub spellcheck_fixed_language: String,
    #[serde(default)]
    pub rss_sources: Vec<RssSource>,
    #[serde(default)]
    pub podcast_sources: Vec<RssSource>,
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
    pub rss_removed_default_vi: Vec<String>,
    #[serde(default)]
    pub rss_default_vi_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_cs: Vec<String>,
    #[serde(default)]
    pub rss_default_cs_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_pl: Vec<String>,
    #[serde(default)]
    pub rss_default_pl_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_fr: Vec<String>,
    #[serde(default)]
    pub rss_default_fr_keys: Vec<String>,
    #[serde(default)]
    pub rss_removed_default_sr: Vec<String>,
    #[serde(default)]
    pub rss_default_sr_keys: Vec<String>,
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
    #[serde(default)]
    pub rss_cooldown_blocked_secs: u64,
    #[serde(default)]
    pub rss_cooldown_not_found_secs: u64,
    #[serde(default)]
    pub rss_cooldown_rate_limited_secs: u64,
    /// Se true, invia crash report anonimi a Sentry
    #[serde(default = "default_true")]
    pub send_crash_reports: bool,
    /// Se true, usa il nome legacy "Novapad" per il titolo finestra e collegamenti.
    #[serde(default)]
    pub use_legacy_name: bool,
}

fn default_true() -> bool {
    true
}

fn default_dialogue_voice_rate() -> i32 {
    0
}

fn default_dialogue_voice_pitch() -> i32 {
    0
}

fn default_dialogue_voice_volume() -> i32 {
    100
}

fn default_dialogue_tts_engine() -> TtsEngine {
    TtsEngine::Edge
}

fn default_indent_tab_width() -> u32 {
    4
}

fn default_indent_space_width() -> u32 {
    4
}

fn default_audiobook_split_minutes() -> u32 {
    5
}

fn default_audiobook_split_start_number() -> u32 {
    1
}

impl Default for AppSettings {
    fn default() -> Self {
        AppSettings {
            open_behavior: OpenBehavior::NewTab,
            language: Language::Italian,
            modified_marker_position: ModifiedMarkerPosition::End,
            tts_engine: TtsEngine::Edge,
            tts_voice: "it-IT-IsabellaNeural".to_string(),
            use_dialogue_voice: false,
            dialogue_voice: String::new(),
            dialogue_voice_rate: 0,
            dialogue_voice_pitch: 0,
            dialogue_voice_volume: 100,
            dialogue_tts_engine: TtsEngine::Edge,
            tts_only_multilingual: false,
            tts_manual_tuning: false,
            split_on_newline: false,
            word_wrap: true,
            wrap_width: 80,
            smart_quotes: false,
            strip_markdown_keep_bullets: false,
            quote_prefix: "> ".to_string(),
            indentation_mode: IndentationMode::Default,
            indent_tab_width: default_indent_tab_width(),
            indent_space_width: default_indent_space_width(),
            move_cursor_during_reading: false,
            audiobook_skip_seconds: 60,
            audiobook_playback_speed: 1.0,
            audiobook_playback_pitch: 0.0,
            audiobook_playback_volume: 1.0,
            audiobook_m4b_bitrate: 128,
            audiobook_split: 0,
            audiobook_split_by_text: false,
            audiobook_split_text: String::new(),
            audiobook_split_text_requires_newline: true,
            audiobook_split_by_epub_chapter: false,
            audiobook_split_by_time: false,
            audiobook_split_minutes: default_audiobook_split_minutes(),
            audiobook_split_start_number: default_audiobook_split_start_number(),
            subtitle_read_mode: SubtitleReadMode::User,
            subtitle_offset_ms: 0,
            subtitle_mix_export_on_play: true,
            subtitle_mix_ducking: false,
            podcast_include_microphone: true,
            podcast_microphone_device_id: PODCAST_DEVICE_DEFAULT.to_string(),
            podcast_microphone_gain: 1.5,
            podcast_include_system_audio: true,
            podcast_system_device_id: PODCAST_DEVICE_DEFAULT.to_string(),
            podcast_system_gain: 1.0,
            podcast_output_format: PodcastFormat::Mp3,
            podcast_mp3_bitrate: 128,
            podcast_save_folder: default_podcast_save_folder(),
            audiobook_save_folder: default_audiobook_save_folder(),
            podcast_include_video: false,
            podcast_monitor_id: String::new(),
            podcast_cache_limit_mb: 500,
            podcast_index_api_key: String::new(),
            podcast_index_api_secret: String::new(),
            youtube_include_timestamps: true,
            last_seen_changelog_version: String::new(),
            favorite_voices: Vec::new(),
            dictionary: Vec::new(),
            dictionary_translation_language: "auto".to_string(),
            dictionary_search_history: Vec::new(),
            wikipedia_language: "auto".to_string(),
            text_color: 0x000000,
            text_size: 12,
            tts_rate: 0,
            tts_pitch: 0,
            tts_volume: 100,
            editor_font_face: String::new(),
            show_voice_panel: false,
            show_favorite_panel: false,
            check_updates_on_startup: true,
            prompt_program: "cmd.exe".to_string(),
            prompt_auto_scroll: true,
            prompt_strip_ansi: true,
            prompt_beep_on_idle: true,
            prompt_prevent_sleep: true,
            prompt_announce_lines: true,
            interpreter_path: "python.exe".to_string(),
            context_menu_open_with: false,
            confirm_delete_rss_podcast: true,
            announce_unread_rss_podcast_items: true,
            spellcheck_enabled: false,
            spellcheck_language_mode: SpellcheckLanguageMode::FollowEditorLanguage,
            spellcheck_fixed_language: "en-US".to_string(),
            rss_sources: Vec::new(),
            rss_removed_default_en: Vec::new(),
            rss_default_en_keys: Vec::new(),
            rss_removed_default_it: Vec::new(),
            rss_default_it_keys: Vec::new(),
            rss_removed_default_es: Vec::new(),
            rss_default_es_keys: Vec::new(),
            rss_removed_default_pt: Vec::new(),
            rss_default_pt_keys: Vec::new(),
            rss_removed_default_vi: Vec::new(),
            rss_default_vi_keys: Vec::new(),
            rss_removed_default_cs: Vec::new(),
            rss_default_cs_keys: Vec::new(),
            rss_removed_default_pl: Vec::new(),
            rss_default_pl_keys: Vec::new(),
            rss_removed_default_fr: Vec::new(),
            rss_default_fr_keys: Vec::new(),
            rss_removed_default_sr: Vec::new(),
            rss_default_sr_keys: Vec::new(),
            podcast_sources: Vec::new(),
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
            rss_cooldown_blocked_secs: 3600,
            rss_cooldown_not_found_secs: 86400,
            rss_cooldown_rate_limited_secs: 300,
            send_crash_reports: true, // Attivato di default
            use_legacy_name: false,
        }
    }
}

pub fn app_display_name(settings: &AppSettings) -> &'static str {
    if settings.use_legacy_name {
        "Novapad"
    } else {
        "Sonarpad"
    }
}

#[cfg(not(feature = "standalone"))]
fn wide(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(Some(0)).collect()
}

#[cfg(not(feature = "standalone"))]
fn is_portable_folder(exe_dir: &std::path::Path) -> bool {
    exe_dir
        .file_name()
        .and_then(|n| n.to_str())
        .map(|n| {
            n.eq_ignore_ascii_case("sonarpad portable")
                || n.eq_ignore_ascii_case("novapad portable")
        })
        .unwrap_or(false)
}

#[cfg(not(feature = "standalone"))]
fn copy_dir_recursive(src: &std::path::Path, dst: &std::path::Path) -> std::io::Result<()> {
    let entries = std::fs::read_dir(src)?;
    for entry_result in entries {
        let entry = match entry_result {
            Ok(entry) => entry,
            Err(e) => {
                crate::log_debug(&format!(
                    "settings_migrate_read_dir_entry_failed {}: {}",
                    src.display(),
                    e
                ));
                continue;
            }
        };
        let file_type = match entry.file_type() {
            Ok(file_type) => file_type,
            Err(e) => {
                crate::log_debug(&format!(
                    "settings_migrate_file_type_failed {}: {}",
                    entry.path().display(),
                    e
                ));
                continue;
            }
        };
        let target = dst.join(entry.file_name());
        if file_type.is_dir() {
            if let Err(e) = std::fs::create_dir_all(&target) {
                crate::log_debug(&format!(
                    "settings_migrate_create_dir_failed {}: {}",
                    target.display(),
                    e
                ));
                continue;
            }
            if let Err(e) = copy_dir_recursive(&entry.path(), &target) {
                crate::log_debug(&format!(
                    "settings_migrate_copy_dir_failed {} -> {}: {}",
                    entry.path().display(),
                    target.display(),
                    e
                ));
            }
        } else if file_type.is_file() {
            if let Err(e) = std::fs::copy(entry.path(), &target) {
                crate::log_debug(&format!(
                    "settings_migrate_copy_file_failed {} -> {}: {}",
                    entry.path().display(),
                    target.display(),
                    e
                ));
            }
        } else {
            crate::log_debug(&format!(
                "settings_migrate_skip_special {}",
                entry.path().display()
            ));
        }
    }
    Ok(())
}

#[cfg(not(feature = "standalone"))]
fn migrate_legacy_settings_dir(legacy_dir: &std::path::Path, new_dir: &std::path::Path) {
    if new_dir.exists() || !legacy_dir.exists() {
        return;
    }
    match std::fs::rename(legacy_dir, new_dir) {
        Ok(_) => {
            crate::log_debug(&format!(
                "settings_migrate_renamed {} -> {}",
                legacy_dir.display(),
                new_dir.display()
            ));
        }
        Err(e) => {
            crate::log_debug(&format!(
                "settings_migrate_rename_failed {} -> {}: {}",
                legacy_dir.display(),
                new_dir.display(),
                e
            ));
            if let Err(e) = std::fs::create_dir_all(new_dir) {
                crate::log_debug(&format!(
                    "settings_migrate_create_dir_failed {}: {}",
                    new_dir.display(),
                    e
                ));
                return;
            }
            if let Err(e) = copy_dir_recursive(legacy_dir, new_dir) {
                crate::log_debug(&format!(
                    "settings_migrate_copy_failed {} -> {}: {}",
                    legacy_dir.display(),
                    new_dir.display(),
                    e
                ));
            }
        }
    }
}

#[cfg(not(feature = "standalone"))]
fn migrate_legacy_log_files(settings_dir: &std::path::Path) {
    let legacy_log = settings_dir.join("Novapad.log");
    let new_log = settings_dir.join("Sonarpad.log");
    if legacy_log.exists() && !new_log.exists() {
        match std::fs::rename(&legacy_log, &new_log) {
            Ok(_) => {
                crate::log_debug(&format!(
                    "settings_migrate_log_renamed {} -> {}",
                    legacy_log.display(),
                    new_log.display()
                ));
            }
            Err(e) => {
                crate::log_debug(&format!(
                    "settings_migrate_log_rename_failed {} -> {}: {}",
                    legacy_log.display(),
                    new_log.display(),
                    e
                ));
                if let Err(e) = std::fs::copy(&legacy_log, &new_log) {
                    crate::log_debug(&format!(
                        "settings_migrate_log_copy_failed {} -> {}: {}",
                        legacy_log.display(),
                        new_log.display(),
                        e
                    ));
                }
            }
        }
    }
    let legacy_lock = settings_dir.join("Novapad.log.lock");
    let new_lock = settings_dir.join("Sonarpad.log.lock");
    if legacy_lock.exists()
        && !new_lock.exists()
        && let Err(e) = std::fs::rename(&legacy_lock, &new_lock)
    {
        crate::log_debug(&format!(
            "settings_migrate_log_lock_rename_failed {} -> {}: {}",
            legacy_lock.display(),
            new_lock.display(),
            e
        ));
    }
}

#[cfg(not(feature = "standalone"))]
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
            crate::log_if_err!(std::fs::remove_file(&probe));
            true
        }
        Err(_) => false,
    }
}

#[cfg(feature = "standalone")]
fn resolve_settings_dir() -> PathBuf {
    let exe_path = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("."));
    let exe_dir = exe_path
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .to_path_buf();

    // Standalone: preferisce config/ locale, ma fallback su %APPDATA% se non scrivibile
    let portable_dir = exe_dir.join("config");
    let appdata_dir = std::env::var_os("APPDATA")
        .map(PathBuf::from)
        .map(|p| p.join("Sonarpad"))
        .unwrap_or_else(|| portable_dir.clone());

    if dir_is_writable(&portable_dir) {
        crate::log_if_err!(std::fs::create_dir_all(&portable_dir));
        portable_dir
    } else {
        crate::log_if_err!(std::fs::create_dir_all(&appdata_dir));
        if dir_is_writable(&appdata_dir) {
            appdata_dir
        } else {
            crate::log_if_err!(std::fs::create_dir_all(&portable_dir));
            portable_dir
        }
    }
}

#[cfg(not(feature = "standalone"))]
fn resolve_settings_dir() -> PathBuf {
    let exe_path = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("."));
    let exe_dir = exe_path
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .to_path_buf();

    // Portable: <exe_dir>\config\settings.json
    let portable_dir = exe_dir.join("config");

    // Non-portable: %APPDATA%\Sonarpad\settings.json
    let appdata_root = std::env::var_os("APPDATA").map(PathBuf::from);
    let appdata_dir = appdata_root
        .as_ref()
        .map(|p| p.join("Sonarpad"))
        .unwrap_or_else(|| portable_dir.clone());
    if let Some(root) = appdata_root.as_ref() {
        let legacy_dir = root.join("Novapad");
        migrate_legacy_settings_dir(&legacy_dir, &appdata_dir);
        migrate_legacy_log_files(&appdata_dir);
    }

    // 1) "sonarpad portable" -> portable forzato
    // 2) drive removibile -> portable
    let preferred_dir = if is_portable_folder(&exe_dir)
        || matches!(exe_drive_type(&exe_path), Some(t) if t == DRIVE_REMOVABLE)
    {
        portable_dir.clone()
    }
    // 3) default -> AppData\Sonarpad
    else {
        appdata_dir
    };

    if dir_is_writable(&preferred_dir) {
        preferred_dir
    } else {
        crate::log_if_err!(std::fs::create_dir_all(&portable_dir));
        portable_dir
    }
}

pub fn settings_dir() -> PathBuf {
    use std::sync::OnceLock;
    static DIR: OnceLock<PathBuf> = OnceLock::new();
    DIR.get_or_init(resolve_settings_dir).clone()
}

fn get_settings_path() -> PathBuf {
    resolve_settings_dir().join("settings.json")
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
        if lower.starts_with("sv") {
            return Language::Swedish;
        }
        if lower.starts_with("vi") {
            return Language::Vietnamese;
        }
        if lower.starts_with("cs") {
            return Language::Czech;
        }
        if lower.starts_with("pl") {
            return Language::Polish;
        }
        if lower.starts_with("fr") {
            return Language::French;
        }
        if lower.starts_with("sr") {
            return Language::Serbian;
        }
        if lower.starts_with("uk") {
            return Language::Ukrainian;
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
    base.push("Sonarpad Recordings");
    base.to_string_lossy().to_string()
}

pub fn default_audiobook_save_folder() -> String {
    let mut base = known_folder_path(&FOLDERID_Documents).unwrap_or_else(|| {
        std::env::var_os("USERPROFILE")
            .map(PathBuf::from)
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
            .join("Documents")
    });
    base.push("Sonarpad Audiobooks");
    base.to_string_lossy().to_string()
}

fn legacy_podcast_save_folder() -> String {
    let mut base = known_folder_path(&FOLDERID_Documents).unwrap_or_else(|| {
        std::env::var_os("USERPROFILE")
            .map(PathBuf::from)
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
            .join("Documents")
    });
    base.push("Novapad Recordings");
    base.to_string_lossy().to_string()
}

fn migrate_legacy_podcast_folder(legacy: &std::path::Path, target: &std::path::Path) {
    if !legacy.exists() || target.exists() {
        return;
    }
    if let Err(e) = std::fs::rename(legacy, target) {
        crate::log_debug(&format!(
            "podcast_folder_rename_failed {} -> {}: {}",
            legacy.display(),
            target.display(),
            e
        ));
        if let Err(e) = copy_dir_recursive_simple(legacy, target) {
            crate::log_debug(&format!(
                "podcast_folder_copy_failed {} -> {}: {}",
                legacy.display(),
                target.display(),
                e
            ));
        } else if let Err(e) = std::fs::remove_dir_all(legacy) {
            crate::log_debug(&format!(
                "podcast_folder_remove_failed {}: {}",
                legacy.display(),
                e
            ));
        }
    }
}

fn copy_dir_recursive_simple(src: &std::path::Path, dst: &std::path::Path) -> std::io::Result<()> {
    let entries = std::fs::read_dir(src)?;
    if !dst.exists() {
        std::fs::create_dir_all(dst)?;
    }
    for entry_result in entries {
        let entry = entry_result?;
        let file_type = entry.file_type()?;
        let target = dst.join(entry.file_name());
        if file_type.is_dir() {
            copy_dir_recursive_simple(&entry.path(), &target)?;
        } else if file_type.is_file() {
            std::fs::copy(entry.path(), &target)?;
        }
    }
    Ok(())
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
        let path = raw.to_string().unwrap_or_default();
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
    if path.exists()
        && let Ok(data) = std::fs::read_to_string(&path)
        && let Ok(settings) = serde_json::from_str(&data)
    {
        let normalized = normalize_settings(settings);
        save_settings(normalized.clone());
        return normalized;
    }

    let normalized = normalize_settings(default_settings);
    save_settings(normalized.clone());
    normalized
}

fn normalize_settings(mut settings: AppSettings) -> AppSettings {
    let valid_indent = [2, 4, 6, 8];
    if !valid_indent.contains(&settings.indent_tab_width) {
        settings.indent_tab_width = default_indent_tab_width();
    }
    if !valid_indent.contains(&settings.indent_space_width) {
        settings.indent_space_width = default_indent_space_width();
    }
    if settings.podcast_save_folder.trim().is_empty() {
        settings.podcast_save_folder = default_podcast_save_folder();
    }
    if settings.audiobook_save_folder.trim().is_empty() {
        settings.audiobook_save_folder = default_audiobook_save_folder();
    }
    if settings
        .podcast_save_folder
        .trim()
        .eq_ignore_ascii_case(&legacy_podcast_save_folder())
    {
        let legacy_path = PathBuf::from(legacy_podcast_save_folder());
        let new_path = PathBuf::from(default_podcast_save_folder());
        migrate_legacy_podcast_folder(&legacy_path, &new_path);
        settings.podcast_save_folder = new_path.to_string_lossy().to_string();
    }
    if settings.podcast_mp3_bitrate == 0 {
        settings.podcast_mp3_bitrate = 128;
    }
    if settings.audiobook_m4b_bitrate == 0 {
        settings.audiobook_m4b_bitrate = 128;
    }
    settings.audiobook_m4b_bitrate = settings.audiobook_m4b_bitrate.clamp(64, 256);
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
    settings.podcast_cache_limit_mb = settings.podcast_cache_limit_mb.clamp(100, 2048);
    if settings.spellcheck_fixed_language.trim().is_empty() {
        settings.spellcheck_fixed_language = "en-US".to_string();
    }
    if settings.dictionary_translation_language.trim().is_empty() {
        settings.dictionary_translation_language = "auto".to_string();
    }
    settings
        .dictionary_search_history
        .retain(|s| !s.trim().is_empty());
    if settings.dictionary_search_history.len() > 30 {
        settings.dictionary_search_history.truncate(30);
    }
    if settings.wikipedia_language.trim().is_empty() {
        settings.wikipedia_language = "auto".to_string();
    }
    settings.rss_cooldown_blocked_secs = settings.rss_cooldown_blocked_secs.clamp(60, 86_400);
    settings.rss_cooldown_not_found_secs = settings.rss_cooldown_not_found_secs.clamp(300, 604_800);
    settings.rss_cooldown_rate_limited_secs =
        settings.rss_cooldown_rate_limited_secs.clamp(30, 3_600);
    settings
}

fn dpapi_protect(data: &[u8]) -> Option<Vec<u8>> {
    unsafe {
        let in_blob = CRYPT_INTEGER_BLOB {
            cbData: data.len() as u32,
            pbData: data.as_ptr() as *mut u8,
        };
        let mut out_blob = CRYPT_INTEGER_BLOB::default();
        let ok = CryptProtectData(
            &in_blob,
            None,
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut out_blob,
        )
        .is_ok();
        if !ok || out_blob.pbData.is_null() {
            return None;
        }
        let out = std::slice::from_raw_parts(out_blob.pbData, out_blob.cbData as usize).to_vec();
        LocalFree(HLOCAL(out_blob.pbData as *mut std::ffi::c_void));
        Some(out)
    }
}

fn dpapi_unprotect(data: &[u8]) -> Option<Vec<u8>> {
    unsafe {
        let in_blob = CRYPT_INTEGER_BLOB {
            cbData: data.len() as u32,
            pbData: data.as_ptr() as *mut u8,
        };
        let mut out_blob = CRYPT_INTEGER_BLOB::default();
        let ok = CryptUnprotectData(
            &in_blob,
            None,
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut out_blob,
        )
        .is_ok();
        if !ok || out_blob.pbData.is_null() {
            return None;
        }
        let out = std::slice::from_raw_parts(out_blob.pbData, out_blob.cbData as usize).to_vec();
        LocalFree(HLOCAL(out_blob.pbData as *mut std::ffi::c_void));
        Some(out)
    }
}

pub fn encrypt_podcast_index_secret(secret: &str) -> String {
    if secret.trim().is_empty() {
        return String::new();
    }
    dpapi_protect(secret.as_bytes())
        .map(hex::encode)
        .unwrap_or_default()
}

pub fn decrypt_podcast_index_secret(secret: &str) -> Option<String> {
    if secret.trim().is_empty() {
        return None;
    }
    let decoded = match hex::decode(secret) {
        Ok(decoded) => decoded,
        Err(_) => return Some(secret.to_string()),
    };
    let bytes = dpapi_unprotect(&decoded)?;
    String::from_utf8(bytes).ok()
}

pub fn save_settings(settings: AppSettings) {
    let path = get_settings_path();
    if let Some(parent) = path.parent()
        && let Err(e) = std::fs::create_dir_all(parent)
    {
        crate::log_debug(&format!("Failed to create settings directory: {}", e));
    }
    match serde_json::to_string_pretty(&settings) {
        Ok(json) => {
            if let Err(e) = std::fs::write(&path, json) {
                crate::log_debug(&format!(
                    "Failed to save settings to {}: {}",
                    path.display(),
                    e
                ));
            }
        }
        Err(e) => {
            crate::log_debug(&format!("Failed to serialize settings: {}", e));
        }
    }
}

pub fn save_settings_with_default_copy(settings: AppSettings, _keep_default_copy: bool) {
    save_settings(settings);
}

const CONTEXT_MENU_EXTENSIONS: &[&str] = &[
    "txt", "md", "pdf", "epub", "mp3", "m4a", "mp4", "aac", "mkv", "avi", "mov", "m4v", "webm",
    "mpg", "mpeg", "ts", "m2ts", "mts", "wmv", "asf", "flv", "vob", "3gp", "flac", "ogg", "opus",
    "wma", "aiff", "m4b", "doc", "docx", "xls", "xlsx", "rtf", "htm", "html", "ppt", "pptx", "py",
    "java", "js", "rb", "pl", "php", "lua", "ps1", "sh", "gdoc", "gsheet", "gslides",
];

pub fn register_application_capabilities() {
    let exe_path = match std::env::current_exe() {
        Ok(path) => path,
        Err(err) => {
            crate::log_debug(&format!("Capabilities: failed to get exe path: {err}"));
            return;
        }
    };
    let exe_path_str = exe_path.to_string_lossy();
    let app_name = "Sonarpad";
    let prog_id_prefix = "Sonarpad.Assoc";

    // 1. Register App Class and Program IDs for each extension
    for ext in CONTEXT_MENU_EXTENSIONS {
        let prog_id = format!("{}.{}", prog_id_prefix, ext);
        let class_key = format!("Software\\Classes\\{}", prog_id);
        if let Some(key) = create_registry_key(&class_key) {
            set_registry_string_value(key, None, &format!("{} Document", app_name));
            unsafe {
                RegCloseKey(key);
            }
        }

        let shell_open_key = format!("{}\\shell\\open\\command", class_key);
        if let Some(key) = create_registry_key(&shell_open_key) {
            set_registry_string_value(key, None, &format!("\"{}\" \"%1\"", exe_path_str));
            unsafe {
                RegCloseKey(key);
            }
        }

        let icon_key = format!("{}\\DefaultIcon", class_key);
        if let Some(key) = create_registry_key(&icon_key) {
            set_registry_string_value(key, None, &format!("\"{}\",0", exe_path_str));
            unsafe {
                RegCloseKey(key);
            }
        }
    }

    // 2. Register Application Capabilities
    let capabilities_key = "Software\\Sonarpad\\Capabilities";
    if let Some(key) = create_registry_key(capabilities_key) {
        set_registry_string_value(key, Some("ApplicationName"), app_name);
        set_registry_string_value(
            key,
            Some("ApplicationDescription"),
            "A modern Notepad with PDF, DOCX, XLS and EPUB support and TTS capabilities.",
        );
        unsafe {
            RegCloseKey(key);
        }
    }

    let file_assoc_key = format!("{}\\FileAssociations", capabilities_key);
    if let Some(key) = create_registry_key(&file_assoc_key) {
        for ext in CONTEXT_MENU_EXTENSIONS {
            let prog_id = format!("{}.{}", prog_id_prefix, ext);
            set_registry_string_value(key, Some(&format!(".{}", ext)), &prog_id);
        }
        unsafe {
            RegCloseKey(key);
        }
    }

    // 3. Register in RegisteredApplications
    if let Some(key) = create_registry_key("Software\\RegisteredApplications") {
        set_registry_string_value(key, Some(app_name), capabilities_key);
        unsafe {
            RegCloseKey(key);
        }
    }
}

pub fn sync_context_menu(settings: &AppSettings) {
    let label = crate::i18n::tr(settings.language, "context_menu.open_with");
    let exe_path = match std::env::current_exe() {
        Ok(path) => path,
        Err(err) => {
            crate::log_debug(&format!("Context menu: failed to get exe path: {err}"));
            return;
        }
    };
    let exe_path_str = exe_path.to_string_lossy();
    let command = format!("\"{}\" \"%1\"", exe_path_str);
    let icon = format!("\"{}\",0", exe_path_str);

    for ext in CONTEXT_MENU_EXTENSIONS {
        let base_key = format!(
            "Software\\Classes\\SystemFileAssociations\\.{}\\shell\\OpenWithSonarpad",
            ext
        );
        if settings.context_menu_open_with {
            create_context_menu_entry(&base_key, &label, &command, &icon);
        } else {
            delete_context_menu_entry(&base_key);
        }
    }
}

pub fn cleanup_legacy_context_menu_entries() {
    for ext in CONTEXT_MENU_EXTENSIONS {
        let base_key = format!(
            "Software\\Classes\\SystemFileAssociations\\.{}\\shell\\OpenWithNovapad",
            ext
        );
        delete_context_menu_entry(&base_key);
    }
}

pub fn sync_start_menu_shortcuts(settings: &AppSettings) {
    let (from, to) = if settings.use_legacy_name {
        ("Sonarpad", "Novapad")
    } else {
        ("Novapad", "Sonarpad")
    };
    for dir in start_menu_program_dirs() {
        rename_start_menu_entries(&dir, from, to);
    }
}

fn start_menu_program_dirs() -> Vec<PathBuf> {
    let mut dirs = Vec::new();
    if let Ok(appdata) = std::env::var("APPDATA") {
        dirs.push(
            PathBuf::from(appdata)
                .join("Microsoft")
                .join("Windows")
                .join("Start Menu")
                .join("Programs"),
        );
    }
    if let Ok(program_data) = std::env::var("ProgramData") {
        dirs.push(
            PathBuf::from(program_data)
                .join("Microsoft")
                .join("Windows")
                .join("Start Menu")
                .join("Programs"),
        );
    }
    dirs
}

fn rename_start_menu_entries(dir: &std::path::Path, from: &str, to: &str) {
    if !dir.exists() {
        return;
    }
    let from_link = dir.join(format!("{from}.lnk"));
    let to_link = dir.join(format!("{to}.lnk"));
    if from_link.exists()
        && !to_link.exists()
        && let Err(e) = std::fs::rename(&from_link, &to_link)
    {
        crate::log_debug(&format!(
            "start_menu_rename_failed {} -> {}: {}",
            from_link.display(),
            to_link.display(),
            e
        ));
    }

    let from_folder = dir.join(from);
    let to_folder = dir.join(to);
    if from_folder.exists()
        && !to_folder.exists()
        && let Err(e) = std::fs::rename(&from_folder, &to_folder)
    {
        crate::log_debug(&format!(
            "start_menu_folder_rename_failed {} -> {}: {}",
            from_folder.display(),
            to_folder.display(),
            e
        ));
        return;
    }

    // If folder exists (either renamed or already there), rename inner shortcut.
    let target_folder = if to_folder.exists() {
        to_folder
    } else {
        from_folder
    };
    if target_folder.exists() {
        let inner_from = target_folder.join(format!("{from}.lnk"));
        let inner_to = target_folder.join(format!("{to}.lnk"));
        if inner_from.exists()
            && !inner_to.exists()
            && let Err(e) = std::fs::rename(&inner_from, &inner_to)
        {
            crate::log_debug(&format!(
                "start_menu_inner_rename_failed {} -> {}: {}",
                inner_from.display(),
                inner_to.display(),
                e
            ));
        }
    }
}

fn create_context_menu_entry(base_key: &str, label: &str, command: &str, icon: &str) {
    if let Some(key) = create_registry_key(base_key) {
        set_registry_string_value(key, None, label);
        unsafe {
            RegCloseKey(key);
        }
    }

    let icon_key = format!("{base_key}\\DefaultIcon");
    if let Some(key) = create_registry_key(&icon_key) {
        set_registry_string_value(key, None, icon);
        unsafe {
            RegCloseKey(key);
        }
    }

    let command_key = format!("{base_key}\\command");
    if let Some(key) = create_registry_key(&command_key) {
        set_registry_string_value(key, None, command);
        unsafe {
            RegCloseKey(key);
        }
    }
}

fn delete_context_menu_entry(base_key: &str) {
    let base_key_wide = to_wide(base_key);
    let status = unsafe { RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(base_key_wide.as_ptr())) };
    if status != ERROR_SUCCESS && status != ERROR_FILE_NOT_FOUND {
        crate::log_debug(&format!(
            "Context menu: failed to delete key {base_key}: {status:?}"
        ));
    }
}

fn create_registry_key(path: &str) -> Option<windows::Win32::System::Registry::HKEY> {
    let path_wide = to_wide(path);
    let mut key = windows::Win32::System::Registry::HKEY::default();
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(path_wide.as_ptr()),
            0,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE,
            None,
            &mut key,
            None,
        )
    };
    if status == ERROR_SUCCESS {
        Some(key)
    } else {
        crate::log_debug(&format!(
            "Context menu: failed to create key {path}: {status:?}"
        ));
        None
    }
}

fn set_registry_string_value(
    key: windows::Win32::System::Registry::HKEY,
    value_name: Option<&str>,
    value: &str,
) -> bool {
    let value_wide = to_wide(value);
    let name_wide;
    let name_ptr = if let Some(name) = value_name {
        name_wide = to_wide(name);
        PCWSTR(name_wide.as_ptr())
    } else {
        PCWSTR::null()
    };
    let value_bytes = unsafe {
        std::slice::from_raw_parts(value_wide.as_ptr() as *const u8, value_wide.len() * 2)
    };
    let status = unsafe { RegSetValueExW(key, name_ptr, 0, REG_SZ, Some(value_bytes)) };
    if status != ERROR_SUCCESS {
        crate::log_debug(&format!("Context menu: failed to set value: {status:?}"));
    }
    status == ERROR_SUCCESS
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

pub fn move_rss_feed_up(settings: &mut AppSettings, index: usize) -> Option<usize> {
    if index == 0 || index >= settings.rss_sources.len() {
        return None;
    }
    settings.rss_sources.swap(index, index - 1);
    Some(index - 1)
}

pub fn move_rss_feed_down(settings: &mut AppSettings, index: usize) -> Option<usize> {
    if index + 1 >= settings.rss_sources.len() {
        return None;
    }
    settings.rss_sources.swap(index, index + 1);
    Some(index + 1)
}

pub fn move_rss_feed_to_top(settings: &mut AppSettings, index: usize) -> Option<usize> {
    move_rss_feed_to_index(settings, index, 0)
}

pub fn move_rss_feed_to_bottom(settings: &mut AppSettings, index: usize) -> Option<usize> {
    let len = settings.rss_sources.len();
    if len == 0 {
        return None;
    }
    move_rss_feed_to_index(settings, index, len - 1)
}

pub fn move_rss_feed_to_index(
    settings: &mut AppSettings,
    index: usize,
    target_index: usize,
) -> Option<usize> {
    let len = settings.rss_sources.len();
    if index >= len {
        return None;
    }
    let target = target_index.min(len.saturating_sub(1));
    if target == index {
        return Some(index);
    }
    let item = settings.rss_sources.remove(index);
    settings.rss_sources.insert(target, item);
    Some(target)
}

pub fn move_podcast_feed_up(settings: &mut AppSettings, index: usize) -> Option<usize> {
    if index == 0 || index >= settings.podcast_sources.len() {
        return None;
    }
    settings.podcast_sources.swap(index, index - 1);
    Some(index - 1)
}

pub fn move_podcast_feed_down(settings: &mut AppSettings, index: usize) -> Option<usize> {
    if index + 1 >= settings.podcast_sources.len() {
        return None;
    }
    settings.podcast_sources.swap(index, index + 1);
    Some(index + 1)
}

pub fn move_podcast_feed_to_top(settings: &mut AppSettings, index: usize) -> Option<usize> {
    move_podcast_feed_to_index(settings, index, 0)
}

pub fn move_podcast_feed_to_bottom(settings: &mut AppSettings, index: usize) -> Option<usize> {
    let len = settings.podcast_sources.len();
    if len == 0 {
        return None;
    }
    move_podcast_feed_to_index(settings, index, len - 1)
}

pub fn move_podcast_feed_to_index(
    settings: &mut AppSettings,
    index: usize,
    target_index: usize,
) -> Option<usize> {
    let len = settings.podcast_sources.len();
    if index >= len {
        return None;
    }
    let target = target_index.min(len.saturating_sub(1));
    if target == index {
        return Some(index);
    }
    let item = settings.podcast_sources.remove(index);
    settings.podcast_sources.insert(target, item);
    Some(target)
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

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum SortOrder {
    TitleAsc,
    TitleDesc,
    DateNewest,
    DateOldest,
}

pub fn sort_rss_sources(settings: &mut AppSettings, order: SortOrder) {
    settings.rss_sources.sort_by(|a, b| match order {
        SortOrder::TitleAsc => {
            let ta = &a.title;
            let tb = &b.title;
            ta.to_lowercase().cmp(&tb.to_lowercase())
        }
        SortOrder::TitleDesc => {
            let ta = &a.title;
            let tb = &b.title;
            tb.to_lowercase().cmp(&ta.to_lowercase())
        }
        SortOrder::DateNewest => b
            .last_updated
            .unwrap_or(0)
            .cmp(&a.last_updated.unwrap_or(0)),
        SortOrder::DateOldest => a
            .last_updated
            .unwrap_or(0)
            .cmp(&b.last_updated.unwrap_or(0)),
    });
}

pub fn sort_podcast_sources(settings: &mut AppSettings, order: SortOrder) {
    settings.podcast_sources.sort_by(|a, b| match order {
        SortOrder::TitleAsc => a.title.to_lowercase().cmp(&b.title.to_lowercase()),
        SortOrder::TitleDesc => b.title.to_lowercase().cmp(&a.title.to_lowercase()),
        SortOrder::DateNewest => b
            .last_updated
            .unwrap_or(0)
            .cmp(&a.last_updated.unwrap_or(0)),
        SortOrder::DateOldest => a
            .last_updated
            .unwrap_or(0)
            .cmp(&b.last_updated.unwrap_or(0)),
    });
}
