use crate::editor_manager::get_edit_text;
use crate::file_handler::{is_epub_path, read_epub_chapters};
use crate::i18n;
use crate::settings;
use crate::settings::{
    AudiobookPartNamingMode, AudiobookResult, DictionaryEntry, Language, TRUSTED_CLIENT_TOKEN,
    TtsEngine,
};
use crate::{get_active_edit, log_debug, save_audio_dialog, show_error, with_state};
use chrono::Local;
use cpal::Sample;
use futures_util::{SinkExt, StreamExt, future::join_all};
use rand::Rng;
use rodio::{Decoder, OutputStreamBuilder, Sink, Source};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::io::{BufWriter, Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use tokio::sync::mpsc;
use tokio_tungstenite::{
    connect_async, tungstenite, tungstenite::client::IntoClientRequest,
    tungstenite::http::HeaderValue, tungstenite::protocol::Message,
};
use url::Url;
use uuid::Uuid;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Power::{ES_CONTINUOUS, ES_SYSTEM_REQUIRED, SetThreadExecutionState};
use windows::Win32::UI::Controls::RichEdit::{CHARRANGE, EM_EXGETSEL, EM_GETTEXTRANGE, TEXTRANGEW};
use windows::Win32::UI::WindowsAndMessaging::{
    PostMessageW, SendMessageW, WM_APP, WM_GETTEXTLENGTH,
};
use windows::core::PWSTR;

use crate::audio_utils::WavWriter;
use crate::subtitle_wasapi::{decode_mp3_to_pcm, resample_pcm};

pub const WSS_URL_BASE: &str =
    "wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1";
pub const MAX_TTS_TEXT_LEN: usize = 3000;
pub const MAX_TTS_TEXT_LEN_LONG: usize = 2000;
pub const MAX_TTS_FIRST_CHUNK_LEN_LONG: usize = 800;
pub const TTS_LONG_TEXT_THRESHOLD: usize = MAX_TTS_TEXT_LEN;
pub(crate) const PAUSE_TAG_MIN_MS: u32 = 50;
pub(crate) const PAUSE_TAG_MAX_MS: u32 = 60_000;
// Some voices silently truncate long SSML payloads without returning an error.
// Keep Edge chunks conservative to avoid partial audiobook exports.
const EDGE_TTS_MAX_BYTES: usize = 1800;
const KEEP_EDGE_TEMP_AFTER_CONVERSION: bool = false;

pub const WM_TTS_PLAYBACK_DONE: u32 = WM_APP + 3;
pub const WM_TTS_PLAYBACK_ERROR: u32 = WM_APP + 5;
pub const WM_TTS_CHUNK_START: u32 = WM_APP + 7;

pub enum TtsCommand {
    Pause,
    Resume,
    Stop,
}

pub struct TtsSession {
    pub id: u64,
    pub command_tx: mpsc::UnboundedSender<TtsCommand>,
    pub cancel: Arc<AtomicBool>,
    pub paused: bool,
    pub initial_caret_pos: i32,
}

#[derive(Clone)]
pub struct TtsChunk {
    pub text_to_read: String,
    pub original_len: usize,
    pub override_voice: Option<VoiceOverride>,
}

type TtsAudioPacket = (Vec<u8>, usize, String);
type EdgeAudioTx<'a> = Option<&'a mpsc::Sender<Result<TtsAudioPacket, String>>>;

#[derive(Clone)]
pub struct VoiceOverride {
    pub engine: TtsEngine,
    pub voice: String,
    pub rate: Option<i32>,
    pub pitch: Option<i32>,
    pub volume: Option<i32>,
}

fn parse_voice_tag_override(tag: &str, default_engine: TtsEngine) -> Option<VoiceOverride> {
    let raw = tag.trim();
    if raw.is_empty() {
        return None;
    }
    let lower = raw.to_ascii_lowercase();
    let mut engine: Option<TtsEngine> = None;
    let mut voice: Option<String> = None;

    for key in ["engine", "tts"] {
        if let Some(pos) = lower.find(key)
            && let Some(eq_pos) = lower[pos..].find('=')
        {
            let val_start = pos + eq_pos + 1;
            let val = raw[val_start..].trim_start();
            let val = if let Some(rest) = val.strip_prefix('"') {
                rest.split('"').next().unwrap_or("").to_string()
            } else {
                val.split_whitespace().next().unwrap_or("").to_string()
            };
            engine = match val.to_ascii_lowercase().as_str() {
                "edge" => Some(TtsEngine::Edge),
                "sapi4" => Some(TtsEngine::Sapi4),
                "sapi5" => Some(TtsEngine::Sapi5),
                _ => None,
            };
        }
    }
    for key in ["voice", "name"] {
        if let Some(pos) = lower.find(key)
            && let Some(eq_pos) = lower[pos..].find('=')
        {
            let val_start = pos + eq_pos + 1;
            let val = raw[val_start..].trim_start();
            let val = if let Some(rest) = val.strip_prefix('"') {
                rest.split('"').next().unwrap_or("").to_string()
            } else {
                val.split_whitespace().next().unwrap_or("").to_string()
            };
            if !val.is_empty() {
                voice = Some(val);
            }
        }
    }

    let mut rate: Option<i32> = None;
    let mut pitch: Option<i32> = None;
    let mut volume: Option<i32> = None;

    for key in ["rate", "speed", "pitch", "volume"] {
        if let Some(pos) = lower.find(key)
            && let Some(eq_pos) = lower[pos..].find('=')
        {
            let val_start = pos + eq_pos + 1;
            let val = raw[val_start..].trim_start();
            let val = if let Some(rest) = val.strip_prefix('"') {
                rest.split('"').next().unwrap_or("")
            } else {
                val.split_whitespace().next().unwrap_or("")
            };
            if let Ok(parsed) = val.parse::<i32>() {
                match key {
                    "rate" | "speed" => rate = Some(parsed),
                    "pitch" => pitch = Some(parsed),
                    "volume" => volume = Some(parsed),
                    _ => {}
                }
            }
        }
    }

    const KNOWN_KEYS: &[&str] = &[
        "engine", "tts", "voice", "name", "rate", "speed", "pitch", "volume",
    ];
    let is_known_attr = |token: &str| {
        let t = token.to_ascii_lowercase();
        KNOWN_KEYS.iter().any(|k| t.starts_with(&format!("{k}=")))
    };

    let mut tokens: Vec<&str> = raw.split_whitespace().collect();
    if let Some(first) = tokens.first().copied() {
        let first_lower = first.to_ascii_lowercase();
        if engine.is_none() {
            engine = match first_lower.as_str() {
                "edge" => Some(TtsEngine::Edge),
                "sapi4" => Some(TtsEngine::Sapi4),
                "sapi5" => Some(TtsEngine::Sapi5),
                _ => None,
            };
            if engine.is_some() {
                tokens.remove(0);
            }
        }
    }
    if voice.is_none() && !tokens.is_empty() {
        let name_tokens: Vec<&str> = tokens.into_iter().filter(|t| !is_known_attr(t)).collect();
        let merged = name_tokens.join(" ");
        if !merged.is_empty() {
            voice = Some(merged);
        }
    }

    let voice = voice?;
    let engine = engine.unwrap_or(default_engine);

    Some(VoiceOverride {
        engine,
        voice,
        rate,
        pitch,
        volume,
    })
}

pub(crate) fn has_voice_tags(text: &str) -> bool {
    text.to_ascii_lowercase().contains("<voice")
}

fn utf16_len(text: &str) -> usize {
    text.encode_utf16().count()
}

fn decode_basic_xml_entities(text: &str) -> String {
    text.replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
}

pub(crate) fn split_voice_tag_spans(
    text: &str,
    default_engine: TtsEngine,
) -> Vec<(String, Option<VoiceOverride>, usize)> {
    if text.is_empty() {
        return Vec::new();
    }
    let lower = text.to_ascii_lowercase();
    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut pending_len = 0usize;
    let mut current_override: Option<VoiceOverride> = None;

    let mut i = 0usize;
    let bytes = lower.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'<' {
            let remaining = &lower[i..];
            let is_open = remaining.starts_with("<voice");
            let is_close = remaining.starts_with("</voice");
            if (is_open || is_close)
                && let Some(end_rel) = remaining.find('>')
            {
                let end = i + end_rel + 1;
                if i > cursor {
                    let chunk = &text[cursor..i];
                    if !chunk.is_empty() {
                        let mut orig_len = utf16_len(chunk);
                        if pending_len > 0 {
                            orig_len += pending_len;
                            pending_len = 0;
                        }
                        segments.push((
                            decode_basic_xml_entities(chunk),
                            current_override.clone(),
                            orig_len,
                        ));
                    }
                }
                let tag_len = utf16_len(&text[i..end]);
                pending_len += tag_len;
                if is_open {
                    let tag_inner = text[i + "<voice".len()..end - 1].trim();
                    current_override = parse_voice_tag_override(tag_inner, default_engine);
                } else {
                    current_override = None;
                }
                cursor = end;
                i = end;
                continue;
            }
        }
        i += 1;
    }

    if cursor < text.len() {
        let tail = &text[cursor..];
        if !tail.is_empty() {
            let mut orig_len = utf16_len(tail);
            if pending_len > 0 {
                orig_len += pending_len;
            }
            segments.push((
                decode_basic_xml_entities(tail),
                current_override.clone(),
                orig_len,
            ));
        }
    } else if pending_len > 0
        && let Some(last) = segments.last_mut()
    {
        last.2 += pending_len;
    }

    if segments.is_empty() {
        segments.push((text.to_string(), None, utf16_len(text)));
    }
    for (i, (txt, ov, len)) in segments.iter().enumerate() {
        crate::log_debug(&format!(
            "split_voice_tag_spans[{}]: text_preview={:?} override_engine={} override_voice={} rate={:?} pitch={:?} vol={:?} len={}",
            i,
            preview_for_log(txt, 120),
            ov.as_ref()
                .map(|o| match o.engine {
                    TtsEngine::Edge => "edge",
                    TtsEngine::Sapi5 => "sapi5",
                    TtsEngine::Sapi4 => "sapi4",
                })
                .unwrap_or("(none)"),
            ov.as_ref().map(|o| o.voice.as_str()).unwrap_or("(none)"),
            ov.as_ref().and_then(|o| o.rate),
            ov.as_ref().and_then(|o| o.pitch),
            ov.as_ref().and_then(|o| o.volume),
            len,
        ));
    }
    segments
}

pub struct TtsPlaybackOptions {
    pub hwnd: HWND,
    pub engine: TtsEngine,
    pub cleaned: String,
    pub voice: String,
    pub chunks: Vec<TtsChunk>,
    pub initial_caret_pos: i32,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
}

struct TtsQueuedPlayback {
    hwnd: HWND,
    engine: TtsEngine,
    text: String,
    voice: String,
    split_on_newline: bool,
    dictionary: Vec<DictionaryEntry>,
    initial_caret_pos: i32,
    rate: i32,
    pitch: i32,
    volume: i32,
}

pub struct AudiobookCommonOptions<'a> {
    pub voice: &'a str,
    pub output: &'a Path,
    pub progress_hwnd: HWND,
    pub cancel: Arc<AtomicBool>,
    pub language: Language,
    pub part_naming_mode: AudiobookPartNamingMode,
    pub audiobook_bitrate_kbps: u32,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub sapi4_threads: Option<u32>,
}

fn post_audiobook_progress(hwnd: HWND, current: usize) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        if let Err(e) = PostMessageW(hwnd, crate::WM_UPDATE_PROGRESS, WPARAM(current), LPARAM(0)) {
            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
        }
    }
}

fn set_audiobook_progress_total(hwnd: HWND, total: usize) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        if let Err(e) = PostMessageW(
            hwnd,
            crate::app_windows::audiobook_window::WM_SET_PROGRESS_TOTAL,
            WPARAM(total.max(1)),
            LPARAM(0),
        ) {
            crate::log_debug(&format!("Failed to post WM_SET_PROGRESS_TOTAL: {}", e));
        }
    }
}

fn set_audiobook_progress_phase(hwnd: HWND, finalizing: bool) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        if let Err(e) = PostMessageW(
            hwnd,
            crate::app_windows::audiobook_window::WM_SET_PROGRESS_PHASE,
            WPARAM(usize::from(finalizing)),
            LPARAM(0),
        ) {
            crate::log_debug(&format!("Failed to post WM_SET_PROGRESS_PHASE: {}", e));
        }
    }
}

fn post_finalization_progress_range(
    progress_hwnd: HWND,
    progress_10000: u32,
    start: usize,
    span: usize,
    total: usize,
) {
    if progress_hwnd.0 == 0 || total == 0 || span == 0 {
        return;
    }
    let clamped = progress_10000.min(10_000) as usize;
    let offset = (clamped * span) / 10_000;
    let current = (start + offset).min(total.saturating_sub(1));
    post_audiobook_progress(progress_hwnd, current);
}

fn cancelled_message(language: Language) -> String {
    i18n::tr(language, "tts.cancelled")
}

pub fn prevent_sleep(enable: bool) {
    unsafe {
        if enable {
            SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
        } else {
            SetThreadExecutionState(ES_CONTINUOUS);
        }
    }
}

fn post_tts_chunk_offset(hwnd: HWND, session_id: u64, offset: usize) {
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        if let Err(e) = PostMessageW(
            hwnd,
            WM_TTS_CHUNK_START,
            WPARAM(session_id as usize),
            LPARAM(offset as isize),
        ) {
            crate::log_debug(&format!("Failed to post WM_TTS_CHUNK_START: {}", e));
        }
    }
}

pub fn post_tts_error(hwnd: HWND, session_id: u64, message: String) {
    log_debug(&format!("TTS error: {message}"));
    let payload = Box::new(message);
    unsafe {
        if let Err(e) = PostMessageW(
            hwnd,
            WM_TTS_PLAYBACK_ERROR,
            WPARAM(session_id as usize),
            LPARAM(Box::into_raw(payload) as isize),
        ) {
            crate::log_debug(&format!("Failed to post WM_TTS_PLAYBACK_ERROR: {}", e));
        }
    }
}

pub fn start_tts_from_caret(hwnd: HWND) {
    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let (language, split_on_newline, tts_engine, dictionary, tts_rate, tts_pitch, tts_volume) = {
        with_state(hwnd, |state| {
            (
                state.settings.language,
                state.settings.split_on_newline,
                state.settings.tts_engine,
                state.settings.dictionary.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
            )
        })
    }
    .unwrap_or((
        Language::Italian,
        true,
        TtsEngine::Edge,
        Vec::new(),
        0,
        0,
        100,
    ));

    let (mut text, initial_caret_pos) = get_text_from_caret(hwnd_edit);
    if with_state(hwnd, |state| {
        state.tts_sentence_nav_anchor = Some((hwnd_edit, initial_caret_pos.max(0)));
    })
    .is_none()
    {
        crate::log_debug("Failed to update TTS sentence navigation anchor");
    }
    let dialogue_settings =
        { with_state(hwnd, |state| state.settings.clone()) }.unwrap_or_default();
    text = crate::dialogue_voice::apply_dialogue_tags_from_settings(&text, &dialogue_settings);
    if text.trim().is_empty() {
        show_error(hwnd, language, &settings::tts_no_text_message(language));
        return;
    }
    let voice = {
        with_state(hwnd, |state| state.settings.tts_voice.clone())
            .unwrap_or_else(|| "it-IT-IsabellaNeural".to_string())
    };
    let has_tags = has_voice_tags(&text);

    if has_tags && tts_engine != TtsEngine::Edge {
        queue_tts_playback_from_text(TtsQueuedPlayback {
            hwnd,
            engine: tts_engine,
            text,
            voice,
            split_on_newline,
            dictionary,
            initial_caret_pos,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
        });
        return;
    }

    match tts_engine {
        TtsEngine::Edge => queue_tts_playback_from_text(TtsQueuedPlayback {
            hwnd,
            engine: tts_engine,
            text,
            voice,
            split_on_newline,
            dictionary,
            initial_caret_pos,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
        }),
        TtsEngine::Sapi4 => {
            stop_tts_playback(hwnd);
            let voice_idx = if let Some(hash_pos) = voice.find("#") {
                let rest = &voice[hash_pos + 1..];
                if let Some(pipe_pos) = rest.find("|") {
                    rest[..pipe_pos].parse::<i32>().unwrap_or(1)
                } else {
                    rest.parse::<i32>().unwrap_or(1)
                }
            } else {
                1
            };
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(hwnd, |state| {
                    state.tts_session = Some(TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to update TTS session state");
            }
            crate::sapi4_engine::play_sapi4(
                voice_idx, text, tts_rate, tts_pitch, tts_volume, cancel, command_rx,
            );
        }
        TtsEngine::Sapi5 => {
            // Stop any existing playback
            stop_tts_playback(hwnd);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if {
                with_state(hwnd, |state| {
                    state.tts_session = Some(TtsSession {
                        id: state.tts_next_session_id,
                        command_tx,
                        cancel: cancel.clone(),
                        paused: false,
                        initial_caret_pos,
                    });
                    state.tts_next_session_id += 1;
                })
            }
            .is_none()
            {
                crate::log_debug("Failed to update TTS session state");
            }

            let chunks = split_into_tts_chunks(&text, split_on_newline, &dictionary, tts_engine);
            let chunk_strings: Vec<String> = chunks.into_iter().map(|c| c.text_to_read).collect();
            if let Err(e) = crate::sapi5_engine::play_sapi(
                chunk_strings,
                voice,
                tts_rate,
                tts_pitch,
                tts_volume,
                cancel,
                command_rx,
            ) {
                crate::log_debug(&format!("SAPI5 playback error: {}", e));
            }
        }
    }
}

pub fn speak_text_once(hwnd: HWND, text: String) {
    let text = text.trim().to_string();
    if text.is_empty() {
        return;
    }
    let (split_on_newline, tts_engine, dictionary, voice, tts_rate, tts_pitch, tts_volume) =
        with_state(hwnd, |state| {
            (
                state.settings.split_on_newline,
                state.settings.tts_engine,
                state.settings.dictionary.clone(),
                state.settings.tts_voice.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
            )
        })
        .unwrap_or((
            true,
            TtsEngine::Edge,
            Vec::new(),
            "it-IT-IsabellaNeural".to_string(),
            0,
            0,
            100,
        ));

    let initial_caret_pos = 0;
    match tts_engine {
        TtsEngine::Edge => queue_tts_playback_from_text(TtsQueuedPlayback {
            hwnd,
            engine: tts_engine,
            text,
            voice,
            split_on_newline,
            dictionary,
            initial_caret_pos,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
        }),
        TtsEngine::Sapi4 => {
            stop_tts_playback(hwnd);
            let voice_idx = if let Some(hash_pos) = voice.find("#") {
                let rest = &voice[hash_pos + 1..];
                if let Some(pipe_pos) = rest.find("|") {
                    rest[..pipe_pos].parse::<i32>().unwrap_or(1)
                } else {
                    rest.parse::<i32>().unwrap_or(1)
                }
            } else {
                1
            };
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if with_state(hwnd, |state| {
                state.tts_session = Some(TtsSession {
                    id: state.tts_next_session_id,
                    command_tx,
                    cancel: cancel.clone(),
                    paused: false,
                    initial_caret_pos,
                });
                state.tts_next_session_id += 1;
            })
            .is_none()
            {
                crate::log_debug("Failed to update subtitle TTS session state");
            }
            crate::sapi4_engine::play_sapi4(
                voice_idx, text, tts_rate, tts_pitch, tts_volume, cancel, command_rx,
            );
        }
        TtsEngine::Sapi5 => {
            stop_tts_playback(hwnd);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            if with_state(hwnd, |state| {
                state.tts_session = Some(TtsSession {
                    id: state.tts_next_session_id,
                    command_tx,
                    cancel: cancel.clone(),
                    paused: false,
                    initial_caret_pos,
                });
                state.tts_next_session_id += 1;
            })
            .is_none()
            {
                crate::log_debug("Failed to update subtitle TTS session state");
            }

            let chunks = split_into_tts_chunks(&text, split_on_newline, &dictionary, tts_engine);
            let chunk_strings: Vec<String> = chunks.into_iter().map(|c| c.text_to_read).collect();
            if let Err(e) = crate::sapi5_engine::play_sapi(
                chunk_strings,
                voice,
                tts_rate,
                tts_pitch,
                tts_volume,
                cancel,
                command_rx,
            ) {
                crate::log_debug(&format!("SAPI5 subtitle playback error: {}", e));
            }
        }
    }
}

fn queue_tts_playback_from_text(options: TtsQueuedPlayback) {
    std::thread::spawn(move || {
        let chunks = split_into_tts_chunks(
            &options.text,
            options.split_on_newline,
            &options.dictionary,
            options.engine,
        );
        let payload = Box::new(TtsPlaybackOptions {
            hwnd: options.hwnd,
            engine: options.engine,
            cleaned: options.text,
            voice: options.voice,
            chunks,
            initial_caret_pos: options.initial_caret_pos,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
        });
        unsafe {
            let payload_ptr = Box::into_raw(payload);
            if PostMessageW(
                options.hwnd,
                crate::WM_TTS_START,
                WPARAM(0),
                LPARAM(payload_ptr as isize),
            )
            .is_err()
            {
                crate::log_debug("Failed to post WM_TTS_START");
                let _unused_box = Box::from_raw(payload_ptr);
            }
        }
    });
}

pub fn toggle_tts_pause(hwnd: HWND) {
    if {
        with_state(hwnd, |state| {
            let Some(session) = &mut state.tts_session else {
                return;
            };
            if session.paused {
                prevent_sleep(true);
                if let Err(e) = session.command_tx.send(TtsCommand::Resume) {
                    crate::log_debug(&format!("Failed to send Resume command: {}", e));
                }
                session.paused = false;
            } else {
                prevent_sleep(false);
                if let Err(e) = session.command_tx.send(TtsCommand::Pause) {
                    crate::log_debug(&format!("Failed to send Pause command: {}", e));
                }
                session.paused = true;
            }
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access TTS session state for pause/resume");
    }
}

pub fn stop_tts_playback(hwnd: HWND) {
    crate::telemetry::set_tts_active(false);
    prevent_sleep(false);
    if {
        with_state(hwnd, |state| {
            if let Some(session) = &state.tts_session {
                session.cancel.store(true, Ordering::SeqCst);
                if let Err(e) = session.command_tx.send(TtsCommand::Stop) {
                    crate::log_debug(&format!("Failed to send Stop command: {}", e));
                }
            }
            state.tts_session = None;
            state.tts_pending_start_pos = None;
        })
    }
    .is_none()
    {
        crate::log_debug("Failed to access TTS session state for stop");
    }
}

fn handle_tts_command(
    cmd: TtsCommand,
    sink: &Sink,
    cancel_flag: &AtomicBool,
    paused: &mut bool,
) -> bool {
    match cmd {
        TtsCommand::Pause => {
            sink.pause();
            *paused = true;
            false
        }
        TtsCommand::Resume => {
            sink.play();
            *paused = false;
            false
        }
        TtsCommand::Stop => {
            cancel_flag.store(true, Ordering::SeqCst);
            sink.stop();
            true
        }
    }
}

fn temp_wav_path(prefix: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    let id = Uuid::new_v4().simple().to_string();
    path.push(format!("sonarpad_{prefix}_{id}.wav"));
    path
}

fn synthesize_sapi5_bytes(
    text: &str,
    voice: &str,
    rate: i32,
    pitch: i32,
    volume: i32,
    language: Language,
) -> Result<Vec<u8>, String> {
    let path = temp_wav_path("sapi5");
    let cancel = Arc::new(AtomicBool::new(false));
    let chunks = vec![text.to_string()];
    crate::sapi5_engine::speak_sapi_to_file(
        crate::sapi5_engine::SapiExportOptions {
            chunks: &chunks,
            voice_name: voice,
            output_path: &path,
            language,
            rate,
            pitch,
            volume,
            audiobook_bitrate_kbps: 128,
            cancel,
        },
        |_chunk_idx| {},
    )?;
    let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
    if let Err(e) = std::fs::remove_file(&path) {
        crate::log_debug(&format!("Failed to remove temp SAPI5 wav: {}", e));
    }
    Ok(bytes)
}

fn synthesize_sapi4_bytes(
    text: &str,
    voice: &str,
    rate: i32,
    pitch: i32,
    volume: i32,
) -> Result<Vec<u8>, String> {
    let path = temp_wav_path("sapi4");
    let cancel = Arc::new(AtomicBool::new(false));
    let voice_idx = parse_sapi4_voice_index(voice);
    let chunks = vec![text.to_string()];
    crate::sapi4_engine::speak_sapi4_to_file(
        &chunks,
        voice_idx,
        &path,
        crate::sapi4_engine::Sapi4Options {
            rate,
            pitch,
            volume,
            mp3_bitrate_kbps: 128,
            cancel,
        },
        |_chunk_idx| {},
    )?;
    let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
    if let Err(e) = std::fs::remove_file(&path) {
        crate::log_debug(&format!("Failed to remove temp SAPI4 wav: {}", e));
    }
    Ok(bytes)
}

async fn synthesize_segment_bytes(
    engine: TtsEngine,
    text: &str,
    voice: &str,
    rate: i32,
    pitch: i32,
    volume: i32,
    language: Language,
) -> Result<Vec<u8>, String> {
    match engine {
        TtsEngine::Edge => {
            let request_id = Uuid::new_v4().simple().to_string();
            download_audio_chunk(text, voice, &request_id, rate, pitch, volume, language).await
        }
        TtsEngine::Sapi5 => {
            let text = text.to_string();
            let voice = voice.to_string();
            tokio::task::spawn_blocking(move || {
                synthesize_sapi5_bytes(&text, &voice, rate, pitch, volume, language)
            })
            .await
            .map_err(|e| e.to_string())?
        }
        TtsEngine::Sapi4 => {
            let text = text.to_string();
            let voice = voice.to_string();
            tokio::task::spawn_blocking(move || {
                synthesize_sapi4_bytes(&text, &voice, rate, pitch, volume)
            })
            .await
            .map_err(|e| e.to_string())?
        }
    }
}

fn decode_wav_to_pcm(bytes: &[u8]) -> Result<(Vec<f32>, u32, u16), String> {
    let cursor = Cursor::new(bytes.to_vec());
    let decoder = Decoder::new(cursor).map_err(|e| e.to_string())?;
    let sample_rate = decoder.sample_rate();
    let channels = decoder.channels();
    let samples: Vec<f32> = decoder.map(|s| s.to_sample::<f32>()).collect();
    Ok((samples, sample_rate, channels))
}

fn write_wav_from_pcm(
    path: &Path,
    samples: &[f32],
    sample_rate: u32,
    channels: u16,
) -> Result<(), String> {
    let mut writer =
        WavWriter::create(path, sample_rate, channels, 16).map_err(|e| e.to_string())?;
    writer
        .write_samples_f32(samples)
        .map_err(|e| e.to_string())?;
    writer.finalize().map_err(|e| e.to_string())?;
    Ok(())
}

struct SynthesisConfig {
    engine: TtsEngine,
    voice: String,
    rate: i32,
    pitch: i32,
    volume: i32,
    language: Language,
}

#[derive(Clone, Copy)]
struct TargetAudio {
    sample_rate: u32,
    channels: u16,
}

async fn synthesize_segment_to_wav(
    text: &str,
    config: &SynthesisConfig,
    target: TargetAudio,
) -> Result<PathBuf, String> {
    let wav_path = temp_wav_path("mix");
    let bytes = synthesize_segment_bytes(
        config.engine,
        text,
        &config.voice,
        config.rate,
        config.pitch,
        config.volume,
        config.language,
    )
    .await?;
    let (samples, src_rate, src_channels) = if config.engine == TtsEngine::Edge {
        match decode_mp3_to_pcm(&bytes) {
            Ok(v) => v,
            Err(mp3_err) => match decode_wav_to_pcm(&bytes) {
                Ok(v) => v,
                Err(wav_err) => {
                    return Err(format!(
                        "Segment decode failed (mp3='{}', wav='{}')",
                        mp3_err, wav_err
                    ));
                }
            },
        }
    } else {
        decode_wav_to_pcm(&bytes)?
    };
    let resampled = resample_pcm(
        &samples,
        src_rate,
        src_channels,
        target.sample_rate,
        target.channels,
    );
    write_wav_from_pcm(&wav_path, &resampled, target.sample_rate, target.channels)?;
    Ok(wav_path)
}

pub fn start_tts_playback_with_chunks(options: TtsPlaybackOptions) {
    // Record telemetry
    crate::telemetry::record_action("tts_start", format!("chunks={}", options.chunks.len()));
    crate::telemetry::set_tts_active(true);

    stop_tts_playback(options.hwnd);
    prevent_sleep(true);
    if options.chunks.is_empty() {
        return;
    }

    let language =
        { with_state(options.hwnd, |state| state.settings.language) }.unwrap_or_default();
    let (tx, rx) = mpsc::unbounded_channel::<TtsCommand>();
    let cancel = Arc::new(AtomicBool::new(false));
    let cancel_flag = cancel.clone();
    let session_id = {
        with_state(options.hwnd, |state| {
            let id = state.tts_next_session_id;
            state.tts_next_session_id = state.tts_next_session_id.saturating_add(1);
            state.tts_session = Some(TtsSession {
                id,
                command_tx: tx.clone(),
                cancel: cancel.clone(),
                paused: false,
                initial_caret_pos: options.initial_caret_pos,
            });
            id
        })
        .unwrap_or(0)
    };
    let hwnd_copy = options.hwnd;
    let chunks = options.chunks;
    let cleaned = options.cleaned;
    let voice = options.voice;
    let tts_rate = options.rate;
    let tts_pitch = options.pitch;
    let tts_volume = options.volume;
    let tts_engine = options.engine;

    std::thread::spawn(move || {
        // Wrap entire TTS playback in catch_unwind to prevent panics from crashing the app
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            tts_playback_inner(TtsPlaybackRequest {
                hwnd_copy,
                session_id,
                chunks,
                cleaned,
                voice,
                tts_rate,
                tts_pitch,
                tts_volume,
                tts_engine,
                language,
                cancel_flag,
                rx,
            });
        }));

        if let Err(panic_info) = result {
            let panic_msg = if let Some(s) = panic_info.downcast_ref::<&str>() {
                s.to_string()
            } else if let Some(s) = panic_info.downcast_ref::<String>() {
                s.clone()
            } else {
                "Unknown panic in TTS playback".to_string()
            };
            crate::log_debug(&format!("TTS thread panic caught: {}", panic_msg));
            post_tts_error(
                hwnd_copy,
                session_id,
                format!("Audio playback error: {}", panic_msg),
            );
        }
    });
}

struct TtsPlaybackRequest {
    hwnd_copy: HWND,
    session_id: u64,
    chunks: Vec<crate::TtsChunk>,
    cleaned: String,
    voice: String,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    tts_engine: TtsEngine,
    language: Language,
    cancel_flag: Arc<AtomicBool>,
    rx: mpsc::UnboundedReceiver<TtsCommand>,
}

fn tts_progress_char_weight(ch: char) -> u64 {
    match ch {
        ',' => 35,
        ';' | ':' => 45,
        '.' | '!' | '?' => 70,
        '—' | '–' => 35,
        '\n' | '\r' => 75,
        _ => 10,
    }
}

fn build_tts_progress_weight_prefix(text: &str) -> Vec<u64> {
    let mut prefix = Vec::with_capacity(text.len().saturating_add(1));
    let mut total = 0u64;
    prefix.push(total);
    for ch in text.chars() {
        total = total.saturating_add(tts_progress_char_weight(ch));
        prefix.push(total);
    }
    prefix
}

struct TtsPlaybackProgressSource<S> {
    inner: S,
    hwnd: HWND,
    session_id: u64,
    chunk_start_offset: usize,
    chunk_len: usize,
    total_samples: Option<u64>,
    emitted_samples: u64,
    next_report_sample: u64,
    report_interval_samples: u64,
    last_reported_offset: Option<usize>,
    reported_end: bool,
    progress_weight_prefix: Vec<u64>,
    progress_weight_total: u64,
}

impl<S> TtsPlaybackProgressSource<S>
where
    S: Source<Item = f32>,
{
    fn new(
        inner: S,
        hwnd: HWND,
        session_id: u64,
        chunk_start_offset: usize,
        chunk_len: usize,
        progress_text: String,
    ) -> Self {
        let channels = u64::from(inner.channels()).max(1);
        let sample_rate = u64::from(inner.sample_rate()).max(1);
        let total_samples = inner.total_duration().map(|duration| {
            (duration.as_secs_f64() * sample_rate as f64 * channels as f64).round() as u64
        });
        let report_interval_samples = (sample_rate * channels / 5).max(1);
        let progress_weight_prefix = build_tts_progress_weight_prefix(&progress_text);
        let progress_weight_total = progress_weight_prefix.last().copied().unwrap_or(0);

        Self {
            inner,
            hwnd,
            session_id,
            chunk_start_offset,
            chunk_len,
            total_samples: total_samples.filter(|samples| *samples > 0),
            emitted_samples: 0,
            next_report_sample: report_interval_samples,
            report_interval_samples,
            last_reported_offset: None,
            reported_end: false,
            progress_weight_prefix,
            progress_weight_total,
        }
    }

    fn report_absolute_offset(&mut self, absolute_offset: usize) {
        let clamped = self
            .chunk_start_offset
            .saturating_add(self.chunk_len)
            .min(absolute_offset);
        if self.last_reported_offset != Some(clamped) {
            post_tts_chunk_offset(self.hwnd, self.session_id, clamped);
            self.last_reported_offset = Some(clamped);
        }
    }

    fn report_progress(&mut self) {
        if self.last_reported_offset.is_none() {
            self.report_absolute_offset(self.chunk_start_offset);
        }
        let Some(total_samples) = self.total_samples else {
            return;
        };
        if self.chunk_len == 0 || self.emitted_samples < self.next_report_sample {
            return;
        }
        let current = self.emitted_samples.min(total_samples);
        let offset_in_chunk = self.weighted_offset_in_chunk(current, total_samples);
        self.report_absolute_offset(self.chunk_start_offset.saturating_add(offset_in_chunk));
        while self.next_report_sample <= self.emitted_samples {
            self.next_report_sample = self
                .next_report_sample
                .saturating_add(self.report_interval_samples);
        }
    }

    fn weighted_offset_in_chunk(&self, current_samples: u64, total_samples: u64) -> usize {
        if self.chunk_len == 0 {
            return 0;
        }
        if total_samples == 0 || self.progress_weight_total == 0 {
            return self.chunk_len;
        }
        let weighted_target = ((current_samples as u128 * self.progress_weight_total as u128)
            / total_samples as u128)
            .min(self.progress_weight_total as u128) as u64;
        let consumed_chars = self
            .progress_weight_prefix
            .partition_point(|weight| *weight <= weighted_target)
            .saturating_sub(1);
        let total_chars = self.progress_weight_prefix.len().saturating_sub(1);
        if total_chars == 0 {
            return ((current_samples as u128 * self.chunk_len as u128) / total_samples as u128)
                .min(self.chunk_len as u128) as usize;
        }
        ((consumed_chars as u128 * self.chunk_len as u128) / total_chars as u128)
            .min(self.chunk_len as u128) as usize
    }

    fn report_end(&mut self) {
        if !self.reported_end {
            self.report_absolute_offset(self.chunk_start_offset.saturating_add(self.chunk_len));
            self.reported_end = true;
        }
    }
}

impl<S> Iterator for TtsPlaybackProgressSource<S>
where
    S: Source<Item = f32>,
{
    type Item = f32;

    fn next(&mut self) -> Option<Self::Item> {
        let sample = self.inner.next();
        if sample.is_some() {
            self.emitted_samples = self.emitted_samples.saturating_add(1);
            self.report_progress();
        } else {
            self.report_end();
        }
        sample
    }
}

impl<S> Source for TtsPlaybackProgressSource<S>
where
    S: Source<Item = f32>,
{
    fn current_span_len(&self) -> Option<usize> {
        self.inner.current_span_len()
    }

    fn channels(&self) -> u16 {
        self.inner.channels()
    }

    fn sample_rate(&self) -> u32 {
        self.inner.sample_rate()
    }

    fn total_duration(&self) -> Option<Duration> {
        self.inner.total_duration()
    }
}

fn tts_playback_inner(req: TtsPlaybackRequest) {
    let TtsPlaybackRequest {
        hwnd_copy,
        session_id,
        chunks,
        cleaned,
        voice,
        tts_rate,
        tts_pitch,
        tts_volume,
        tts_engine,
        language,
        cancel_flag,
        rx,
    } = req;
    let mut rx = rx;
    let start = Instant::now();
    let mut total_audio_bytes = 0usize;
    let mut first_audio_at: Option<Instant> = None;
    let log_end = |reason: &str, total_bytes: usize, first_audio_at: Option<Instant>| {
        let elapsed_ms = start.elapsed().as_millis();
        if let Some(first) = first_audio_at {
            let first_ms = first.duration_since(start).as_millis();
            log_debug(&format!(
                "TTS end: reason={} elapsed_ms={} audio_bytes={} first_audio_ms={}",
                reason, elapsed_ms, total_bytes, first_ms
            ));
        } else {
            log_debug(&format!(
                "TTS end: reason={} elapsed_ms={} audio_bytes={} first_audio_ms=na",
                reason, elapsed_ms, total_bytes
            ));
        }
    };
    let engine_label = match tts_engine {
        TtsEngine::Edge => "edge",
        TtsEngine::Sapi4 => "sapi4",
        TtsEngine::Sapi5 => "sapi5",
    };
    log_debug(&format!(
        "TTS start: engine={} voice={voice} chunks={} text_len={}",
        engine_label,
        chunks.len(),
        cleaned.len()
    ));
    let stream_handle = match OutputStreamBuilder::open_default_stream() {
        Ok(handle) => handle,
        Err(_) => {
            post_tts_error(
                hwnd_copy,
                session_id,
                "Audio output device not available.".to_string(),
            );
            log_end(
                "output_device_unavailable",
                total_audio_bytes,
                first_audio_at,
            );
            return;
        }
    };
    let sink = Sink::connect_new(stream_handle.mixer());
    let rt = match tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
    {
        Ok(rt) => rt,
        Err(err) => {
            post_tts_error(hwnd_copy, session_id, err.to_string());
            log_end("runtime_init_failed", total_audio_bytes, first_audio_at);
            return;
        }
    };

    let (audio_tx, mut audio_rx) = mpsc::channel::<Result<TtsAudioPacket, String>>(10);
    let cancel_downloader = cancel_flag.clone();
    let chunks_downloader = chunks.clone();
    let voice_downloader = voice.clone();
    let rate = tts_rate;
    let pitch = tts_pitch;
    let volume = tts_volume;

    rt.spawn(async move {
        let all_edge = tts_engine == TtsEngine::Edge
            && chunks_downloader.iter().all(|chunk| {
                chunk
                    .override_voice
                    .as_ref()
                    .map(|v| v.engine == TtsEngine::Edge)
                    .unwrap_or(true)
            });

        if all_edge {
            const WS_RETRY_MAX: usize = 10;
            const EDGE_FRESH_CHUNK_RETRIES: usize = 6;
            let mut next_index = 0usize;
            let mut attempt = 1usize;
            loop {
                if cancel_downloader.load(Ordering::SeqCst) {
                    return;
                }
                let ws_result = download_edge_chunks_ws(
                    &chunks_downloader[next_index..],
                    &voice_downloader,
                    rate,
                    pitch,
                    volume,
                    cancel_downloader.as_ref(),
                    Some(&audio_tx),
                )
                .await;
                match ws_result {
                    Ok(processed_count) => {
                        if processed_count == 0 {
                            let msg = "Edge WS: no progress".to_string();
                            if let Err(err) = audio_tx.send(Err(msg)).await {
                                crate::log_debug(&format!("Failed to send audio error: {:?}", err));
                            }
                            return;
                        }
                        next_index = (next_index + processed_count).min(chunks_downloader.len());
                        if next_index >= chunks_downloader.len() {
                            return;
                        }
                        crate::log_debug(&format!(
                            "Edge WS: partial progress processed={} remaining={} (attempt {}/{})",
                            processed_count,
                            chunks_downloader.len().saturating_sub(next_index),
                            attempt,
                            WS_RETRY_MAX
                        ));
                    }
                    Err(e) => {
                        if cancel_downloader.load(Ordering::SeqCst) {
                            return;
                        }
                        if is_edge_forbidden_error(&e) {
                            crate::log_debug(
                                "Edge WS reuse returned 403; switching to fresh Edge sessions for remaining chunks",
                            );
                            for (offset, chunk_obj) in chunks_downloader[next_index..].iter().enumerate() {
                                if cancel_downloader.load(Ordering::SeqCst) {
                                    return;
                                }
                                let (chunk_voice, chunk_rate, chunk_pitch, chunk_volume) =
                                    if let Some(ov) = &chunk_obj.override_voice {
                                        (
                                            ov.voice.as_str(),
                                            ov.rate.unwrap_or(0),
                                            ov.pitch.unwrap_or(0),
                                            ov.volume.unwrap_or(100),
                                        )
                                    } else {
                                        (voice_downloader.as_str(), rate, pitch, volume)
                                    };
                                let options = EdgeStreamOptions {
                                    voice: chunk_voice,
                                    rate: chunk_rate,
                                    pitch: chunk_pitch,
                                    volume: chunk_volume,
                                    language,
                                    cancel: cancel_downloader.as_ref(),
                                    progress_hwnd: hwnd_copy,
                                    allow_http_fallback: true,
                                };
                                match download_edge_chunk_ws_with_retry(
                                    chunk_obj,
                                    &options,
                                    next_index + offset,
                                    EDGE_FRESH_CHUNK_RETRIES,
                                )
                                .await
                                {
                                    Ok(audio) => {
                                        let len = chunk_obj.original_len;
                                        if audio_tx
                                            .send(Ok((audio, len, chunk_obj.text_to_read.clone())))
                                            .await
                                            .is_err()
                                        {
                                            return;
                                        }
                                    }
                                    Err(fresh_err) => {
                                        if let Err(err) = audio_tx.send(Err(fresh_err)).await {
                                            crate::log_debug(&format!(
                                                "Failed to send audio error: {:?}",
                                                err
                                            ));
                                        }
                                        return;
                                    }
                                }
                            }
                            return;
                        }
                        let retry_forever = is_edge_retry_forever_error(&e);
                        crate::log_debug(&format!(
                            "Edge WS reuse failed (attempt {}/{}): {}",
                            attempt,
                            if retry_forever {
                                "inf".to_string()
                            } else {
                                WS_RETRY_MAX.to_string()
                            },
                            e
                        ));
                        if !retry_forever && attempt >= WS_RETRY_MAX {
                            if let Err(err) = audio_tx.send(Err(e)).await {
                                crate::log_debug(&format!("Failed to send audio error: {:?}", err));
                            }
                            return;
                        }
                        tokio::time::sleep(Duration::from_millis(edge_retry_delay_ms(&e, attempt)))
                            .await;
                        attempt = attempt.saturating_add(1);
                    }
                }
            }
        }

        // Mixed-engine path: synthesise chunks sequentially to avoid
        // COM conflicts between SAPI4/SAPI5 running on parallel threads.
        for chunk_obj in &chunks_downloader {
            if cancel_downloader.load(Ordering::SeqCst) {
                break;
            }
            let (engine, chunk_voice, chunk_rate, chunk_pitch, chunk_volume) =
                if let Some(ov) = &chunk_obj.override_voice {
                    (
                        ov.engine,
                        ov.voice.as_str(),
                        ov.rate.unwrap_or(0),
                        ov.pitch.unwrap_or(0),
                        ov.volume.unwrap_or(100),
                    )
                } else {
                    (tts_engine, voice_downloader.as_str(), rate, pitch, volume)
                };
            let result = synthesize_segment_bytes(
                engine,
                &chunk_obj.text_to_read,
                chunk_voice,
                chunk_rate,
                chunk_pitch,
                chunk_volume,
                language,
            )
            .await;
            match result {
                Ok(data) => {
                    if audio_tx
                        .send(Ok((
                            data,
                            chunk_obj.original_len,
                            chunk_obj.text_to_read.clone(),
                        )))
                        .await
                        .is_err()
                    {
                        break;
                    }
                }
                Err(e) => {
                    if let Err(err) = audio_tx.send(Err(e)).await {
                        crate::log_debug(&format!("Failed to send audio error: {:?}", err));
                    }
                    break;
                }
            }
        }
    });

    let mut appended_any = false;
    let mut paused = false;
    let mut current_offset: usize = 0;
    let mut end_reason = "completed";

    loop {
        if cancel_flag.load(Ordering::SeqCst) {
            end_reason = "cancelled";
            break;
        }

        let packet = rt.block_on(async {
            loop {
                if cancel_flag.load(Ordering::SeqCst) {
                    return None;
                }
                while let Ok(cmd) = rx.try_recv() {
                    if handle_tts_command(cmd, &sink, cancel_flag.as_ref(), &mut paused) {
                        return None;
                    }
                }
                if paused {
                    tokio::time::sleep(Duration::from_millis(100)).await;
                    continue;
                }
                tokio::select! {
                    res = audio_rx.recv() => return res,
                    cmd_opt = rx.recv() => {
                        if let Some(cmd) = cmd_opt
                            && handle_tts_command(cmd, &sink, cancel_flag.as_ref(), &mut paused)
                        {
                            return None;
                        }
                    }
                }
            }
        });

        let Some(res) = packet else {
            if !appended_any {
                end_reason = "no_audio";
                crate::log_debug(&format!(
                    "TTS no audio received (engine={}, cancelled={})",
                    engine_label,
                    cancel_flag.load(Ordering::SeqCst)
                ));
            }
            break;
        };
        let (audio, orig_len, progress_text) = match res {
            Ok(data) => data,
            Err(e) => {
                post_tts_error(hwnd_copy, session_id, e);
                end_reason = "download_error";
                break;
            }
        };

        if audio.is_empty() {
            current_offset = current_offset.saturating_add(orig_len.max(1));
            continue;
        }

        total_audio_bytes = total_audio_bytes.saturating_add(audio.len());
        if first_audio_at.is_none() {
            first_audio_at = Some(Instant::now());
            let first_ms = first_audio_at
                .as_ref()
                .map(|t| t.duration_since(start).as_millis())
                .unwrap_or(0);
            log_debug(&format!("TTS first audio: elapsed_ms={}", first_ms));
        }

        let cursor = std::io::Cursor::new(audio);
        let source = match Decoder::new(cursor) {
            Ok(source) => source,
            Err(_) => {
                post_tts_error(hwnd_copy, session_id, "Failed to decode audio.".to_string());
                end_reason = "decode_error";
                break;
            }
        };
        let progress_source = TtsPlaybackProgressSource::new(
            source,
            hwnd_copy,
            session_id,
            current_offset,
            orig_len,
            progress_text,
        );

        sink.append(progress_source);
        appended_any = true;
        while !sink.empty() {
            if cancel_flag.load(Ordering::SeqCst) {
                sink.stop();
                end_reason = "cancelled";
                break;
            }
            while let Ok(cmd) = rx.try_recv() {
                if handle_tts_command(cmd, &sink, cancel_flag.as_ref(), &mut paused) {
                    end_reason = "stopped";
                    break;
                }
            }
            if cancel_flag.load(Ordering::SeqCst) {
                sink.stop();
                end_reason = "cancelled";
                break;
            }
            std::thread::sleep(Duration::from_millis(50));
        }
        if cancel_flag.load(Ordering::SeqCst) {
            end_reason = "cancelled";
            break;
        }
        current_offset = current_offset.saturating_add(orig_len.max(1));
    }

    if appended_any {
        unsafe {
            if let Err(e) = PostMessageW(
                hwnd_copy,
                WM_TTS_PLAYBACK_DONE,
                WPARAM(session_id as usize),
                LPARAM(0),
            ) {
                crate::log_debug(&format!("Failed to post WM_TTS_PLAYBACK_DONE: {}", e));
            }
        }
    }
    log_end(end_reason, total_audio_bytes, first_audio_at);
}

pub fn play_edge_bytes_async(bytes: Vec<u8>, volume: i32) -> Arc<AtomicBool> {
    let cancel = Arc::new(AtomicBool::new(false));
    let cancel_flag = cancel.clone();
    std::thread::spawn(move || {
        let stream_handle = match OutputStreamBuilder::open_default_stream() {
            Ok(handle) => handle,
            Err(_) => {
                crate::log_debug("Edge TTS: audio output device not available.");
                return;
            }
        };
        let sink = Sink::connect_new(stream_handle.mixer());
        let vol = (volume as f32 / 100.0).clamp(0.0, 1.0);
        sink.set_volume(vol);
        let cursor = std::io::Cursor::new(bytes);
        let source = match Decoder::new(cursor) {
            Ok(source) => source,
            Err(err) => {
                crate::log_debug(&format!("Edge TTS: decoder failed: {}", err));
                return;
            }
        };
        sink.append(source);
        loop {
            if cancel_flag.load(Ordering::SeqCst) {
                sink.stop();
                break;
            }
            if sink.empty() {
                break;
            }
            std::thread::sleep(Duration::from_millis(20));
        }
    });
    cancel
}

fn is_edge_connection_retry_error(err: &str) -> bool {
    let lower = err.to_ascii_lowercase();
    err.contains("os error 10054")
        || lower.contains("connection reset")
        || lower.contains("forcibly closed by the remote host")
}

fn is_edge_forbidden_error(err: &str) -> bool {
    err.to_ascii_lowercase().contains("403 forbidden")
}

fn is_edge_retry_forever_error(err: &str) -> bool {
    is_edge_connection_retry_error(err) || is_edge_forbidden_error(err)
}

fn edge_retry_delay_ms(err: &str, attempt: usize) -> u64 {
    let lower = err.to_ascii_lowercase();
    let base_ms = if is_edge_retry_forever_error(err) || lower.contains("timeout") {
        250
    } else {
        400
    };
    ((attempt as u64) * (base_ms as u64)).min(2000)
}

pub async fn download_audio_chunk(
    text: &str,
    voice: &str,
    request_id: &str,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    language: Language,
) -> Result<Vec<u8>, String> {
    download_audio_chunk_with_cancel(DownloadAudioChunkRequest {
        text,
        voice,
        request_id,
        tts_rate,
        tts_pitch,
        tts_volume,
        language,
        cancel: None,
    })
    .await
}

struct DownloadAudioChunkRequest<'a> {
    text: &'a str,
    voice: &'a str,
    request_id: &'a str,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    language: Language,
    cancel: Option<&'a AtomicBool>,
}

async fn download_audio_chunk_with_cancel(
    request: DownloadAudioChunkRequest<'_>,
) -> Result<Vec<u8>, String> {
    let max_retries = 40;
    let mut attempt = 1usize;

    let last_error = loop {
        if request
            .cancel
            .is_some_and(|flag| flag.load(Ordering::Relaxed))
        {
            return Err(cancelled_message(request.language));
        }
        match download_audio_chunk_attempt(
            request.text,
            request.voice,
            request.request_id,
            request.tts_rate,
            request.tts_pitch,
            request.tts_volume,
        )
        .await
        {
            Ok(data) => return Ok(data),
            Err(e) => {
                let retry_forever = is_edge_retry_forever_error(&e);
                let max_label = if retry_forever {
                    "inf".to_string()
                } else {
                    max_retries.to_string()
                };
                let msg = i18n::tr_f(
                    request.language,
                    "tts.chunk_download_retry",
                    &[
                        ("attempt", &attempt.to_string()),
                        ("max", &max_label),
                        ("err", &e),
                    ],
                );
                log_debug(&msg);
                if !retry_forever && attempt >= max_retries {
                    break e;
                }
                tokio::time::sleep(Duration::from_millis(edge_retry_delay_ms(&e, attempt))).await;
                attempt = attempt.saturating_add(1);
            }
        }
    };
    Err(i18n::tr_f(
        request.language,
        "tts.chunk_download_error",
        &[("err", &last_error)],
    ))
}

async fn download_audio_chunk_attempt(
    text: &str,
    voice: &str,
    request_id: &str,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
) -> Result<Vec<u8>, String> {
    let sec_ms_gec = generate_sec_ms_gec();
    let sec_ms_gec_version = "1-132.0.2917.39";

    let url_str = format!(
        "{}?TrustedClientToken={}&ConnectionId={}&Sec-MS-GEC={}&Sec-MS-GEC-Version={}",
        WSS_URL_BASE, TRUSTED_CLIENT_TOKEN, request_id, sec_ms_gec, sec_ms_gec_version
    );
    let url = Url::parse(&url_str).map_err(|err| err.to_string())?;

    let mut request = url
        .as_str()
        .into_client_request()
        .map_err(|e: tungstenite::Error| e.to_string())?;
    let headers = request.headers_mut();
    headers.insert("Pragma", HeaderValue::from_static("no-cache"));
    headers.insert("Cache-Control", HeaderValue::from_static("no-cache"));
    headers.insert(
        "Origin",
        HeaderValue::from_static("chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold"),
    );
    headers.insert("User-Agent", HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0"));
    headers.insert(
        "Accept-Encoding",
        HeaderValue::from_static("gzip, deflate, br"),
    );
    headers.insert(
        "Accept-Language",
        HeaderValue::from_static("en-US,en;q=0.9"),
    );
    let cookie = format!("muid={};", generate_muid());
    headers.insert(
        "Cookie",
        HeaderValue::from_str(&cookie).map_err(|err| err.to_string())?,
    );

    let connect_timeout = Duration::from_secs(5);
    let (ws_stream, _) = match tokio::time::timeout(connect_timeout, connect_async(request)).await {
        Ok(res) => res.map_err(|e: tungstenite::Error| e.to_string())?,
        Err(_) => {
            return Err("WebSocket connect timeout".to_string());
        }
    };
    let (mut write, mut read): (
        futures_util::stream::SplitSink<_, Message>,
        futures_util::stream::SplitStream<_>,
    ) = ws_stream.split();

    let config_msg = format!(
        "X-Timestamp:{}\r\nContent-Type:application/json; charset=utf-8\r\nPath:speech.config\r\n\r\n{{\"context\":{{\"synthesis\":{{\"audio\":{{\"metadataoptions\":{{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"}},\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}}}}}",
        get_date_string()
    );
    write
        .send(Message::Text(config_msg.into()))
        .await
        .map_err(|e: tungstenite::Error| e.to_string())?;

    let sanitized_text = sanitize_edge_text(text);
    if !is_edge_text_usable(&sanitized_text) {
        crate::log_debug("Edge WS: skipping empty chunk after sanitization (HTTP fallback path).");
        return Ok(Vec::new());
    }
    let ssml = mkssml(&sanitized_text, voice, tts_rate, tts_pitch, tts_volume);
    let ssml_msg = format!(
        "X-RequestId:{}\r\nContent-Type:application/ssml+xml\r\nX-Timestamp:{}Z\r\nPath:ssml\r\n\r\n{}",
        request_id,
        get_date_string(),
        ssml
    );
    write
        .send(Message::Text(ssml_msg.into()))
        .await
        .map_err(|e: tungstenite::Error| e.to_string())?;

    let mut audio_data = Vec::new();
    while let Some(msg) = read.next().await {
        let msg: Message = msg.map_err(|e: tungstenite::Error| e.to_string())?;
        match msg {
            Message::Text(text) if text.contains("Path:turn.end") => {
                break;
            }
            Message::Text(_) => {}
            Message::Binary(data) => match parse_edge_binary_audio_payload(&data) {
                Ok(Some(audio)) => {
                    audio_data.extend_from_slice(&audio);
                }
                Ok(None) => continue,
                Err(err) => {
                    log_debug(&err);
                    continue;
                }
            },
            Message::Close(_) => break,
            _ => {}
        }
    }
    Ok(audio_data)
}

async fn read_edge_audio_turn(
    read: &mut futures_util::stream::SplitStream<
        tokio_tungstenite::WebSocketStream<
            tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>,
        >,
    >,
) -> Result<Vec<u8>, String> {
    let mut audio_data = Vec::new();
    while let Some(msg) = read.next().await {
        let msg: Message = msg.map_err(|e: tungstenite::Error| e.to_string())?;
        match msg {
            Message::Text(text) if text.contains("Path:turn.end") => {
                break;
            }
            Message::Text(_) => {}
            Message::Binary(data) => match parse_edge_binary_audio_payload(&data) {
                Ok(Some(audio)) => {
                    audio_data.extend_from_slice(&audio);
                }
                Ok(None) => continue,
                Err(err) => {
                    log_debug(&err);
                    continue;
                }
            },
            Message::Close(_) => break,
            _ => {}
        }
    }
    Ok(audio_data)
}

async fn read_edge_audio_turn_to_writer(
    read: &mut futures_util::stream::SplitStream<
        tokio_tungstenite::WebSocketStream<
            tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>,
        >,
    >,
    buffer: &mut Vec<u8>,
) -> Result<bool, String> {
    let mut audio_received = false;
    while let Some(msg) = read.next().await {
        let msg: Message = msg.map_err(|e: tungstenite::Error| e.to_string())?;
        match msg {
            Message::Text(text) if text.contains("Path:turn.end") => {
                break;
            }
            Message::Text(_) => {}
            Message::Binary(data) => match parse_edge_binary_audio_payload(&data) {
                Ok(Some(audio)) => {
                    buffer.extend_from_slice(&audio);
                    audio_received = true;
                }
                Ok(None) => continue,
                Err(err) => {
                    log_debug(&err);
                    continue;
                }
            },
            Message::Close(_) => break,
            _ => {}
        }
    }
    Ok(audio_received)
}

async fn download_edge_chunks_ws(
    chunks: &[TtsChunk],
    default_voice: &str,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    cancel: &AtomicBool,
    audio_tx: EdgeAudioTx<'_>,
) -> Result<usize, String> {
    let first_audio_timeout = Duration::from_secs(60);
    const FIRST_AUDIO_RETRIES: usize = 1;
    let sec_ms_gec = generate_sec_ms_gec();
    let sec_ms_gec_version = "1-132.0.2917.39";
    let request_id = Uuid::new_v4().simple().to_string();

    let url_str = format!(
        "{}?TrustedClientToken={}&ConnectionId={}&Sec-MS-GEC={}&Sec-MS-GEC-Version={}",
        WSS_URL_BASE, TRUSTED_CLIENT_TOKEN, request_id, sec_ms_gec, sec_ms_gec_version
    );
    let url = Url::parse(&url_str).map_err(|err| err.to_string())?;

    let mut request = url
        .as_str()
        .into_client_request()
        .map_err(|e: tungstenite::Error| e.to_string())?;
    let headers = request.headers_mut();
    headers.insert("Pragma", HeaderValue::from_static("no-cache"));
    headers.insert("Cache-Control", HeaderValue::from_static("no-cache"));
    headers.insert(
        "Origin",
        HeaderValue::from_static("chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold"),
    );
    headers.insert("User-Agent", HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0"));
    headers.insert(
        "Accept-Encoding",
        HeaderValue::from_static("gzip, deflate, br"),
    );
    headers.insert(
        "Accept-Language",
        HeaderValue::from_static("en-US,en;q=0.9"),
    );
    let cookie = format!("muid={};", generate_muid());
    headers.insert(
        "Cookie",
        HeaderValue::from_str(&cookie).map_err(|err| err.to_string())?,
    );

    let connect_timeout = Duration::from_secs(10);
    let (ws_stream, _) = match tokio::time::timeout(connect_timeout, connect_async(request)).await {
        Ok(res) => res.map_err(|e: tungstenite::Error| e.to_string())?,
        Err(_) => {
            return Err("WebSocket connect timeout".to_string());
        }
    };
    crate::log_debug("Edge WS: connected.");
    let (mut write, mut read) = ws_stream.split();

    let config_msg = format!(
        "X-Timestamp:{}\r\nContent-Type:application/json; charset=utf-8\r\nPath:speech.config\r\n\r\n{{\"context\":{{\"synthesis\":{{\"audio\":{{\"metadataoptions\":{{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"}},\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}}}}}",
        get_date_string()
    );
    write
        .send(Message::Text(config_msg.into()))
        .await
        .map_err(|e: tungstenite::Error| e.to_string())?;

    let mut sent_count = 0usize;
    let mut processed_count = 0usize;
    for (idx, chunk) in chunks.iter().enumerate() {
        if cancel.load(Ordering::Relaxed) {
            return Err("Cancelled".to_string());
        }
        let (voice, chunk_rate, chunk_pitch, chunk_volume) = if let Some(ov) = &chunk.override_voice
        {
            (
                ov.voice.as_str(),
                ov.rate.unwrap_or(0),
                ov.pitch.unwrap_or(0),
                ov.volume.unwrap_or(100),
            )
        } else {
            (default_voice, tts_rate, tts_pitch, tts_volume)
        };
        crate::log_debug(&format!(
            "Edge WS chunk {}: voice={} rate={} pitch={} volume={} override={:?} text={:?}",
            idx,
            voice,
            chunk_rate,
            chunk_pitch,
            chunk_volume,
            chunk
                .override_voice
                .as_ref()
                .map(|ov| (ov.rate, ov.pitch, ov.volume)),
            &chunk.text_to_read,
        ));
        let req_id = Uuid::new_v4().simple().to_string();
        let sanitized_text = sanitize_edge_text(&chunk.text_to_read);
        if !is_edge_text_usable(&sanitized_text) {
            crate::log_debug(&format!(
                "Edge WS: skipping unusable chunk {} text={:?}",
                idx, &chunk.text_to_read
            ));
            processed_count = processed_count.saturating_add(1);
            continue;
        }
        let ssml = mkssml(
            &sanitized_text,
            voice,
            chunk_rate,
            chunk_pitch,
            chunk_volume,
        );
        let ssml_msg = format!(
            "X-RequestId:{}\r\nContent-Type:application/ssml+xml\r\nX-Timestamp:{}Z\r\nPath:ssml\r\n\r\n{}",
            req_id,
            get_date_string(),
            ssml
        );
        if let Err(e) = write.send(Message::Text(ssml_msg.into())).await {
            if processed_count > 0 {
                crate::log_debug(&format!(
                    "Edge WS: write failed after partial progress, processed={}: {}",
                    processed_count, e
                ));
                return Ok(processed_count);
            }
            return Err(e.to_string());
        }
        let audio = if sent_count == 0 {
            let mut audio = None;
            for attempt in 1..=FIRST_AUDIO_RETRIES {
                if cancel.load(Ordering::Relaxed) {
                    return Err("Cancelled".to_string());
                }
                match tokio::time::timeout(first_audio_timeout, read_edge_audio_turn(&mut read))
                    .await
                {
                    Ok(res) => {
                        match res {
                            Ok(data) => {
                                audio = Some(data);
                            }
                            Err(e) => {
                                if processed_count > 0 {
                                    crate::log_debug(&format!(
                                        "Edge WS: read failed after partial progress, processed={}: {}",
                                        processed_count, e
                                    ));
                                    return Ok(processed_count);
                                }
                                return Err(e);
                            }
                        }
                        break;
                    }
                    Err(_) => {
                        crate::log_debug(&format!(
                            "Edge WS: first audio timeout (attempt {}/{} chunk_index={} text_len={} voice={})",
                            attempt,
                            FIRST_AUDIO_RETRIES,
                            idx,
                            chunk.text_to_read.len(),
                            voice
                        ));
                        if attempt < FIRST_AUDIO_RETRIES {
                            tokio::time::sleep(Duration::from_secs((attempt * 2) as u64)).await;
                        }
                    }
                }
            }
            audio.ok_or_else(|| "Edge WS: first audio timeout".to_string())?
        } else {
            match read_edge_audio_turn(&mut read).await {
                Ok(data) => data,
                Err(e) => {
                    if processed_count > 0 {
                        crate::log_debug(&format!(
                            "Edge WS: read failed after partial progress, processed={}: {}",
                            processed_count, e
                        ));
                        return Ok(processed_count);
                    }
                    return Err(e);
                }
            }
        };
        if audio.is_empty() {
            crate::log_debug(&format!(
                "Edge WS: received empty audio chunk, skipping chunk_index={}",
                idx
            ));
            processed_count = processed_count.saturating_add(1);
            continue;
        }
        if let Some(tx) = audio_tx {
            let len = chunk.original_len;
            if tx
                .send(Ok((audio, len, chunk.text_to_read.clone())))
                .await
                .is_err()
            {
                return Ok(processed_count);
            }
            sent_count = sent_count.saturating_add(1);
        } else {
            sent_count = sent_count.saturating_add(1);
        }
        processed_count = processed_count.saturating_add(1);
    }
    Ok(processed_count)
}

struct EdgeStreamOptions<'a> {
    voice: &'a str,
    rate: i32,
    pitch: i32,
    volume: i32,
    language: Language,
    cancel: &'a AtomicBool,
    progress_hwnd: HWND,
    allow_http_fallback: bool,
}

async fn download_edge_chunk_ws(
    chunk: &TtsChunk,
    options: &EdgeStreamOptions<'_>,
    idx: usize,
) -> Result<Vec<u8>, String> {
    let first_audio_timeout = Duration::from_secs(60);
    let sec_ms_gec = generate_sec_ms_gec();
    let sec_ms_gec_version = "1-132.0.2917.39";
    let request_id = Uuid::new_v4().simple().to_string();

    let url_str = format!(
        "{}?TrustedClientToken={}&ConnectionId={}&Sec-MS-GEC={}&Sec-MS-GEC-Version={}",
        WSS_URL_BASE, TRUSTED_CLIENT_TOKEN, request_id, sec_ms_gec, sec_ms_gec_version
    );
    let url = Url::parse(&url_str).map_err(|err| err.to_string())?;

    let mut request = url
        .as_str()
        .into_client_request()
        .map_err(|e: tungstenite::Error| e.to_string())?;
    let headers = request.headers_mut();
    headers.insert("Pragma", HeaderValue::from_static("no-cache"));
    headers.insert("Cache-Control", HeaderValue::from_static("no-cache"));
    headers.insert(
        "Origin",
        HeaderValue::from_static("chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold"),
    );
    headers.insert("User-Agent", HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0"));
    headers.insert(
        "Accept-Encoding",
        HeaderValue::from_static("gzip, deflate, br"),
    );
    headers.insert(
        "Accept-Language",
        HeaderValue::from_static("en-US,en;q=0.9"),
    );
    let cookie = format!("muid={};", generate_muid());
    headers.insert(
        "Cookie",
        HeaderValue::from_str(&cookie).map_err(|err| err.to_string())?,
    );

    let connect_timeout = Duration::from_secs(10);
    let (ws_stream, _) = match tokio::time::timeout(connect_timeout, connect_async(request)).await {
        Ok(res) => res.map_err(|e: tungstenite::Error| e.to_string())?,
        Err(_) => {
            return Err("WebSocket connect timeout".to_string());
        }
    };
    crate::log_debug(&format!("Edge WS: connected (chunk {}).", idx));
    let (mut write, mut read) = ws_stream.split();

    let config_msg = format!(
        "X-Timestamp:{}\r\nContent-Type:application/json; charset=utf-8\r\nPath:speech.config\r\n\r\n{{\"context\":{{\"synthesis\":{{\"audio\":{{\"metadataoptions\":{{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"}},\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}}}}}",
        get_date_string()
    );
    write
        .send(Message::Text(config_msg.into()))
        .await
        .map_err(|e: tungstenite::Error| e.to_string())?;

    if options.cancel.load(Ordering::Relaxed) {
        return Err("Cancelled".to_string());
    }
    let (voice, chunk_rate, chunk_pitch, chunk_volume) = if let Some(ov) = &chunk.override_voice {
        (
            ov.voice.as_str(),
            ov.rate.unwrap_or(0),
            ov.pitch.unwrap_or(0),
            ov.volume.unwrap_or(100),
        )
    } else {
        (options.voice, options.rate, options.pitch, options.volume)
    };
    let req_id = Uuid::new_v4().simple().to_string();
    let sanitized_text = sanitize_edge_text(&chunk.text_to_read);
    if !is_edge_text_usable(&sanitized_text) {
        crate::log_debug(&format!(
            "Edge WS: skipping empty chunk after sanitization (chunk_index={}).",
            idx
        ));
        return Ok(Vec::new());
    }
    let ssml = mkssml(
        &sanitized_text,
        voice,
        chunk_rate,
        chunk_pitch,
        chunk_volume,
    );
    let ssml_msg = format!(
        "X-RequestId:{}\r\nContent-Type:application/ssml+xml\r\nX-Timestamp:{}Z\r\nPath:ssml\r\n\r\n{}",
        req_id,
        get_date_string(),
        ssml
    );
    write
        .send(Message::Text(ssml_msg.into()))
        .await
        .map_err(|e: tungstenite::Error| e.to_string())?;

    let mut chunk_buffer: Vec<u8> = Vec::new();
    let audio_received = match tokio::time::timeout(
        first_audio_timeout,
        read_edge_audio_turn_to_writer(&mut read, &mut chunk_buffer),
    )
    .await
    {
        Ok(res) => res?,
        Err(_) => {
            crate::log_debug(&format!(
                "Edge WS: first audio timeout (chunk_index={} text_len={} voice={})",
                idx,
                chunk.text_to_read.len(),
                voice
            ));
            false
        }
    };

    if !audio_received {
        crate::log_debug(&format!(
            "Edge WS: no audio sent to writer (chunk_index={} bytes_len={} utf16_len={}).",
            idx,
            chunk.text_to_read.len(),
            utf16_len(&chunk.text_to_read)
        ));
        return Err("Edge WS: no audio sent".to_string());
    }

    let text_utf16_len = utf16_len(&chunk.text_to_read);
    let text_bytes_len = chunk.text_to_read.len();
    let audio_bytes_len = chunk_buffer.len();
    crate::log_debug(&format!(
        "Edge WS: chunk completed (chunk_index={} text_bytes={} utf16_len={} audio_bytes={})",
        idx, text_bytes_len, text_utf16_len, audio_bytes_len
    ));
    if audio_bytes_len < 4096 && text_utf16_len > 600 {
        crate::log_debug(&format!(
            "Edge WS: suspiciously small audio payload (chunk_index={} utf16_len={} audio_bytes={})",
            idx, text_utf16_len, audio_bytes_len
        ));
    }

    Ok(chunk_buffer)
}

async fn download_edge_chunk_ws_with_retry(
    chunk: &TtsChunk,
    options: &EdgeStreamOptions<'_>,
    idx: usize,
    max_retries: usize,
) -> Result<Vec<u8>, String> {
    let mut attempt = 1usize;
    let last_err = loop {
        if options.cancel.load(Ordering::Relaxed) {
            return Err(cancelled_message(options.language));
        }
        match download_edge_chunk_ws(chunk, options, idx).await {
            Ok(audio) => return Ok(audio),
            Err(err) => {
                let err_str = err.as_str();
                crate::log_debug(&format!(
                    "Edge WS: chunk download failed (attempt {}/{} chunk_index={}): {}",
                    attempt,
                    if is_edge_retry_forever_error(err_str) {
                        "inf".to_string()
                    } else {
                        max_retries.to_string()
                    },
                    idx,
                    err_str
                ));
                let retry_forever = is_edge_retry_forever_error(err_str);
                if !retry_forever && attempt >= max_retries {
                    break err;
                }
                tokio::time::sleep(Duration::from_millis(edge_retry_delay_ms(err_str, attempt)))
                    .await;
                attempt = attempt.saturating_add(1);
            }
        }
    };
    if options.allow_http_fallback && !options.cancel.load(Ordering::Relaxed) {
        crate::log_debug(&format!(
            "Edge WS: falling back to HTTP chunk (chunk_index={})",
            idx
        ));
        let request_id = Uuid::new_v4().simple().to_string();
        return download_audio_chunk_with_cancel(DownloadAudioChunkRequest {
            text: &chunk.text_to_read,
            voice: options.voice,
            request_id: &request_id,
            tts_rate: options.rate,
            tts_pitch: options.pitch,
            tts_volume: options.volume,
            language: options.language,
            cancel: Some(options.cancel),
        })
        .await;
    }
    Err(last_err)
}

pub(crate) fn format_audiobook_part_filename(
    stem: &str,
    ext: &str,
    number: u32,
    width: usize,
    naming_mode: AudiobookPartNamingMode,
) -> String {
    match naming_mode {
        AudiobookPartNamingMode::TitleNumber => {
            format!("{stem} Part {number:0width$}.{ext}")
        }
        AudiobookPartNamingMode::NumberOnly => format!("{number:0width$}.{ext}"),
        AudiobookPartNamingMode::NumberTitle => format!("{number:0width$} - {stem}.{ext}"),
    }
}

fn time_split_part_output(
    base: &Path,
    part_index: u32,
    start_number: u32,
    naming_mode: AudiobookPartNamingMode,
) -> PathBuf {
    const TIME_SPLIT_PART_WIDTH: usize = 4;
    let stem = base
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let ext = base
        .extension()
        .and_then(|s| s.to_str())
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .unwrap_or("mp3");
    let number = start_number.saturating_add(part_index);
    base.with_file_name(format_audiobook_part_filename(
        stem,
        ext,
        number,
        TIME_SPLIT_PART_WIDTH,
        naming_mode,
    ))
}

fn split_part_output(
    base: &Path,
    part_index: usize,
    total_parts: usize,
    naming_mode: AudiobookPartNamingMode,
) -> PathBuf {
    let stem = base
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let ext = base
        .extension()
        .and_then(|s| s.to_str())
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .unwrap_or("mp3");
    let width = std::cmp::max(2, total_parts.to_string().len());
    let number = (part_index + 1) as u32;
    base.with_file_name(format_audiobook_part_filename(
        stem,
        ext,
        number,
        width,
        naming_mode,
    ))
}

fn parse_part_index_from_stem(
    stem_name: &str,
    stem: &str,
    naming_mode: AudiobookPartNamingMode,
) -> Option<usize> {
    match naming_mode {
        AudiobookPartNamingMode::TitleNumber => stem_name
            .strip_prefix(&format!("{stem} Part "))?
            .parse::<usize>()
            .ok(),
        AudiobookPartNamingMode::NumberOnly => stem_name.parse::<usize>().ok(),
        AudiobookPartNamingMode::NumberTitle => stem_name
            .strip_suffix(&format!(" - {stem}"))?
            .parse::<usize>()
            .ok(),
    }
}

fn collect_split_part_files(
    base: &Path,
    naming_mode: AudiobookPartNamingMode,
) -> Result<Vec<PathBuf>, String> {
    let stem = base
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook");
    let ext = base
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
        .unwrap_or_else(|| "m4b".to_string());
    let parent = base
        .parent()
        .ok_or_else(|| "Output path has no parent directory".to_string())?;
    let mut found: Vec<(usize, PathBuf)> = Vec::new();
    for entry in std::fs::read_dir(parent).map_err(|e| e.to_string())? {
        let entry = entry.map_err(|e| e.to_string())?;
        let path = entry.path();
        let Some(stem_name) = path.file_stem().and_then(|s| s.to_str()) else {
            continue;
        };
        let Some(path_ext) = path.extension().and_then(|s| s.to_str()) else {
            continue;
        };
        if !path_ext.eq_ignore_ascii_case(&ext) {
            continue;
        }
        let Some(idx) = parse_part_index_from_stem(stem_name, stem, naming_mode) else {
            continue;
        };
        if idx == 0 {
            continue;
        }
        found.push((idx, path));
    }
    found.sort_by_key(|(idx, _)| *idx);
    Ok(found.into_iter().map(|(_, p)| p).collect())
}

fn merge_m4b_parts_with_chapters(
    part_files: &[PathBuf],
    output: &Path,
    chapter_titles: Option<&[String]>,
) -> Result<(), String> {
    crate::ffmpeg_export::merge_audio_files_with_chapters_copy(part_files, output, chapter_titles)
}

fn split_parts_output_in_subfolder(output: &Path) -> PathBuf {
    let stem = output
        .file_stem()
        .and_then(|s| s.to_str())
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .unwrap_or("audiobook");
    let file_name = output
        .file_name()
        .map(|s| s.to_owned())
        .unwrap_or_else(|| "audiobook.mp3".into());
    let base_dir = output
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .unwrap_or_default();
    base_dir.join(stem).join(file_name)
}

fn finish_time_split_part(
    part_output: &Path,
    wav_output: Option<PathBuf>,
    options: &AudiobookCommonOptions,
    is_aac: bool,
    is_mp3: bool,
) -> Result<(), String> {
    if !is_aac && !is_mp3 {
        return Ok(());
    }
    let Some(wav_path) = wav_output else {
        return Ok(());
    };
    let res = if is_aac {
        let settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut progress = |_p: u32| {};
        crate::ffmpeg_export::convert_audio_file(
            &wav_path,
            part_output,
            &settings,
            None,
            Some(&mut progress),
        )
    } else {
        let settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut progress = |_p: u32| {};
        crate::ffmpeg_export::convert_audio_file(
            &wav_path,
            part_output,
            &settings,
            None,
            Some(&mut progress),
        )
    };
    if let Err(e) = std::fs::remove_file(&wav_path) {
        crate::log_debug(&format!(
            "Failed to remove temp wav after time split: {}",
            e
        ));
    }
    if let Err(e) = res {
        return Err(i18n::tr_f(
            options.language,
            "sapi5.mf_error",
            &[("err", &e)],
        ));
    }
    Ok(())
}

fn filter_time_split_chunks(chunks: &[String], engine: TtsEngine) -> Vec<String> {
    let mut out = Vec::new();
    let edge_max_bytes = EDGE_TTS_MAX_BYTES.saturating_sub(512).min(3000);
    for (idx, chunk) in chunks.iter().enumerate() {
        let trimmed = chunk.trim();
        if trimmed.is_empty() {
            crate::log_debug(&format!("Time-split: skipping empty chunk index={}", idx));
            continue;
        }
        let len_utf16 = utf16_len(chunk);
        let len_trim_utf16 = utf16_len(trimmed);
        let byte_len = chunk.len();
        crate::log_debug(&format!(
            "Time-split: chunk index={} utf16_len={} trimmed_utf16_len={} bytes_len={}",
            idx, len_utf16, len_trim_utf16, byte_len
        ));
        if engine == TtsEngine::Edge && byte_len > edge_max_bytes {
            crate::log_debug(&format!(
                "Time-split: chunk index={} exceeds edge byte limit ({} > {}), re-splitting",
                idx, byte_len, edge_max_bytes
            ));
            for (sub_idx, sub) in split_long_sentence_edge_with_limit(chunk, edge_max_bytes)
                .into_iter()
                .enumerate()
            {
                if sub.trim().is_empty() {
                    crate::log_debug(&format!(
                        "Time-split: skipping empty sub-chunk parent_index={} sub_index={}",
                        idx, sub_idx
                    ));
                    continue;
                }
                crate::log_debug(&format!(
                    "Time-split: sub-chunk parent_index={} sub_index={} bytes_len={}",
                    idx,
                    sub_idx,
                    sub.len()
                ));
                out.push(sub);
            }
            continue;
        }
        if len_utf16 > MAX_TTS_TEXT_LEN {
            crate::log_debug(&format!(
                "Time-split: chunk index={} exceeds max len ({} > {}), re-splitting",
                idx, len_utf16, MAX_TTS_TEXT_LEN
            ));
            for (sub_idx, sub) in split_text_for_engine(chunk, engine).into_iter().enumerate() {
                if sub.trim().is_empty() {
                    crate::log_debug(&format!(
                        "Time-split: skipping empty sub-chunk parent_index={} sub_index={}",
                        idx, sub_idx
                    ));
                    continue;
                }
                crate::log_debug(&format!(
                    "Time-split: sub-chunk parent_index={} sub_index={} bytes_len={}",
                    idx,
                    sub_idx,
                    sub.len()
                ));
                out.push(sub);
            }
            continue;
        }
        out.push(chunk.clone());
    }
    out
}

fn run_split_audiobook_by_time_edge(
    chunks: &[String],
    split_minutes: u32,
    start_number: u32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }
    let filtered_chunks = filter_time_split_chunks(chunks, TtsEngine::Edge);
    if filtered_chunks.is_empty() {
        return Err("No readable text after time-split filtering".to_string());
    }
    let split_seconds = split_minutes.saturating_mul(60) as f64;
    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
    let is_mp3 = extension == "mp3";

    let edge_chunks: Vec<TtsChunk> = filtered_chunks
        .iter()
        .map(|chunk| TtsChunk {
            text_to_read: chunk.clone(),
            original_len: utf16_len(chunk),
            override_voice: None,
        })
        .collect();

    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|err| err.to_string())?;

    rt.block_on(async {
        let stream_options = EdgeStreamOptions {
            voice: options.voice,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            language: options.language,
            cancel: options.cancel.as_ref(),
            progress_hwnd: options.progress_hwnd,
            allow_http_fallback: true,
        };

        const MAX_PARALLEL_CHUNKS: usize = 8;
        const CHUNK_RETRIES: usize = 6;

        let mut pending: BTreeMap<usize, Vec<u8>> = BTreeMap::new();
        let mut next_to_write = 0usize;
        let mut part_index: u32 = 0;
        let mut current_duration = 0.0f64;
        let mut part_writer: Option<BufWriter<std::fs::File>> = None;
        let mut wav_writer: Option<WavWriter> = None;
        let mut wav_path: Option<PathBuf> = None;
        let mut wav_rate: Option<u32> = None;
        let mut wav_channels: Option<u16> = None;
        let mut current_global_progress: usize = 0;

        let stream_options_ref = &stream_options;
        let mut stream = futures_util::stream::iter(edge_chunks.iter().enumerate())
            .map(|(idx, chunk)| {
                let opts = stream_options_ref;
                async move {
                    let res =
                        download_edge_chunk_ws_with_retry(chunk, opts, idx, CHUNK_RETRIES).await;
                    (idx, res)
                }
            })
            .buffer_unordered(MAX_PARALLEL_CHUNKS);

        while let Some((idx, res)) = stream.next().await {
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }
            let audio = res?;
            pending.insert(idx, audio);

        while let Some(data) = pending.remove(&next_to_write) {
            if data.is_empty() {
                crate::log_debug(&format!(
                    "Edge WS: skipping empty audio payload after sanitization (chunk_index={}).",
                    next_to_write
                ));
                current_global_progress = current_global_progress.saturating_add(1);
                if options.progress_hwnd.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            options.progress_hwnd,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
                next_to_write = next_to_write.saturating_add(1);
                continue;
            }
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }

                let (samples, src_rate, src_channels) = match decode_mp3_to_pcm(&data) {
                    Ok(v) => v,
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Edge WS: skipping undecodable audio chunk (chunk_index={}): {}",
                            next_to_write, err
                        ));
                        current_global_progress = current_global_progress.saturating_add(1);
                        if options.progress_hwnd.0 != 0 {
                            unsafe {
                                if let Err(e) = PostMessageW(
                                    options.progress_hwnd,
                                    crate::WM_UPDATE_PROGRESS,
                                    WPARAM(current_global_progress),
                                    LPARAM(0),
                                ) {
                                    crate::log_debug(&format!(
                                        "Failed to post WM_UPDATE_PROGRESS: {}",
                                        e
                                    ));
                                }
                            }
                        }
                        next_to_write = next_to_write.saturating_add(1);
                        continue;
                    }
                };
                let chunk_duration = samples.len() as f64 / (src_rate as f64 * src_channels as f64);

                if split_seconds > 0.0
                    && current_duration > 0.0
                    && current_duration + chunk_duration > split_seconds
                {
                    if let Some(mut writer) = part_writer.take()
                        && let Err(e) = writer.flush()
                    {
                        crate::log_debug(&format!("Failed to flush time-split part: {}", e));
                    }
                    if let Some(mut w) = wav_writer.take()
                        && let Err(e) = w.finalize()
                    {
                        crate::log_debug(&format!("Failed to finalize time-split wav: {}", e));
                    }
                    let part_output = time_split_part_output(
                        options.output,
                        part_index,
                        start_number,
                        options.part_naming_mode,
                    );
                    finish_time_split_part(
                        &part_output,
                        wav_path.take(),
                        &options,
                        is_aac,
                        is_mp3,
                    )?;
                    part_index = part_index.saturating_add(1);
                    current_duration = 0.0;
                    wav_rate = None;
                    wav_channels = None;
                }

                if part_writer.is_none() && wav_writer.is_none() {
                    let part_output = time_split_part_output(
                        options.output,
                        part_index,
                        start_number,
                        options.part_naming_mode,
                    );
                    if is_aac || is_mp3 {
                        let tmp_wav = part_output.with_extension("wav.tmp");
                        wav_path = Some(tmp_wav.clone());
                        wav_rate = Some(src_rate);
                        wav_channels = Some(src_channels);
                        let writer = WavWriter::create(&tmp_wav, src_rate, src_channels, 16)
                            .map_err(|e| e.to_string())?;
                        wav_writer = Some(writer);
                    } else {
                        let file =
                            std::fs::File::create(&part_output).map_err(|e| e.to_string())?;
                        part_writer = Some(BufWriter::new(file));
                    }
                }

                if let Some(ref mut w) = wav_writer {
                    let target_rate = wav_rate.unwrap_or(src_rate);
                    let target_channels = wav_channels.unwrap_or(src_channels);
                    let out_samples = if src_rate != target_rate || src_channels != target_channels
                    {
                        resample_pcm(
                            &samples,
                            src_rate,
                            src_channels,
                            target_rate,
                            target_channels,
                        )
                    } else {
                        samples
                    };
                    w.write_samples_f32(&out_samples)
                        .map_err(|e| e.to_string())?;
                } else if let Some(ref mut writer) = part_writer {
                    writer.write_all(&data).map_err(|e| e.to_string())?;
                }

                current_duration += chunk_duration;
                current_global_progress = current_global_progress.saturating_add(1);
                if options.progress_hwnd.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            options.progress_hwnd,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
                next_to_write = next_to_write.saturating_add(1);
            }
        }

        if let Some(mut writer) = part_writer.take()
            && let Err(e) = writer.flush()
        {
            crate::log_debug(&format!("Failed to flush time-split part: {}", e));
        }
        if let Some(mut w) = wav_writer.take()
            && let Err(e) = w.finalize()
        {
            crate::log_debug(&format!("Failed to finalize time-split wav: {}", e));
        }
        if next_to_write > 0 {
            let part_output = time_split_part_output(
                options.output,
                part_index,
                start_number,
                options.part_naming_mode,
            );
            finish_time_split_part(&part_output, wav_path.take(), &options, is_aac, is_mp3)?;
        }
        Ok(())
    })
}

fn run_split_audiobook_by_time_sapi(
    engine: TtsEngine,
    chunks: &[String],
    split_minutes: u32,
    start_number: u32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }
    let filtered_chunks = filter_time_split_chunks(chunks, engine);
    if filtered_chunks.is_empty() {
        return Err("No readable text after time-split filtering".to_string());
    }
    let split_seconds = split_minutes.saturating_mul(60) as f64;
    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
    let is_mp3 = extension == "mp3";

    let rt = tokio::runtime::Runtime::new().map_err(|e| e.to_string())?;
    let mut current_duration = 0.0f64;
    let mut part_index: u32 = 0;
    let mut wav_writer: Option<WavWriter> = None;
    let mut wav_path: Option<PathBuf> = None;
    let mut wav_rate: Option<u32> = None;
    let mut wav_channels: Option<u16> = None;
    let mut current_global_progress: usize = 0;

    for chunk in filtered_chunks.iter() {
        if options.cancel.load(Ordering::Relaxed) {
            return Err(cancelled_message(options.language));
        }
        let bytes = rt.block_on(synthesize_segment_bytes(
            engine,
            chunk,
            options.voice,
            options.rate,
            options.pitch,
            options.volume,
            options.language,
        ))?;
        let (samples, src_rate, src_channels) = decode_wav_to_pcm(&bytes)?;
        let chunk_duration = samples.len() as f64 / (src_rate as f64 * src_channels as f64);

        if split_seconds > 0.0
            && current_duration > 0.0
            && current_duration + chunk_duration > split_seconds
        {
            if let Some(mut w) = wav_writer.take()
                && let Err(e) = w.finalize()
            {
                crate::log_debug(&format!("Failed to finalize time-split wav: {}", e));
            }
            let part_output = time_split_part_output(
                options.output,
                part_index,
                start_number,
                options.part_naming_mode,
            );
            finish_time_split_part(&part_output, wav_path.take(), &options, is_aac, is_mp3)?;
            part_index = part_index.saturating_add(1);
            current_duration = 0.0;
            wav_rate = None;
            wav_channels = None;
        }

        if wav_writer.is_none() {
            let part_output = time_split_part_output(
                options.output,
                part_index,
                start_number,
                options.part_naming_mode,
            );
            if is_aac || is_mp3 {
                let tmp_wav = part_output.with_extension("wav.tmp");
                wav_path = Some(tmp_wav.clone());
                wav_rate = Some(src_rate);
                wav_channels = Some(src_channels);
                let writer = WavWriter::create(&tmp_wav, src_rate, src_channels, 16)
                    .map_err(|e| e.to_string())?;
                wav_writer = Some(writer);
            } else {
                wav_rate = Some(src_rate);
                wav_channels = Some(src_channels);
                let writer = WavWriter::create(&part_output, src_rate, src_channels, 16)
                    .map_err(|e| e.to_string())?;
                wav_writer = Some(writer);
            }
        }

        if let Some(ref mut w) = wav_writer {
            let target_rate = wav_rate.unwrap_or(src_rate);
            let target_channels = wav_channels.unwrap_or(src_channels);
            let out_samples = if src_rate != target_rate || src_channels != target_channels {
                resample_pcm(
                    &samples,
                    src_rate,
                    src_channels,
                    target_rate,
                    target_channels,
                )
            } else {
                samples
            };
            w.write_samples_f32(&out_samples)
                .map_err(|e| e.to_string())?;
        }

        current_duration += chunk_duration;
        current_global_progress = current_global_progress.saturating_add(1);
        if options.progress_hwnd.0 != 0 {
            unsafe {
                if let Err(e) = PostMessageW(
                    options.progress_hwnd,
                    crate::WM_UPDATE_PROGRESS,
                    WPARAM(current_global_progress),
                    LPARAM(0),
                ) {
                    crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                }
            }
        }
    }

    if let Some(mut w) = wav_writer.take()
        && let Err(e) = w.finalize()
    {
        crate::log_debug(&format!("Failed to finalize time-split wav: {}", e));
    }
    if !chunks.is_empty() {
        let part_output = time_split_part_output(
            options.output,
            part_index,
            start_number,
            options.part_naming_mode,
        );
        finish_time_split_part(&part_output, wav_path.take(), &options, is_aac, is_mp3)?;
    }

    Ok(())
}

async fn download_edge_chunks_ws_parallel_to_writer(
    chunks: &[TtsChunk],
    start_index: usize,
    options: &EdgeStreamOptions<'_>,
    writer: &mut dyn std::io::Write,
    current_global_progress: &mut usize,
) -> Result<usize, String> {
    const MAX_PARALLEL_CHUNKS: usize = 8;
    const CHUNK_RETRIES: usize = 6;
    let is_lithuanian_voice = options.voice.to_ascii_lowercase().starts_with("lt-");

    if start_index >= chunks.len() {
        return Ok(chunks.len());
    }

    let mut pending: BTreeMap<usize, Vec<u8>> = BTreeMap::new();
    let mut next_to_write = start_index;
    let mut written_chunks = 0usize;
    let mut skipped_empty_chunks = 0usize;
    let mut total_audio_bytes = 0usize;

    let mut stream = futures_util::stream::iter(chunks.iter().enumerate().skip(start_index))
        .map(|(idx, chunk)| async move {
            let mut validate_attempt = 0usize;
            loop {
                if options.cancel.load(Ordering::Relaxed) {
                    break (idx, Err("Cancelled".to_string()));
                }

                validate_attempt = validate_attempt.saturating_add(1);
                let res = if is_lithuanian_voice {
                    match download_edge_chunk_ws_adaptive_lt(chunk, options, idx, CHUNK_RETRIES, 0)
                        .await
                    {
                        Ok(mut audio) => {
                            if is_lt_audio_suspicious(&chunk.text_to_read, audio.len())
                                && let Ok(strict_audio) = download_edge_chunk_ws_strict_small_lt(
                                    chunk,
                                    options,
                                    idx,
                                    CHUNK_RETRIES,
                                )
                                .await
                                && strict_audio.len() > audio.len()
                            {
                                audio = strict_audio;
                            }
                            Ok(audio)
                        }
                        Err(err) => Err(err),
                    }
                } else {
                    download_edge_chunk_ws_with_retry(chunk, options, idx, CHUNK_RETRIES).await
                };

                let audio = match res {
                    Ok(audio) => audio,
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Edge WS: chunk {} fetch failed on validation attempt {}: {}. Retrying...",
                            idx, validate_attempt, err
                        ));
                        continue;
                    }
                };

                match decode_mp3_to_pcm(&audio) {
                    Ok((samples, _rate, _channels)) if !samples.is_empty() => {
                        break (idx, Ok(audio));
                    }
                    Ok((_samples, _rate, _channels)) => {
                        crate::log_debug(&format!(
                            "Edge WS: chunk {} invalid audio (empty decoded samples) on validation attempt {}. Retrying...",
                            idx, validate_attempt
                        ));
                    }
                    Err(err) => {
                        crate::log_debug(&format!(
                            "Edge WS: chunk {} invalid audio (decode error) on validation attempt {}: {}. Retrying...",
                            idx, validate_attempt, err
                        ));
                    }
                }
            }
        })
        .buffer_unordered(MAX_PARALLEL_CHUNKS);

    while let Some((idx, res)) = stream.next().await {
        if options.cancel.load(Ordering::Relaxed) {
            return Err("Cancelled".to_string());
        }
        let audio = res?;
        pending.insert(idx, audio);
        while let Some(data) = pending.remove(&next_to_write) {
            if data.is_empty() {
                skipped_empty_chunks = skipped_empty_chunks.saturating_add(1);
                *current_global_progress += 1;
                if options.progress_hwnd.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            options.progress_hwnd,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(*current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
                next_to_write = next_to_write.saturating_add(1);
                continue;
            }
            writer.write_all(&data).map_err(|err| err.to_string())?;
            writer.flush().map_err(|err| err.to_string())?;
            written_chunks = written_chunks.saturating_add(1);
            total_audio_bytes = total_audio_bytes.saturating_add(data.len());

            *current_global_progress += 1;
            if options.progress_hwnd.0 != 0 {
                unsafe {
                    if let Err(e) = PostMessageW(
                        options.progress_hwnd,
                        crate::WM_UPDATE_PROGRESS,
                        WPARAM(*current_global_progress),
                        LPARAM(0),
                    ) {
                        crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                    }
                }
            }
            next_to_write = next_to_write.saturating_add(1);
        }
    }

    let expected_chunks = chunks.len().saturating_sub(start_index);
    crate::log_debug(&format!(
        "Edge WS export summary: expected_chunks={} written_chunks={} skipped_empty_chunks={} total_audio_bytes={}",
        expected_chunks, written_chunks, skipped_empty_chunks, total_audio_bytes
    ));

    Ok(chunks.len())
}

fn get_text_from_caret(hwnd_edit: HWND) -> (String, i32) {
    unsafe {
        let mut range = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut range as *mut _ as isize),
        );
        let caret_pos = range.cpMin.min(range.cpMax).max(0);
        let total_len = SendMessageW(hwnd_edit, WM_GETTEXTLENGTH, WPARAM(0), LPARAM(0)).0 as i32;
        if total_len <= 0 {
            return (String::new(), 0);
        }
        let full_text = get_text_range(hwnd_edit, 0, total_len);

        // Se siamo all'inizio, o se siamo alla fine del testo, leggi tutto dall'inizio
        if caret_pos == 0 {
            return (full_text, 0);
        }

        // Se la posizione del cursore Š oltre la lunghezza del testo (fine file),
        // ricomincia a leggere dall'inizio come richiesto.
        if caret_pos >= total_len {
            return (full_text, 0);
        }

        let prefix = get_text_range(hwnd_edit, 0, caret_pos);
        let caret_utf16 = prefix.encode_utf16().count() as i32;

        let wide: Vec<u16> = full_text.encode_utf16().collect();

        let adjusted_pos = adjust_tts_caret_pos(&full_text, caret_utf16);
        let adjusted_pos = adjusted_pos.max(0) as usize;
        if adjusted_pos >= wide.len() {
            return (full_text, 0);
        }
        (
            String::from_utf16_lossy(&wide[adjusted_pos..]),
            adjusted_pos as i32,
        )
    }
}

fn get_text_range(hwnd_edit: HWND, start: i32, end: i32) -> String {
    unsafe {
        let len = (end - start).max(0) as usize;
        if len == 0 {
            return String::new();
        }
        let mut buf = vec![0u16; len + 1];
        let mut text_range = TEXTRANGEW {
            chrg: CHARRANGE {
                cpMin: start,
                cpMax: end,
            },
            lpstrText: PWSTR(buf.as_mut_ptr()),
        };
        let copied = SendMessageW(
            hwnd_edit,
            EM_GETTEXTRANGE,
            WPARAM(0),
            LPARAM(&mut text_range as *mut _ as isize),
        )
        .0 as usize;
        let used = copied.min(len);
        String::from_utf16_lossy(&buf[..used])
    }
}

fn adjust_tts_caret_pos(text: &str, pos: i32) -> i32 {
    if pos <= 0 {
        return 0;
    }
    let mut items: Vec<(usize, usize, bool)> = Vec::new();
    let mut offset = 0usize;
    for ch in text.chars() {
        let start = offset;
        let len = ch.len_utf16();
        let end = start + len;
        let is_word = ch.is_alphanumeric() || ch == '_';
        items.push((start, end, is_word));
        offset = end;
    }
    if offset == 0 {
        return pos;
    }
    let mut pos_usize = pos as usize;
    if pos_usize > offset {
        pos_usize = offset;
    }

    let mut prev: Option<usize> = None;
    let mut next: Option<usize> = None;
    for (idx, (start, end, _)) in items.iter().enumerate() {
        if *end <= pos_usize {
            prev = Some(idx);
            continue;
        }
        if *start >= pos_usize {
            next = Some(idx);
            break;
        }
        next = Some(idx);
        break;
    }

    let prev_is_word = prev
        .and_then(|idx| items.get(idx))
        .map(|v| v.2)
        .unwrap_or(false);
    let next_is_word = next
        .and_then(|idx| items.get(idx))
        .map(|v| v.2)
        .unwrap_or(false);
    if prev_is_word
        && next_is_word
        && let Some(mut idx) = prev
    {
        while idx > 0 && items[idx - 1].2 {
            idx -= 1;
        }
        return items[idx].0 as i32;
    }
    pos
}
fn generate_sec_ms_gec() -> String {
    let win_epoch = 11644473600i64;
    let ticks = Local::now().timestamp() + win_epoch;
    let ticks = (ticks - (ticks % 300)) * 10_000_000;
    let str_to_hash = format!("{}{}", ticks, TRUSTED_CLIENT_TOKEN);
    let mut hasher = Sha256::new();
    hasher.update(str_to_hash);
    hex::encode(hasher.finalize()).to_uppercase()
}

fn generate_muid() -> String {
    let mut rng = rand::thread_rng();
    let mut bytes = [0u8; 16];
    rng.fill(&mut bytes);
    hex::encode(bytes).to_uppercase()
}

fn get_date_string() -> String {
    Local::now()
        .format("%a %b %d %Y %H:%M:%S GMT+0000 (Coordinated Universal Time)")
        .to_string()
}

fn format_rate(rate: i32) -> String {
    format!("{:+}%", rate)
}

fn format_pitch(pitch: i32) -> String {
    format!("{:+}Hz", pitch)
}

fn format_volume(volume: i32) -> String {
    let delta = volume.saturating_sub(100);
    format!("{:+}%", delta)
}

fn mkssml(text: &str, voice: &str, tts_rate: i32, tts_pitch: i32, tts_volume: i32) -> String {
    let lang = voice.split('-').collect::<Vec<_>>();
    let lang = if lang.len() >= 3 {
        lang[0..2].join("-")
    } else {
        "en-US".to_string()
    };
    let rendered_text = render_edge_ssml_text_with_pause_tags(text);
    format!(
        "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='{}'><voice name='{}'><prosody pitch='{}' rate='{}' volume='{}'>{}</prosody></voice></speak>",
        lang,
        voice,
        format_pitch(tts_pitch),
        format_rate(tts_rate),
        format_volume(tts_volume),
        rendered_text
    )
}

fn parse_pause_tag_milliseconds(tag: &str) -> Option<u32> {
    let trimmed = tag.trim();
    let inner = trimmed
        .strip_prefix('<')?
        .strip_suffix('>')?
        .trim()
        .trim_end_matches('/')
        .trim();
    let rest = inner.strip_prefix("pause")?.trim();
    if rest.is_empty() {
        return None;
    }
    for token in rest.split_whitespace() {
        let value = token
            .strip_prefix("ms=")
            .or_else(|| token.strip_prefix("milliseconds="))
            .unwrap_or(token)
            .trim_matches(['"', '\'']);
        if let Ok(ms) = value.parse::<u32>()
            && (PAUSE_TAG_MIN_MS..=PAUSE_TAG_MAX_MS).contains(&ms)
        {
            return Some(ms);
        }
    }
    None
}

fn render_ssml_text_with_pause_tags(text: &str, pause_markup: impl Fn(u32) -> String) -> String {
    let lower = text.to_ascii_lowercase();
    let mut out = String::with_capacity(text.len());
    let mut cursor = 0usize;
    let mut i = 0usize;
    let bytes = lower.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'<' {
            let remaining = &lower[i..];
            if remaining.starts_with("<pause")
                && let Some(end_rel) = remaining.find('>')
            {
                let end = i + end_rel + 1;
                let tag = &text[i..end];
                if let Some(ms) = parse_pause_tag_milliseconds(&lower[i..end]) {
                    if i > cursor {
                        out.push_str(&escape_xml(&text[cursor..i]));
                    }
                    out.push_str(&pause_markup(ms));
                    cursor = end;
                    i = end;
                    continue;
                }
                if i > cursor {
                    out.push_str(&escape_xml(&text[cursor..i]));
                }
                out.push_str(&escape_xml(tag));
                cursor = end;
                i = end;
                continue;
            }
        }
        i += 1;
    }
    if cursor < text.len() {
        out.push_str(&escape_xml(&text[cursor..]));
    }
    out
}

pub(crate) fn render_edge_ssml_text_with_pause_tags(text: &str) -> String {
    let lower = text.to_ascii_lowercase();
    let mut out = String::with_capacity(text.len());
    let mut cursor = 0usize;
    let mut i = 0usize;
    let bytes = lower.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'<' {
            let remaining = &lower[i..];
            if remaining.starts_with("<pause")
                && let Some(end_rel) = remaining.find('>')
            {
                let end = i + end_rel + 1;
                if let Some(ms) = parse_pause_tag_milliseconds(&lower[i..end]) {
                    out.push_str(&text[cursor..i]);
                    out.push_str(&format!(
                        "<mstts:silence type=\"Sentenceboundary\" value=\"{ms}ms\"/>"
                    ));
                    cursor = end;
                    i = end;
                    continue;
                }
            }
        } else if bytes[i] == b'&' {
            let remaining = &lower[i..];
            if remaining.starts_with("&lt;pause")
                && let Some(end_rel) = remaining.find("&gt;")
            {
                let end = i + end_rel + "&gt;".len();
                let decoded = decode_basic_xml_entities(&lower[i..end]);
                if let Some(ms) = parse_pause_tag_milliseconds(&decoded) {
                    out.push_str(&text[cursor..i]);
                    out.push_str(&format!(
                        "<mstts:silence type=\"Sentenceboundary\" value=\"{ms}ms\"/>"
                    ));
                    cursor = end;
                    i = end;
                    continue;
                }
            }
        }
        i += 1;
    }
    if cursor < text.len() {
        out.push_str(&text[cursor..]);
    }
    out
}

pub(crate) fn render_sapi_ssml_text_with_pause_tags(text: &str) -> String {
    render_ssml_text_with_pause_tags(text, |ms| format!("<silence msec=\"{ms}\"/>"))
}

pub fn remove_long_dash_runs(line: &str) -> String {
    let mut out = String::with_capacity(line.len());
    let mut dash_run = 0;
    for ch in line.chars() {
        if ch == '-' {
            dash_run += 1;
            continue;
        }
        if dash_run > 0 {
            if dash_run < 3 {
                out.extend(std::iter::repeat_n('-', dash_run));
            }
            dash_run = 0;
        }
        out.push(ch);
    }
    if dash_run > 0 && dash_run < 3 {
        out.extend(std::iter::repeat_n('-', dash_run));
    }
    out
}

pub fn strip_dashed_lines(text: &str) -> String {
    text.lines()
        .filter_map(|line| {
            let trimmed = line.trim();
            if trimmed.is_empty() {
                return Some(String::new());
            }
            let cleaned = remove_long_dash_runs(line);
            if cleaned.trim().is_empty() {
                None
            } else {
                Some(cleaned)
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn normalize_weird_spaces(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            // Remove zero-width formatting chars that can confuse TTS segmentation.
            '\u{200B}' | '\u{200C}' | '\u{200D}' | '\u{2060}' | '\u{FEFF}' => {}
            // Convert uncommon Unicode spaces to ASCII space.
            '\u{00A0}' | '\u{1680}' | '\u{2000}' | '\u{2001}' | '\u{2002}' | '\u{2003}'
            | '\u{2004}' | '\u{2005}' | '\u{2006}' | '\u{2007}' | '\u{2008}' | '\u{2009}'
            | '\u{200A}' | '\u{202F}' | '\u{205F}' | '\u{3000}'
            // Visible symbols for spaces/tabs used in some copied texts.
            | '\u{2420}' | '\u{2423}' | '\u{2409}' => out.push(' '),
            // Visible symbols for LF/CR/newline.
            '\u{240A}' | '\u{240D}' | '\u{2424}' => out.push('\n'),
            _ => out.push(ch),
        }
    }
    out
}

#[derive(Clone, Copy)]
enum TtsSanitizeProfile {
    Safe,
    Strict,
}

fn normalize_tts_strict_chars(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        let mapped = match ch {
            // Smart/directional double quotes -> plain quote
            '“' | '”' | '„' | '‟' | '«' | '»' | '〝' | '〞' => '"',
            // Smart/directional single quotes/apostrophes -> plain apostrophe
            '‘' | '’' | '‚' | '‛' | '‹' | '›' | '`' | '´' => '\'',
            // Long dashes/minus variants -> hyphen-minus
            '–' | '—' | '―' | '−' => '-',
            _ => ch,
        };

        // Horizontal ellipsis -> three dots
        if mapped == '…' {
            out.push_str("...");
            continue;
        }

        // Keep only letters/digits/whitespace and a conservative punctuation set.
        let safe_punct = matches!(
            mapped,
            '.' | ','
                | ';'
                | ':'
                | '!'
                | '?'
                | '"'
                | '\''
                | '-'
                | '('
                | ')'
                | '['
                | ']'
                | '{'
                | '}'
                | '/'
                | '\\'
                | '@'
                | '#'
                | '%'
                | '&'
                | '+'
                | '='
                | '*'
                | '_'
                | '<'
                | '>'
                | '|'
                | '~'
        );

        if mapped.is_alphanumeric() || mapped.is_whitespace() || safe_punct {
            out.push(mapped);
        } else {
            out.push(' ');
        }
    }
    out
}

fn preview_for_log(text: &str, max_chars: usize) -> String {
    let mut out = String::with_capacity(max_chars);
    for ch in text.chars().take(max_chars) {
        if ch == '\n' || ch == '\r' || ch == '\t' {
            out.push(' ');
        } else {
            out.push(ch);
        }
    }
    out
}

fn normalize_for_tts_with_profile(
    text: &str,
    split_on_newline: bool,
    profile: TtsSanitizeProfile,
) -> String {
    let normalized_spaces = normalize_weird_spaces(text);
    let mut normalized = if split_on_newline {
        normalized_spaces
    } else {
        normalized_spaces.replace('\n', " ").replace('\r', "")
    };
    match profile {
        TtsSanitizeProfile::Safe => normalized.replace(['«', '»'], ""),
        TtsSanitizeProfile::Strict => {
            normalized = normalize_tts_strict_chars(&normalized);
            normalized
        }
    }
}

pub fn normalize_for_tts(text: &str, split_on_newline: bool) -> String {
    normalize_for_tts_with_profile(text, split_on_newline, TtsSanitizeProfile::Safe)
}

fn apply_dictionary(text: &str, dictionary: &[DictionaryEntry]) -> String {
    if dictionary.is_empty() {
        return text.to_string();
    }
    let mut out = text.to_string();
    for entry in dictionary {
        if entry.original.is_empty() {
            continue;
        }
        out = if entry.match_case {
            out.replace(&entry.original, &entry.replacement)
        } else {
            replace_case_insensitive(&out, &entry.original, &entry.replacement)
        };
    }
    out
}

fn replace_case_insensitive(text: &str, needle: &str, replacement: &str) -> String {
    if needle.is_empty() {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut cursor = 0;
    while cursor < text.len() {
        if let Some(match_len) = match_len_case_insensitive(text, cursor, needle) {
            out.push_str(replacement);
            cursor += match_len;
            continue;
        }
        let Some(ch) = text[cursor..].chars().next() else {
            break;
        };
        out.push(ch);
        cursor += ch.len_utf8();
    }
    out
}

fn match_len_case_insensitive(text: &str, start: usize, needle: &str) -> Option<usize> {
    let mut consumed = 0;
    let mut text_chars = text[start..].chars();
    for needle_ch in needle.chars() {
        let text_ch = text_chars.next()?;
        if text_ch.to_lowercase().to_string() != needle_ch.to_lowercase().to_string() {
            return None;
        }
        consumed += text_ch.len_utf8();
    }
    Some(consumed)
}

pub(crate) fn prepare_tts_text(
    text: &str,
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
) -> String {
    let normalized = normalize_for_tts(text, split_on_newline);
    apply_dictionary(&normalized, dictionary)
}

fn normalize_newlines(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}

pub(crate) struct MarkerEntry {
    pub(crate) pos: usize,
    pub(crate) label: String,
}

fn marker_label_for_position(text: &str, pos: usize, marker: &str) -> String {
    let start = text[..pos].rfind('\n').map(|idx| idx + 1).unwrap_or(0);
    let end = text[pos..]
        .find('\n')
        .map(|idx| pos + idx)
        .unwrap_or(text.len());
    let line = text[start..end].trim();
    if line.is_empty() {
        marker.to_string()
    } else {
        line.to_string()
    }
}

pub(crate) fn collect_marker_entries(
    text: &str,
    marker: &str,
    require_newline: bool,
) -> (String, Vec<MarkerEntry>) {
    let normalized = normalize_newlines(text);
    if marker.trim().is_empty() {
        return (normalized, Vec::new());
    }

    let mut entries = Vec::new();
    for (idx, _) in normalized.match_indices(marker) {
        // If require_newline is true, we ensure the marker is at the start of a line.
        // normalized only has \n (no \r), but we check both just in case or for robustness.
        // We also allow the start of the file (idx == 0).
        if require_newline && idx > 0 {
            let prefix = &normalized[..idx];
            if !prefix.ends_with('\n') {
                continue;
            }
        }
        let label = marker_label_for_position(&normalized, idx, marker);
        entries.push(MarkerEntry { pos: idx, label });
    }

    (normalized, entries)
}

fn split_text_by_positions(text: &str, positions: &[usize]) -> Option<Vec<String>> {
    if positions.is_empty() {
        return None;
    }

    let mut positions = positions.to_vec();
    positions.sort_unstable();
    positions.dedup();

    let mut parts = Vec::new();
    let mut start = 0usize;
    for pos in positions.iter() {
        if *pos == 0 {
            continue;
        }
        // Ensure we don't go out of bounds and that we advance
        if *pos > start && *pos <= text.len() {
            parts.push(text[start..*pos].to_string());
            start = *pos;
        }
    }
    // Push the remainder of the text
    if start < text.len() {
        parts.push(text[start..].to_string());
    } else if start == text.len() && parts.is_empty() {
        // Edge case: empty text or positions at end?
        // If text is not empty, but start reached end, we are done.
    }
    Some(parts)
}

pub(crate) fn build_audiobook_parts_from_sections(
    sections: &[String],
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
    engine: TtsEngine,
) -> Option<Vec<Vec<String>>> {
    let mut parts = Vec::new();
    for (idx, section) in sections.iter().enumerate() {
        let cleaned = strip_dashed_lines(section);
        let prepared = prepare_tts_text(&cleaned, split_on_newline, dictionary);
        let chunks = split_text_for_engine(&prepared, engine);
        if chunks.is_empty() {
            crate::log_debug(&format!("EPUB split: skipping empty chapter index={}", idx));
            continue;
        }
        parts.push(chunks);
    }
    if parts.is_empty() { None } else { Some(parts) }
}

pub(crate) fn build_audiobook_parts_by_positions(
    text: &str,
    positions: &[usize],
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
    engine: TtsEngine,
) -> Option<Vec<Vec<String>>> {
    let parts_text = split_text_by_positions(text, positions)?;
    let mut parts_chunks = Vec::new();

    for part_text in parts_text {
        let prepared = prepare_tts_text(&part_text, split_on_newline, dictionary);
        let chunks = split_text_for_engine(&prepared, engine);
        // Even if chunks is empty (e.g. only whitespace part), we might want to keep the structure?
        // But run_tts_audiobook_part handles empty chunks by skipping.
        if !chunks.is_empty() {
            parts_chunks.push(chunks);
        } else {
            // If a part is empty, we push an empty vec to maintain index alignment if needed,
            // though currently the caller just iterates.
            parts_chunks.push(Vec::new());
        }
    }

    if parts_chunks.is_empty() {
        None
    } else {
        Some(parts_chunks)
    }
}

pub(crate) fn build_mixed_audiobook_parts_from_sections(
    sections: &[String],
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
    tts_engine: TtsEngine,
) -> Option<Vec<Vec<TtsChunk>>> {
    let mut parts = Vec::new();
    for (idx, section) in sections.iter().enumerate() {
        let cleaned = strip_dashed_lines(section);
        let chunks = split_into_tts_chunks(&cleaned, split_on_newline, dictionary, tts_engine);
        if chunks.is_empty() {
            crate::log_debug(&format!(
                "EPUB split (mixed): skipping empty chapter index={}",
                idx
            ));
            continue;
        }
        parts.push(chunks);
    }
    if parts.is_empty() { None } else { Some(parts) }
}

pub(crate) fn build_mixed_audiobook_parts_by_positions(
    text: &str,
    positions: &[usize],
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
    tts_engine: TtsEngine,
) -> Option<Vec<Vec<TtsChunk>>> {
    let parts_text = split_text_by_positions(text, positions)?;
    let mut parts_chunks = Vec::new();

    for part_text in parts_text {
        let chunks = split_into_tts_chunks(&part_text, split_on_newline, dictionary, tts_engine);
        parts_chunks.push(chunks);
    }

    if parts_chunks.is_empty() {
        None
    } else {
        Some(parts_chunks)
    }
}

pub fn split_text(text: &str) -> Vec<String> {
    let mut chunks: Vec<String> = Vec::new();
    let text = text.trim();
    if text.is_empty() {
        return chunks;
    }

    let char_len = text.chars().count();
    let is_long = char_len > TTS_LONG_TEXT_THRESHOLD;
    let max_len = if is_long {
        MAX_TTS_TEXT_LEN_LONG
    } else {
        MAX_TTS_TEXT_LEN
    };
    let target_chunk_len = if is_long {
        MAX_TTS_FIRST_CHUNK_LEN_LONG
    } else {
        max_len
    };

    let sentences = split_sentences(text);
    let mut current = String::new();
    let mut current_limit = target_chunk_len;

    for sentence in sentences {
        let sentence_trim = sentence.trim();
        if sentence_trim.is_empty() {
            continue;
        }
        let sentence_len = sentence_trim.chars().count();
        if current.is_empty() {
            if sentence_len > current_limit {
                let parts = split_long_sentence_by_whitespace(sentence_trim, current_limit);
                for part in parts {
                    chunks.push(part);
                }
                current_limit = max_len;
                continue;
            }
            current.push_str(sentence_trim);
        } else {
            let combined_len = current
                .chars()
                .count()
                .saturating_add(1)
                .saturating_add(sentence_len);
            if combined_len > current_limit {
                chunks.push(current.trim().to_string());
                current.clear();
                current_limit = max_len;

                if sentence_len > current_limit {
                    let parts = split_long_sentence_by_whitespace(sentence_trim, current_limit);
                    for part in parts {
                        chunks.push(part);
                    }
                } else {
                    current.push_str(sentence_trim);
                }
            } else {
                current.push(' ');
                current.push_str(sentence_trim);
            }
        }
    }

    if !current.trim().is_empty() {
        chunks.push(current.trim().to_string());
    }

    chunks
}

fn sanitize_edge_text(text: &str) -> String {
    let normalized_spaces = normalize_weird_spaces(text);
    let mut out = String::with_capacity(normalized_spaces.len());
    let mut dot_run = 0usize;
    let mut bang_run = 0usize;
    let mut question_run = 0usize;

    for ch in normalized_spaces.chars() {
        if ch == '.' {
            dot_run += 1;
            continue;
        }
        if ch == '!' {
            bang_run += 1;
            continue;
        }
        if ch == '?' {
            question_run += 1;
            continue;
        }

        flush_edge_punctuation_runs(&mut out, &mut dot_run, &mut bang_run, &mut question_run);
        let code = ch as u32;
        if (0..=8).contains(&code) || (11..=12).contains(&code) || (14..=31).contains(&code) {
            out.push(' ');
        } else {
            out.push(ch);
        }
    }
    flush_edge_punctuation_runs(&mut out, &mut dot_run, &mut bang_run, &mut question_run);
    normalize_edge_terminal_punctuation(&out)
}

fn is_edge_text_usable(text: &str) -> bool {
    text.chars().any(|ch| ch.is_alphanumeric())
}

fn normalize_edge_terminal_punctuation(text: &str) -> String {
    // Edge can fail on some terminal combinations coming from headlines, e.g. `...":`.
    // Keep behavior minimal: only normalize dot + quote + colon sequences to a plain sentence end.
    text.replace(".\":", ". ")
        .replace(".':", ". ")
        .replace(".”:", ". ")
        .replace(".’:", ". ")
}

fn flush_edge_punctuation_runs(
    out: &mut String,
    dot_run: &mut usize,
    bang_run: &mut usize,
    question_run: &mut usize,
) {
    let trim_trailing_spaces = |s: &mut String| {
        while s.chars().last().is_some_and(|ch| ch.is_whitespace()) {
            s.pop();
        }
    };
    if *dot_run >= 3 {
        trim_trailing_spaces(out);
        out.push('.');
    } else if *dot_run > 0 {
        trim_trailing_spaces(out);
        out.extend(std::iter::repeat_n('.', *dot_run));
    }
    if *bang_run >= 3 {
        trim_trailing_spaces(out);
        out.push('!');
    } else if *bang_run > 0 {
        trim_trailing_spaces(out);
        out.extend(std::iter::repeat_n('!', *bang_run));
    }
    if *question_run >= 3 {
        trim_trailing_spaces(out);
        out.push('?');
    } else if *question_run > 0 {
        trim_trailing_spaces(out);
        out.extend(std::iter::repeat_n('?', *question_run));
    }
    *dot_run = 0;
    *bang_run = 0;
    *question_run = 0;
}

fn escape_xml(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

fn adjust_split_for_xml_entity(text: &str, mut split_at: usize) -> usize {
    while split_at > 0 {
        let prefix = &text[..split_at];
        let Some(amp_idx) = prefix.rfind('&') else {
            break;
        };
        if prefix[amp_idx..].contains(';') {
            break;
        }
        split_at = amp_idx;
    }
    split_at
}

fn find_last_newline_or_space_within_limit(text: &str, max_bytes: usize) -> usize {
    let mut last_newline = 0usize;
    let mut last_space = 0usize;
    for (idx, ch) in text.char_indices() {
        let end = idx + ch.len_utf8();
        if end > max_bytes {
            break;
        }
        if ch == '\n' {
            last_newline = end;
        } else if ch.is_whitespace() {
            last_space = end;
        }
    }
    if last_newline > 0 {
        last_newline
    } else {
        last_space
    }
}

fn find_safe_utf8_split_idx(text: &str, max_bytes: usize) -> usize {
    let mut last_end = 0usize;
    for (idx, ch) in text.char_indices() {
        let end = idx + ch.len_utf8();
        if end > max_bytes {
            break;
        }
        last_end = end;
    }
    last_end
}

fn find_edge_split_idx(text: &str, max_bytes: usize) -> usize {
    let total_bytes = text.len();
    if total_bytes <= max_bytes {
        return text.len();
    }

    let mut split_at = find_last_newline_or_space_within_limit(text, max_bytes);
    if split_at == 0 {
        split_at = find_safe_utf8_split_idx(text, max_bytes);
    }
    split_at = adjust_split_for_xml_entity(text, split_at);
    if split_at == 0 {
        split_at = find_safe_utf8_split_idx(text, max_bytes);
    }
    split_at
}

fn split_text_edge(text: &str) -> Vec<String> {
    let cleaned = sanitize_edge_text(text);
    let escaped = escape_xml(&cleaned);
    let sentences = split_sentences(&escaped);
    let mut out = Vec::new();
    let mut current = String::new();

    for sentence in sentences {
        let sentence_trim = sentence.trim();
        if sentence_trim.is_empty() {
            continue;
        }

        let sentence_len = sentence_trim.len();
        if current.is_empty() {
            if sentence_len > EDGE_TTS_MAX_BYTES {
                let parts = split_long_sentence_edge(sentence_trim);
                out.extend(parts);
                continue;
            }
            current.push_str(sentence_trim);
        } else {
            let combined_len = current.len().saturating_add(1).saturating_add(sentence_len);
            if combined_len > EDGE_TTS_MAX_BYTES {
                out.push(current.trim().to_string());
                current.clear();

                if sentence_len > EDGE_TTS_MAX_BYTES {
                    let parts = split_long_sentence_edge(sentence_trim);
                    out.extend(parts);
                } else {
                    current.push_str(sentence_trim);
                }
            } else {
                current.push(' ');
                current.push_str(sentence_trim);
            }
        }
    }

    if !current.trim().is_empty() {
        out.push(current.trim().to_string());
    }

    out
}

fn split_text_edge_with_limit(text: &str, max_bytes: usize) -> Vec<String> {
    let cleaned = sanitize_edge_text(text);
    let escaped = escape_xml(&cleaned);
    let sentences = split_sentences(&escaped);
    let mut out = Vec::new();
    let mut current = String::new();

    for sentence in sentences {
        let sentence_trim = sentence.trim();
        if sentence_trim.is_empty() {
            continue;
        }

        let sentence_len = sentence_trim.len();
        if current.is_empty() {
            if sentence_len > max_bytes {
                out.extend(split_long_sentence_edge_with_limit(
                    sentence_trim,
                    max_bytes,
                ));
            } else {
                current.push_str(sentence_trim);
            }
        } else {
            let combined_len = current.len().saturating_add(1).saturating_add(sentence_len);
            if combined_len > max_bytes {
                out.push(current.trim().to_string());
                current.clear();
                if sentence_len > max_bytes {
                    out.extend(split_long_sentence_edge_with_limit(
                        sentence_trim,
                        max_bytes,
                    ));
                } else {
                    current.push_str(sentence_trim);
                }
            } else {
                current.push(' ');
                current.push_str(sentence_trim);
            }
        }
    }

    if !current.trim().is_empty() {
        out.push(current.trim().to_string());
    }

    out
}

fn split_edge_chunk_in_two(text: &str) -> Option<(String, String)> {
    if text.trim().is_empty() || text.len() < 8 {
        return None;
    }
    let mid = text.len() / 2;
    let mut split_at = find_edge_split_idx(text, mid);
    if split_at == 0 || split_at >= text.len() {
        split_at = find_edge_split_idx(text, text.len().saturating_sub(1));
    }
    if split_at == 0 || split_at >= text.len() {
        return None;
    }
    let (a, b) = text.split_at(split_at);
    let a = a.trim().to_string();
    let b = b.trim().to_string();
    if a.is_empty() || b.is_empty() {
        return None;
    }
    Some((a, b))
}

fn is_lt_audio_suspicious(text: &str, audio_len: usize) -> bool {
    let text_len = utf16_len(text);
    if text_len < 120 {
        return false;
    }
    let min_expected = text_len.saturating_mul(120);
    audio_len < min_expected
}

async fn download_edge_chunk_ws_adaptive_lt(
    chunk: &TtsChunk,
    options: &EdgeStreamOptions<'_>,
    idx: usize,
    max_retries: usize,
    depth: usize,
) -> Result<Vec<u8>, String> {
    let audio = download_edge_chunk_ws_with_retry(chunk, options, idx, max_retries).await?;
    if !is_lt_audio_suspicious(&chunk.text_to_read, audio.len()) {
        return Ok(audio);
    }
    if depth >= 3 {
        return Ok(audio);
    }
    let Some((left, right)) = split_edge_chunk_in_two(&chunk.text_to_read) else {
        return Ok(audio);
    };

    let left_chunk = TtsChunk {
        text_to_read: left,
        original_len: chunk.original_len / 2,
        override_voice: chunk.override_voice.clone(),
    };
    let right_chunk = TtsChunk {
        text_to_read: right,
        original_len: chunk.original_len.saturating_sub(left_chunk.original_len),
        override_voice: chunk.override_voice.clone(),
    };

    let left_audio = Box::pin(download_edge_chunk_ws_adaptive_lt(
        &left_chunk,
        options,
        idx,
        max_retries,
        depth + 1,
    ))
    .await?;
    let right_audio = Box::pin(download_edge_chunk_ws_adaptive_lt(
        &right_chunk,
        options,
        idx,
        max_retries,
        depth + 1,
    ))
    .await?;

    let mut merged = Vec::with_capacity(left_audio.len().saturating_add(right_audio.len()));
    merged.extend_from_slice(&left_audio);
    merged.extend_from_slice(&right_audio);
    Ok(merged)
}

async fn download_edge_chunk_ws_strict_small_lt(
    chunk: &TtsChunk,
    options: &EdgeStreamOptions<'_>,
    idx: usize,
    max_retries: usize,
) -> Result<Vec<u8>, String> {
    const STRICT_LIMIT: usize = 320;
    let parts = split_text_edge_with_limit(&chunk.text_to_read, STRICT_LIMIT);
    if parts.len() <= 1 {
        return download_edge_chunk_ws_with_retry(chunk, options, idx, max_retries).await;
    }
    let mut out = Vec::new();
    for part in parts {
        let sub = TtsChunk {
            text_to_read: part,
            original_len: chunk.original_len,
            override_voice: chunk.override_voice.clone(),
        };
        let audio = download_edge_chunk_ws_with_retry(&sub, options, idx, max_retries).await?;
        out.extend_from_slice(&audio);
    }
    Ok(out)
}

fn split_sentences(text: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0usize;
    let chars: Vec<(usize, char)> = text.char_indices().collect();

    for (index, (idx, ch)) in chars.iter().copied().enumerate() {
        if matches!(ch, '.' | '!' | '?' | ';' | ':') {
            let prev = index
                .checked_sub(1)
                .and_then(|i| chars.get(i))
                .map(|(_, c)| *c);
            let next = chars.get(index + 1).map(|(_, c)| *c);
            let next_is_space = next.map(|c| c.is_whitespace()).unwrap_or(true);
            let is_numeric_separator = matches!(ch, '.' | ':')
                && prev.is_some_and(|c| c.is_ascii_digit())
                && next.is_some_and(|c| c.is_ascii_digit());
            if is_numeric_separator {
                continue;
            }
            if next_is_space {
                let end = idx + ch.len_utf8();
                if end > start {
                    let candidate = &text[start..end];
                    if candidate.chars().any(|c| c.is_alphanumeric()) {
                        out.push(candidate);
                        start = end;
                    }
                }
            }
        }
    }

    if start < text.len() {
        out.push(&text[start..]);
    }

    out
}

fn split_long_sentence_by_whitespace(text: &str, max_chars: usize) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();

    for word in text.split_whitespace() {
        if current.is_empty() {
            if word.chars().count() > max_chars {
                out.push(word.to_string());
            } else {
                current.push_str(word);
            }
            continue;
        }

        let combined_len = current
            .chars()
            .count()
            .saturating_add(1)
            .saturating_add(word.chars().count());
        if combined_len > max_chars {
            out.push(current.trim().to_string());
            current.clear();
            current.push_str(word);
        } else {
            current.push(' ');
            current.push_str(word);
        }
    }

    if !current.trim().is_empty() {
        out.push(current.trim().to_string());
    }

    out
}

fn split_long_sentence_edge(text: &str) -> Vec<String> {
    split_long_sentence_edge_with_limit(text, EDGE_TTS_MAX_BYTES)
}

fn split_long_sentence_edge_with_limit(text: &str, max_bytes: usize) -> Vec<String> {
    let mut out = Vec::new();
    let mut remaining = text;
    while remaining.len() > max_bytes {
        let split_at = find_edge_split_idx(remaining, max_bytes);
        if split_at == 0 {
            // Guard against infinite loop on malformed input: consume at least one char.
            if let Some(first_char) = remaining.chars().next() {
                let advance = first_char.len_utf8();
                let chunk = remaining[..advance].trim();
                if !chunk.is_empty() {
                    out.push(chunk.to_string());
                }
                remaining = &remaining[advance..];
                continue;
            }
            break;
        }
        if split_at >= remaining.len() {
            break;
        }
        let (head, tail) = remaining.split_at(split_at);
        let chunk = head.trim();
        if !chunk.is_empty() {
            out.push(chunk.to_string());
        }
        remaining = tail;
    }
    let tail = remaining.trim();
    if !tail.is_empty() {
        out.push(tail.to_string());
    }
    out
}

fn parse_edge_binary_audio_payload(data: &[u8]) -> Result<Option<Vec<u8>>, String> {
    if data.len() < 2 {
        return Err("Edge WS: binary frame missing header length".to_string());
    }

    let be_len = u16::from_be_bytes([data[0], data[1]]) as usize;
    let le_len = u16::from_le_bytes([data[0], data[1]]) as usize;
    let header_len = if be_len > 0 && data.len() >= be_len + 2 {
        be_len
    } else if le_len > 0 && data.len() >= le_len + 2 {
        le_len
    } else {
        return Err("Edge WS: invalid binary header length".to_string());
    };

    let header_bytes = &data[2..2 + header_len];
    let payload = &data[2 + header_len..];
    let header_text = String::from_utf8_lossy(header_bytes);
    let mut path: Option<&str> = None;
    let mut content_type: Option<&str> = None;
    for line in header_text.split("\r\n") {
        if let Some((k, v)) = line.split_once(':') {
            let key = k.trim();
            let val = v.trim();
            if key.eq_ignore_ascii_case("Path") {
                path = Some(val);
            } else if key.eq_ignore_ascii_case("Content-Type") {
                content_type = Some(val);
            }
        }
    }

    if path != Some("audio") {
        return Err("Edge WS: binary frame path is not audio".to_string());
    }

    match content_type {
        Some(ct) if ct.eq_ignore_ascii_case("audio/mpeg") => {
            if payload.is_empty() {
                return Err("Edge WS: audio frame has empty payload".to_string());
            }
            Ok(Some(payload.to_vec()))
        }
        Some(_) => Err("Edge WS: unexpected binary content type".to_string()),
        None => {
            if payload.is_empty() {
                Ok(None)
            } else {
                Err("Edge WS: missing content type with non-empty payload".to_string())
            }
        }
    }
}

pub(crate) fn split_text_for_engine(text: &str, engine: TtsEngine) -> Vec<String> {
    if engine == TtsEngine::Edge {
        split_text_edge(text)
    } else {
        split_text(text)
    }
}

fn split_by_custom_dictionary(
    text: &str,
    _custom_entries: &[DictionaryEntry],
) -> Vec<(String, Option<VoiceOverride>, usize)> {
    if text.is_empty() {
        Vec::new()
    } else {
        vec![(text.to_string(), None, utf16_len(text))]
    }
}

pub fn split_into_tts_chunks(
    text: &str,
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
    default_engine: TtsEngine,
) -> Vec<TtsChunk> {
    let voice_spans = split_voice_tag_spans(text, default_engine);
    if voice_spans.is_empty() {
        return Vec::new();
    }

    let mut chunks: Vec<TtsChunk> = Vec::new();

    for (span_text, span_override, span_orig_len) in voice_spans {
        let span_engine = span_override
            .as_ref()
            .map(|v| v.engine)
            .unwrap_or(default_engine);
        if span_text.trim().is_empty() {
            if let Some(last) = chunks.last_mut() {
                last.original_len += span_orig_len;
            }
            continue;
        }
        let mut sentences = Vec::new();
        let mut current_sentence = String::new();
        let mut current_len = 0usize;
        let chars: Vec<char> = span_text.chars().collect();
        for (idx, ch) in chars.iter().copied().enumerate() {
            current_sentence.push(ch);
            current_len += 1;
            let is_terminal = matches!(ch, '.' | '!' | '?' | ';' | ':');
            let next_ch = chars.get(idx + 1).copied();
            let edge_dot_run =
                span_engine == TtsEngine::Edge && ch == '.' && matches!(next_ch, Some('.'));
            if is_terminal && !edge_dot_run {
                let should_split = span_engine != TtsEngine::Edge
                    || current_sentence.chars().any(|c| c.is_alphanumeric());
                if should_split && !current_sentence.trim().is_empty() {
                    sentences.push((current_sentence.clone(), current_len));
                    current_sentence.clear();
                    current_len = 0;
                }
            }
        }
        if !current_sentence.trim().is_empty() {
            sentences.push((current_sentence, current_len));
        }

        let extra_len = span_orig_len.saturating_sub(utf16_len(&span_text));
        let mut pending_len = extra_len;

        for (s_text, _s_len) in sentences.into_iter() {
            let cleaned = strip_dashed_lines(&s_text);
            let dict_segments = split_by_custom_dictionary(&cleaned, &[]);
            for (dict_text, _override_voice, dict_len) in dict_segments {
                let prepared = prepare_tts_text(&dict_text, split_on_newline, dictionary);
                if prepared.trim().is_empty() {
                    pending_len += dict_len;
                    continue;
                }
                let orig_len = dict_len + pending_len;
                pending_len = 0;
                if span_engine == TtsEngine::Edge {
                    let parts = split_text_edge(&prepared);
                    if parts.is_empty() {
                        continue;
                    }
                    let base = orig_len / parts.len();
                    let extra = orig_len % parts.len();
                    for (idx, part) in parts.into_iter().enumerate() {
                        let part_len = base + if idx < extra { 1 } else { 0 };
                        chunks.push(TtsChunk {
                            text_to_read: part,
                            original_len: part_len,
                            override_voice: span_override.clone(),
                        });
                    }
                } else {
                    chunks.push(TtsChunk {
                        text_to_read: prepared,
                        original_len: orig_len,
                        override_voice: span_override.clone(),
                    });
                }
            }
            if pending_len > 0
                && let Some(last) = chunks.last_mut()
            {
                last.original_len += pending_len;
            }
        }
    }
    chunks
}

fn split_tts_chunks_by_parts(chunks: &[TtsChunk], parts: usize) -> Vec<Vec<TtsChunk>> {
    if chunks.is_empty() {
        return Vec::new();
    }
    let parts = if parts == 0 { 1 } else { parts };
    let parts = if chunks.len() < parts {
        chunks.len()
    } else {
        parts
    };
    let per_part = chunks.len().div_ceil(parts);
    let mut out = Vec::new();
    for part_idx in 0..parts {
        let start = part_idx * per_part;
        let end = std::cmp::min(start + per_part, chunks.len());
        if start >= end {
            break;
        }
        out.push(chunks[start..end].to_vec());
    }
    out
}

fn start_audiobook_with_text(
    hwnd: HWND,
    mut text: String,
    suggested_name: Option<String>,
    mut epub_chapters: Option<Vec<String>>,
    is_unsaved_doc: bool,
    _doc_path: Option<&Path>,
) {
    let language = { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    if text.trim().is_empty() {
        show_error(hwnd, language, &settings::tts_no_text_message(language));
        return;
    }

    let (
        voice,
        split_on_newline,
        audiobook_split,
        audiobook_split_by_text,
        audiobook_split_text,
        audiobook_split_text_requires_newline,
        audiobook_split_by_epub_chapter,
        audiobook_split_by_time,
        audiobook_split_minutes,
        audiobook_split_start_number,
        audiobook_part_naming_mode,
        initial_audiobook_m4b_bitrate,
        tts_engine,
        dictionary,
        tts_rate,
        tts_pitch,
        tts_volume,
    ) = {
        with_state(hwnd, |state| {
            (
                state.settings.tts_voice.clone(),
                state.settings.split_on_newline,
                state.settings.audiobook_split,
                state.settings.audiobook_split_by_text,
                state.settings.audiobook_split_text.clone(),
                state.settings.audiobook_split_text_requires_newline,
                state.settings.audiobook_split_by_epub_chapter,
                state.settings.audiobook_split_by_time,
                state.settings.audiobook_split_minutes,
                state.settings.audiobook_split_start_number,
                state.settings.audiobook_part_naming_mode,
                state.settings.audiobook_m4b_bitrate,
                state.settings.tts_engine,
                state.settings.dictionary.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
            )
        })
    }
    .unwrap_or((
        "it-IT-IsabellaNeural".to_string(),
        true,
        0,
        false,
        String::new(),
        true,
        false,
        false,
        5,
        1,
        AudiobookPartNamingMode::TitleNumber,
        128,
        TtsEngine::Edge,
        Vec::new(),
        0,
        0,
        100,
    ));
    let dialogue_settings =
        { with_state(hwnd, |state| state.settings.clone()) }.unwrap_or_default();
    text = crate::dialogue_voice::apply_dialogue_tags_from_settings(&text, &dialogue_settings);
    if let Some(chapters) = epub_chapters.as_mut() {
        for chapter in chapters {
            *chapter = crate::dialogue_voice::apply_dialogue_tags_from_settings(
                chapter,
                &dialogue_settings,
            );
        }
    }

    let base_split_option_visible = audiobook_split_by_time
        || audiobook_split_by_text
        || audiobook_split > 1
        || (audiobook_split_by_epub_chapter && epub_chapters.as_ref().is_some());

    let cleaned = strip_dashed_lines(&text);
    let use_epub_split = audiobook_split_by_epub_chapter && epub_chapters.as_ref().is_some();
    let mixed_needed = if use_epub_split {
        epub_chapters
            .as_ref()
            .is_some_and(|chapters| chapters.iter().any(|chapter| has_voice_tags(chapter)))
    } else {
        has_voice_tags(&cleaned)
    };
    let mut split_by_time = audiobook_split_by_time;
    let split_minutes = audiobook_split_minutes.clamp(1, 60);
    let split_start_number = audiobook_split_start_number.clamp(1, 99);
    let mut split_parts = audiobook_split;
    let mut split_by_text = audiobook_split_by_text;
    if use_epub_split {
        split_by_time = false;
        split_by_text = false;
        split_parts = 0;
    }
    if split_by_time {
        split_parts = 0;
        split_by_text = false;
    }
    let mut marker_parts: Option<Vec<Vec<String>>> = None;
    let mut marker_positions: Option<Vec<usize>> = None;
    let mut marker_text: Option<String> = None;
    let mut selected_marker_titles: Option<Vec<String>> = None;
    let mut mixed_marker_parts: Option<Vec<Vec<TtsChunk>>> = None;
    let mut sapi4_threads: Option<u32> = None;

    if use_epub_split {
        let Some(chapters) = epub_chapters.as_ref() else {
            crate::log_debug("EPUB split requested without chapter data.");
            show_error(hwnd, language, &settings::tts_no_text_message(language));
            return;
        };
        if mixed_needed {
            mixed_marker_parts = build_mixed_audiobook_parts_from_sections(
                chapters,
                split_on_newline,
                &dictionary,
                tts_engine,
            );
        } else {
            marker_parts = build_audiobook_parts_from_sections(
                chapters,
                split_on_newline,
                &dictionary,
                tts_engine,
            );
        }
        if marker_parts.is_none() && mixed_marker_parts.is_none() {
            show_error(hwnd, language, &settings::tts_no_text_message(language));
            return;
        }
    }

    if tts_engine == TtsEngine::Sapi4 {
        let title = i18n::tr(language, "audiobook.sapi4_threads_title");

        let body = i18n::tr(language, "audiobook.sapi4_threads_body");

        if let Some(val_str) =
            crate::app_windows::prompt_window::prompt_user(hwnd, &title, &body, "30", language)
        {
            if let Ok(val) = val_str.parse::<u32>() {
                sapi4_threads = Some(val.clamp(1, 100));
            }
        } else {
            // User cancelled the prompt, abort the whole process

            return;
        }
    }

    if split_by_text {
        let (normalized, entries) = collect_marker_entries(
            &cleaned,
            &audiobook_split_text,
            audiobook_split_text_requires_newline,
        );
        if entries.is_empty() {
            split_parts = 0;
        } else {
            let labels: Vec<String> = entries.iter().map(|entry| entry.label.clone()).collect();
            let selected = crate::app_windows::marker_select_window::select_marker_entries(
                hwnd, &labels, language,
            );
            let Some(selected) = selected else {
                return;
            };
            let positions: Vec<usize> = selected
                .iter()
                .filter_map(|idx| entries.get(*idx).map(|e| e.pos))
                .collect();
            selected_marker_titles = Some(
                selected
                    .iter()
                    .filter_map(|idx| entries.get(*idx).map(|e| e.label.clone()))
                    .collect(),
            );
            if positions.is_empty() {
                split_parts = 0;
            } else {
                marker_positions = Some(positions.clone());
                marker_text = Some(normalized.clone());
                marker_parts = build_audiobook_parts_by_positions(
                    &normalized,
                    &positions,
                    split_on_newline,
                    &dictionary,
                    tts_engine,
                );
                if marker_parts.is_none() {
                    split_parts = 0;
                }
            }
        }
    }

    let prepared = if marker_parts.is_some() || mixed_needed {
        String::new()
    } else {
        prepare_tts_text(&cleaned, split_on_newline, &dictionary)
    };

    let chunks = if marker_parts.is_some() || mixed_needed {
        Vec::new()
    } else {
        split_text_for_engine(&prepared, tts_engine)
    };

    let mixed_chunks = if mixed_needed && marker_parts.is_none() && mixed_marker_parts.is_none() {
        Some(split_into_tts_chunks(
            &cleaned,
            split_on_newline,
            &dictionary,
            tts_engine,
        ))
    } else {
        None
    };
    if mixed_needed && mixed_marker_parts.is_none() {
        mixed_marker_parts = match (marker_text.as_ref(), marker_positions.as_ref()) {
            (Some(text), Some(positions)) => build_mixed_audiobook_parts_by_positions(
                text,
                positions,
                split_on_newline,
                &dictionary,
                tts_engine,
            ),
            _ => None,
        };
    }

    let expected_multi_file_split = if split_by_time {
        true
    } else if let Some(parts) = &marker_parts {
        parts.iter().filter(|p| !p.is_empty()).count() > 1
    } else if let Some(parts) = &mixed_marker_parts {
        parts.iter().filter(|p| !p.is_empty()).count() > 1
    } else if split_parts > 1 {
        std::cmp::min(split_parts as usize, chunks.len()) > 1
    } else {
        false
    };
    let split_into_multiple_files = expected_multi_file_split;
    let chapter_titles = if use_epub_split {
        epub_chapters.as_ref().map(|chs| {
            chs.iter()
                .map(|s| s.trim())
                .filter(|s| !s.is_empty())
                .map(|s| s.to_string())
                .collect::<Vec<_>>()
        })
    } else {
        selected_marker_titles.clone()
    };
    crate::log_debug(&format!(
        "Audiobook: split summary split_by_time={} split_parts={} chunks={} marker_parts={} mixed_marker_parts={} expected_multi_file_split={}",
        split_by_time,
        split_parts,
        chunks.len(),
        marker_parts.as_ref().map(|p| p.len()).unwrap_or(0),
        mixed_marker_parts.as_ref().map(|p| p.len()).unwrap_or(0),
        expected_multi_file_split
    ));
    let split_option_visible = if is_unsaved_doc {
        expected_multi_file_split
    } else {
        base_split_option_visible
    };
    let Some(save_result) =
        save_audio_dialog(hwnd, suggested_name.as_deref(), split_option_visible)
    else {
        return;
    };
    let mut output = save_result.path;
    let create_parts_folder = save_result.create_parts_folder;
    let audiobook_m4b_bitrate = {
        with_state(hwnd, |state| state.settings.audiobook_m4b_bitrate)
            .unwrap_or(initial_audiobook_m4b_bitrate)
    };
    crate::log_debug(&format!(
        "Audiobook: settings bitrate resolved to {} kbps for output {:?}",
        audiobook_m4b_bitrate, output
    ));

    if create_parts_folder && split_into_multiple_files {
        let nested_output = split_parts_output_in_subfolder(&output);
        if let Some(parent) = nested_output.parent()
            && let Err(e) = std::fs::create_dir_all(parent)
        {
            show_error(
                hwnd,
                language,
                &i18n::tr_f(language, "app.error_save_file", &[("err", &e.to_string())]),
            );
            return;
        }
        output = nested_output;
    }

    let chunks_len = if let Some(parts) = &mixed_marker_parts {
        parts.iter().map(|part| part.len()).sum()
    } else if let Some(parts) = &marker_parts {
        parts.iter().map(|part| part.len()).sum()
    } else if let Some(chunks) = &mixed_chunks {
        chunks.len()
    } else {
        chunks.len()
    };
    let progress_total = chunks_len;

    let cancel_token = Arc::new(AtomicBool::new(false));
    let progress_hwnd = {
        let h = crate::app_windows::audiobook_window::open(hwnd, progress_total);
        if with_state(hwnd, |state| {
            state.audiobook_progress = h;
            state.audiobook_cancel = Some(cancel_token.clone());
        })
        .is_none()
        {
            crate::log_debug("Failed to update audiobook progress state");
        }
        h
    };

    let cancel_clone = cancel_token.clone();
    std::thread::spawn(move || {
        let final_output = output.clone();
        let extension = final_output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
        let aac_with_chapters = is_aac && expected_multi_file_split;
        let emit_part_files = !is_aac || aac_with_chapters;
        let chapter_work_output = if aac_with_chapters {
            let stem = final_output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = final_output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("m4b");
            let stamp = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|d| d.as_millis())
                .unwrap_or(0);
            final_output.with_file_name(format!("{stem}.chapbuild.{stamp}.{ext}"))
        } else {
            final_output.clone()
        };
        let work_output_for_cleanup = chapter_work_output.clone();

        let options = AudiobookCommonOptions {
            voice: &voice,
            output: &chapter_work_output,
            progress_hwnd,
            cancel: cancel_clone,
            language,
            part_naming_mode: audiobook_part_naming_mode,
            audiobook_bitrate_kbps: audiobook_m4b_bitrate,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
            sapi4_threads,
        };
        let part_naming_mode = options.part_naming_mode;
        let engine_name = match tts_engine {
            TtsEngine::Edge => "edge",
            TtsEngine::Sapi5 => "sapi5",
            TtsEngine::Sapi4 => "sapi4",
        };
        crate::log_debug(&format!(
            "Audiobook: export start engine={} bitrate={} output={:?}",
            engine_name, options.audiobook_bitrate_kbps, options.output
        ));

        let mut time_split_done = false;
        let mut result = if mixed_needed {
            let mut current_global_progress = 0usize;
            let mixed_config = MixedAudiobookConfig {
                main_engine: tts_engine,
            };
            (|| {
                if let Some(ref parts) = mixed_marker_parts {
                    let parts_len = parts.len();
                    for (part_idx, part_chunks) in parts.iter().enumerate() {
                        if part_chunks.is_empty() {
                            continue;
                        }
                        let part_output = if parts_len > 1 && emit_part_files {
                            split_part_output(
                                options.output,
                                part_idx,
                                parts_len,
                                options.part_naming_mode,
                            )
                        } else {
                            options.output.to_path_buf()
                        };
                        render_mixed_audiobook_part(
                            part_chunks,
                            &mut current_global_progress,
                            &part_output,
                            &options,
                            &mixed_config,
                        )?;
                    }
                } else if let Some(ref chunks) = mixed_chunks {
                    let parts_count = if emit_part_files {
                        split_parts as usize
                    } else {
                        1
                    };
                    let parts = split_tts_chunks_by_parts(chunks, parts_count);
                    let parts_len = parts.len();
                    for (part_idx, part_chunks) in parts.iter().enumerate() {
                        let part_output = if parts_len > 1 && emit_part_files {
                            split_part_output(
                                options.output,
                                part_idx,
                                parts_len,
                                options.part_naming_mode,
                            )
                        } else {
                            options.output.to_path_buf()
                        };
                        render_mixed_audiobook_part(
                            part_chunks,
                            &mut current_global_progress,
                            &part_output,
                            &options,
                            &mixed_config,
                        )?;
                    }
                }
                Ok(())
            })()
        } else {
            match tts_engine {
                TtsEngine::Edge => {
                    if let Some(ref parts) = marker_parts {
                        if is_aac && !aac_with_chapters {
                            // Merge all marker parts into one for AAC
                            let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                            run_split_audiobook(&all_chunks, 0, options)
                        } else {
                            run_marker_split_audiobook(parts, options)
                        }
                    } else if split_by_time {
                        time_split_done = true;
                        run_split_audiobook_by_time_edge(
                            &chunks,
                            split_minutes,
                            split_start_number,
                            options,
                        )
                    } else {
                        run_split_audiobook(
                            &chunks,
                            if emit_part_files { split_parts } else { 0 },
                            options,
                        )
                    }
                }
                TtsEngine::Sapi4 => {
                    let voice_idx = parse_sapi4_voice_index(&voice);
                    if let Some(ref parts) = marker_parts {
                        if is_aac && !aac_with_chapters {
                            let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                            run_split_sapi4_audiobook(&all_chunks, voice_idx, 0, options)
                        } else {
                            run_marker_split_sapi4_audiobook(parts, voice_idx, options)
                        }
                    } else if split_by_time {
                        time_split_done = true;
                        run_split_audiobook_by_time_sapi(
                            TtsEngine::Sapi4,
                            &chunks,
                            split_minutes,
                            split_start_number,
                            options,
                        )
                    } else {
                        run_split_sapi4_audiobook(
                            &chunks,
                            voice_idx,
                            if emit_part_files { split_parts } else { 0 },
                            options,
                        )
                    }
                }
                TtsEngine::Sapi5 => {
                    if let Some(ref parts) = marker_parts {
                        if is_aac && !aac_with_chapters {
                            let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                            run_split_sapi_audiobook(&all_chunks, 0, options)
                        } else {
                            run_marker_split_sapi_audiobook(parts, options)
                        }
                    } else if split_by_time {
                        time_split_done = true;
                        run_split_audiobook_by_time_sapi(
                            TtsEngine::Sapi5,
                            &chunks,
                            split_minutes,
                            split_start_number,
                            options,
                        )
                    } else {
                        run_split_sapi_audiobook(
                            &chunks,
                            if emit_part_files { split_parts } else { 0 },
                            options,
                        )
                    }
                }
            }
        };
        if result.is_ok() && split_by_time && !time_split_done {
            let stem = final_output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = final_output
                .extension()
                .and_then(|s| s.to_str())
                .map(str::trim)
                .filter(|s| !s.is_empty())
                .unwrap_or("mp3");
            const TIME_SPLIT_PART_WIDTH: usize = 4;
            let pattern = final_output.with_file_name(match part_naming_mode {
                AudiobookPartNamingMode::TitleNumber => {
                    format!("{stem} Part %0{TIME_SPLIT_PART_WIDTH}d.{ext}")
                }
                AudiobookPartNamingMode::NumberOnly => {
                    format!("%0{TIME_SPLIT_PART_WIDTH}d.{ext}")
                }
                AudiobookPartNamingMode::NumberTitle => {
                    format!("%0{TIME_SPLIT_PART_WIDTH}d - {stem}.{ext}")
                }
            });
            let segment_seconds = split_minutes.saturating_mul(60);
            result = crate::ffmpeg_export::segment_audio_file(
                &final_output,
                &pattern,
                segment_seconds,
                split_start_number,
            );
            if result.is_ok()
                && let Err(e) = std::fs::remove_file(&final_output)
            {
                crate::log_debug(&format!(
                    "Failed to remove original audiobook after segmenting: {}",
                    e
                ));
            }
        }

        let mut success = result.is_ok();
        if success && aac_with_chapters {
            match collect_split_part_files(&work_output_for_cleanup, part_naming_mode) {
                Ok(part_files) if part_files.len() > 1 => {
                    match merge_m4b_parts_with_chapters(
                        &part_files,
                        &final_output,
                        chapter_titles.as_deref(),
                    ) {
                        Ok(()) => {
                            for file in &part_files {
                                if let Err(e) = std::fs::remove_file(file) {
                                    crate::log_debug(&format!(
                                        "M4B chapter merge cleanup: failed to remove part {}: {}",
                                        file.display(),
                                        e
                                    ));
                                }
                            }
                        }
                        Err(e) => {
                            crate::log_debug(&format!("M4B chapter merge failed: {}", e));
                            result = Err(e);
                            success = false;
                        }
                    }
                }
                Ok(_) => {
                    result = Err("M4B chapter merge: not enough part files".to_string());
                    success = false;
                }
                Err(e) => {
                    result = Err(e);
                    success = false;
                }
            }
            if work_output_for_cleanup.exists()
                && let Err(e) = std::fs::remove_file(&work_output_for_cleanup)
            {
                crate::log_debug(&format!(
                    "M4B chapter merge cleanup: failed to remove temp output {}: {}",
                    work_output_for_cleanup.display(),
                    e
                ));
            }
        }

        if success && is_aac && !split_by_time && !aac_with_chapters {
            // Set metadata for the single M4B file
            let file_title = final_output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("Audiobook");
            let mut comment = String::new();

            // If it was supposed to be split, we can provide a "table of contents" in comments
            if split_parts > 1 || marker_parts.is_some() {
                comment.push_str("Chapters:\n");
                // Note: accurate timestamps would require dry-run or parsing generated audio.
                // For now, we provide the labels if using markers.
                if let Some(ref parts) = marker_parts {
                    for (i, _p) in parts.iter().enumerate() {
                        comment.push_str(&format!("Part {}\n", i + 1));
                    }
                }
            }

            if crate::audio_utils::set_file_metadata(
                &final_output,
                Some(file_title),
                Some("Sonarpad"),
                if comment.is_empty() {
                    None
                } else {
                    Some(&comment)
                },
            )
            .is_err()
            {}
        }

        let message = match result {
            Ok(()) => i18n::tr(language, "tts.audiobook_saved"),
            Err(err) => {
                if err == "Cancelled" {
                    cancelled_message(language)
                } else {
                    err
                }
            }
        };
        let payload = Box::new(AudiobookResult { success, message });
        unsafe {
            if let Err(e) = PostMessageW(
                hwnd,
                crate::WM_TTS_AUDIOBOOK_DONE,
                WPARAM(0),
                LPARAM(Box::into_raw(payload) as isize),
            ) {
                crate::log_debug(&format!("Failed to post WM_TTS_AUDIOBOOK_DONE: {}", e));
            }
        }
    });
}

pub fn start_audiobook(hwnd: HWND) {
    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    let (suggested_name, doc_path, split_epub, language, is_unsaved_doc) = {
        with_state(hwnd, |state| {
            state.docs.get(state.current).map(|doc| {
                let p = Path::new(&doc.title);
                (
                    p.file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or(&doc.title)
                        .to_string(),
                    doc.path.clone(),
                    state.settings.audiobook_split_by_epub_chapter,
                    state.settings.language,
                    doc.path.is_none(),
                )
            })
        })
    }
    .flatten()
    .unwrap_or((String::new(), None, false, Language::Italian, true));

    let mut epub_chapters = None;
    if split_epub
        && let Some(ref path) = doc_path
        && is_epub_path(path)
    {
        match read_epub_chapters(path, language) {
            Ok(chapters) => {
                epub_chapters = Some(chapters);
            }
            Err(err) => {
                show_error(hwnd, language, &err);
                return;
            }
        }
    }

    let suggested_name = if suggested_name.is_empty() {
        None
    } else {
        Some(suggested_name)
    };
    start_audiobook_with_text(
        hwnd,
        text,
        suggested_name,
        epub_chapters,
        is_unsaved_doc,
        doc_path.as_deref(),
    );
}

pub fn start_audiobook_from_selection(hwnd: HWND) {
    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let text = crate::editor_manager::get_selected_text(hwnd_edit);
    let Some(text) = text else {
        let language = { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
        show_error(hwnd, language, &settings::tts_no_text_message(language));
        return;
    };
    let suggested_name = {
        with_state(hwnd, |state| {
            state.docs.get(state.current).map(|doc| {
                let p = Path::new(&doc.title);
                (
                    p.file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or(&doc.title)
                        .to_string(),
                    doc.path.clone(),
                    doc.path.is_none(),
                )
            })
        })
    }
    .flatten();
    let (suggested_name, doc_path, is_unsaved_doc) = suggested_name
        .map(|(name, path, unsaved)| (Some(name), path, unsaved))
        .unwrap_or((None, None, true));
    start_audiobook_with_text(
        hwnd,
        text,
        suggested_name,
        None,
        is_unsaved_doc,
        doc_path.as_deref(),
    );
}

pub(crate) fn parse_sapi4_voice_index(voice: &str) -> i32 {
    if let Some(hash_pos) = voice.find('#') {
        let rest = &voice[hash_pos + 1..];
        if let Some(pipe_pos) = rest.find('|') {
            rest[..pipe_pos].parse::<i32>().unwrap_or(1)
        } else {
            rest.parse::<i32>().unwrap_or(1)
        }
    } else {
        1
    }
}

fn run_split_audiobook(
    chunks: &[String],
    split_parts: u32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let parts = if split_parts == 0 {
        1
    } else {
        split_parts as usize
    };
    let total_chunks = chunks.len();

    // Se ci sono meno chunks delle parti richieste, riduciamo le parti
    let parts = if total_chunks < parts {
        total_chunks
    } else {
        parts
    };
    let chunks_per_part = total_chunks.div_ceil(parts);

    let mut current_global_progress = 0;

    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }

        let part_chunks = &chunks[start_idx..end_idx];

        let part_output = if parts > 1 {
            split_part_output(options.output, part_idx, parts, options.part_naming_mode)
        } else {
            options.output.to_path_buf()
        };

        // Create a temporary options struct with the correct output path for this part
        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            part_naming_mode: options.part_naming_mode,
            audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            sapi4_threads: options.sapi4_threads,
        };

        run_tts_audiobook_part(part_chunks, &mut current_global_progress, &part_options)?;
    }
    Ok(())
}

fn run_marker_split_audiobook(
    parts: &[Vec<String>],
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let parts_len = parts.len();
    let mut current_global_progress = 0;

    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if part_chunks.is_empty() {
            continue;
        }
        let part_output = if parts_len > 1 {
            split_part_output(
                options.output,
                part_idx,
                parts_len,
                options.part_naming_mode,
            )
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            part_naming_mode: options.part_naming_mode,
            audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            sapi4_threads: options.sapi4_threads,
        };

        run_tts_audiobook_part(part_chunks, &mut current_global_progress, &part_options)?;
    }
    Ok(())
}

fn run_split_sapi4_audiobook(
    chunks: &[String],
    voice_idx: i32,
    split_parts: u32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let parts_count = if split_parts == 0 {
        1
    } else {
        split_parts as usize
    };
    let total_chunks = chunks.len();
    let parts_count = if total_chunks < parts_count {
        total_chunks
    } else {
        parts_count
    };

    let chunks_per_part = total_chunks.div_ceil(parts_count);
    let mut current_global_progress = 0;

    for part_idx in 0..parts_count {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }

        let part_chunks = &chunks[start_idx..end_idx];
        let part_output = if parts_count > 1 {
            split_part_output(
                options.output,
                part_idx,
                parts_count,
                options.part_naming_mode,
            )
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            part_naming_mode: options.part_naming_mode,
            audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            sapi4_threads: options.sapi4_threads,
        };

        run_sapi4_parallel_part(
            part_chunks,
            voice_idx,
            &mut current_global_progress,
            &part_options,
        )?;
    }
    Ok(())
}

fn run_marker_split_sapi4_audiobook(
    parts: &[Vec<String>],
    voice_idx: i32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let mut current_global_progress = 0;
    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if part_chunks.is_empty() {
            continue;
        }

        let part_output = if parts.len() > 1 {
            split_part_output(
                options.output,
                part_idx,
                parts.len(),
                options.part_naming_mode,
            )
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            part_naming_mode: options.part_naming_mode,
            audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            sapi4_threads: options.sapi4_threads,
        };

        run_sapi4_parallel_part(
            part_chunks,
            voice_idx,
            &mut current_global_progress,
            &part_options,
        )?;
    }
    Ok(())
}

pub(crate) fn run_sapi4_parallel_part(
    chunks: &[String],
    voice_idx: i32,
    global_progress: &mut usize,
    options: &AudiobookCommonOptions,
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }

    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_mp3 = extension == "mp3";

    // Parallel processing for chunks within this part
    let temp_dir = options
        .output
        .parent()
        .unwrap_or(Path::new("."))
        .join(format!(
            "sapi4_tmp_{}",
            options
                .output
                .file_name()
                .and_then(|s| s.to_str())
                .unwrap_or("part")
        ));
    std::fs::create_dir_all(&temp_dir).ok();

    let mut internal_parts = Vec::new();
    let pool_size = options.sapi4_threads.unwrap_or(30) as usize;
    let chunks_count = chunks.len();
    let sub_parts_count = if chunks_count < pool_size {
        chunks_count
    } else {
        pool_size
    };
    let chunks_per_sub = chunks_count.div_ceil(sub_parts_count);

    for i in 0..sub_parts_count {
        let s = i * chunks_per_sub;
        let e = std::cmp::min(s + chunks_per_sub, chunks_count);
        if s >= e {
            break;
        }
        let sub_chunks = chunks[s..e].to_vec();
        let sub_output = temp_dir.join(format!("sub_{}.wav", i));
        internal_parts.push((sub_chunks, sub_output));
    }

    let progress_counter = Arc::new(std::sync::atomic::AtomicUsize::new(*global_progress));
    let (tx, rx) = std::sync::mpsc::channel::<Result<PathBuf, String>>();
    let parts_shared = Arc::new(Mutex::new(internal_parts.into_iter()));
    let mut handles = Vec::new();

    for _ in 0..pool_size {
        let tx = tx.clone();
        let parts_shared = parts_shared.clone();
        let progress_counter = progress_counter.clone();
        let cancel_token = options.cancel.clone();
        let (r, p, v) = (options.rate, options.pitch, options.volume);
        let mp3_bitrate_kbps = options.audiobook_bitrate_kbps;
        let progress_hwnd = options.progress_hwnd;

        let handle = std::thread::spawn(move || {
            loop {
                let part = {
                    let mut guard = parts_shared.lock().unwrap_or_else(|e| e.into_inner());
                    guard.next()
                };
                let Some((sub_chunks, sub_output)) = part else {
                    break;
                };
                if cancel_token.load(Ordering::Relaxed) {
                    tx.send(Err("Cancelled".to_string())).ok();
                    break;
                }

                let res = crate::sapi4_engine::speak_sapi4_to_file(
                    &sub_chunks,
                    voice_idx,
                    &sub_output,
                    crate::sapi4_engine::Sapi4Options {
                        rate: r,
                        pitch: p,
                        volume: v,
                        mp3_bitrate_kbps,
                        cancel: cancel_token.clone(),
                    },
                    |_| {
                        let current = progress_counter.fetch_add(1, Ordering::SeqCst) + 1;
                        if progress_hwnd.0 != 0 {
                            unsafe {
                                if PostMessageW(
                                    progress_hwnd,
                                    crate::WM_UPDATE_PROGRESS,
                                    WPARAM(current),
                                    LPARAM(0),
                                )
                                .ok()
                                .is_some()
                                {}
                            }
                        }
                    },
                );

                if let Err(e) = res {
                    std::fs::remove_file(&sub_output).ok();
                    tx.send(Err(e)).ok();
                    break;
                } else {
                    // Only parallel encode for MP3 because they can be concatenated binary.
                    // AAC/M4B must be joined as WAV first and then encoded as a whole.
                    if is_mp3 {
                        let encoded_sub = sub_output.with_extension("mp3");
                        let ff_settings = crate::ffmpeg_export::ConvertAudioSettings {
                            format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
                            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                                mp3_bitrate_kbps,
                            ),
                        };
                        let mut ff_progress = |_p: u32| {};
                        if let Err(e) = crate::ffmpeg_export::convert_audio_file(
                            &sub_output,
                            &encoded_sub,
                            &ff_settings,
                            None,
                            Some(&mut ff_progress),
                        ) {
                            std::fs::remove_file(&sub_output).ok();
                            tx.send(Err(format!("Parallel audio encode failed: {}", e)))
                                .ok();
                            break;
                        }
                        std::fs::remove_file(&sub_output).ok();
                        tx.send(Ok(encoded_sub)).ok();
                    } else {
                        // Keep as WAV for now (M4B or standard WAV)
                        tx.send(Ok(sub_output)).ok();
                    }
                }
            }
        });
        handles.push(handle);
    }

    {
        let _tx = tx;
    }
    let mut produced_files: Vec<PathBuf> = Vec::new();
    let mut error = None;
    for res in rx {
        match res {
            Ok(path) => produced_files.push(path),
            Err(e) => {
                if error.is_none() {
                    error = Some(e);
                }
            }
        }
    }
    for h in handles {
        h.join().ok();
    }

    if error.is_some() || options.cancel.load(Ordering::Relaxed) {
        std::fs::remove_dir_all(&temp_dir).ok();
        let msg = if let Some(e) = error {
            e
        } else {
            "Cancelled".to_string()
        };
        return Err(if msg == "Cancelled" {
            cancelled_message(options.language)
        } else {
            msg
        });
    }

    produced_files.sort_by_key(|p: &PathBuf| {
        let name = p.file_stem().and_then(|s| s.to_str()).unwrap_or("");
        parse_sapi4_part_index(name)
    });

    let result = if is_mp3 {
        merge_and_finalize_sapi4_mp3(
            &produced_files,
            options.output,
            options.language,
            options.audiobook_bitrate_kbps,
            options.progress_hwnd,
        )
    } else {
        // This covers M4B (AAC) and standard WAV.
        // For M4B, it joins WAV chunks and then encodes to AAC in one pass.
        merge_and_finalize_sapi4_audio(&produced_files, options.output, options.language, options)
    };

    std::fs::remove_dir_all(&temp_dir).ok();
    *global_progress = progress_counter.load(Ordering::SeqCst);
    result
}

fn parse_sapi4_part_index(name: &str) -> usize {
    if let Some(rest) = name.strip_prefix("sub_")
        && let Ok(value) = rest.parse::<usize>()
    {
        return value;
    }
    if let Some(rest) = name.strip_prefix("part_")
        && let Ok(value) = rest.parse::<usize>()
    {
        return value;
    }
    let mut digits = Vec::new();
    for ch in name.chars().rev() {
        if ch.is_ascii_digit() {
            digits.push(ch);
        } else if !digits.is_empty() {
            break;
        }
    }
    if digits.is_empty() {
        return 0;
    }
    digits.reverse();
    digits
        .into_iter()
        .collect::<String>()
        .parse::<usize>()
        .unwrap_or(0)
}

fn merge_and_finalize_sapi4_mp3(
    mp3_files: &[PathBuf],
    output: &Path,
    language: Language,
    bitrate_kbps: u32,
    progress_hwnd: HWND,
) -> Result<(), String> {
    if mp3_files.is_empty() {
        return Ok(());
    }

    let output_name = output
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("audiobook.mp3");
    let concat_output = output.with_file_name(format!("{output_name}.concat.tmp.mp3"));

    let mut out_file = std::fs::File::create(&concat_output).map_err(|e| e.to_string())?;
    for path in mp3_files {
        let mut in_file = std::fs::File::open(path).map_err(|e| e.to_string())?;
        std::io::copy(&mut in_file, &mut out_file).map_err(|e| e.to_string())?;
    }

    set_audiobook_progress_phase(progress_hwnd, true);
    set_audiobook_progress_total(progress_hwnd, 100);
    post_audiobook_progress(progress_hwnd, 0);
    let settings = crate::ffmpeg_export::ConvertAudioSettings {
        format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
        quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(bitrate_kbps),
    };
    let mut last_progress = 0usize;
    let mut progress = |p: u32| {
        let current = ((p as usize) / 100).min(99);
        if current > last_progress {
            last_progress = current;
            post_audiobook_progress(progress_hwnd, current);
        }
    };
    let convert_result = crate::ffmpeg_export::convert_audio_file(
        &concat_output,
        output,
        &settings,
        None,
        Some(&mut progress),
    );
    if let Err(err) = std::fs::remove_file(&concat_output) {
        crate::log_debug(&format!(
            "Failed to remove temporary concatenated MP3 {:?}: {}",
            concat_output, err
        ));
    }
    match convert_result {
        Ok(()) => {
            post_audiobook_progress(progress_hwnd, 100);
            Ok(())
        }
        Err(err) => Err(i18n::tr_f(language, "sapi5.mf_error", &[("err", &err)])),
    }
}

#[cfg(test)]
mod tests {
    use crate::settings::DictionaryEntry;

    use super::{
        TtsEngine, build_audiobook_parts_by_positions, collect_marker_entries, find_edge_split_idx,
        is_edge_text_usable, normalize_for_tts, parse_edge_binary_audio_payload,
        parse_sapi4_part_index, parse_voice_tag_override, prepare_tts_text, preview_for_log,
        render_edge_ssml_text_with_pause_tags, render_sapi_ssml_text_with_pause_tags,
        sanitize_edge_text, split_into_tts_chunks, split_long_sentence_edge_with_limit,
        split_sentences, split_text_for_engine, split_voice_tag_spans, strip_dashed_lines,
    };

    #[test]
    fn prepare_tts_text_keeps_dictionary_entries_case_sensitive_by_default() {
        let dictionary = vec![DictionaryEntry {
            original: "ciao".to_string(),
            replacement: "salve".to_string(),
            match_case: true,
            use_custom_voice: false,
            custom_voice_engine: None,
            custom_voice: None,
        }];

        assert_eq!(
            prepare_tts_text("Ciao ciao", false, &dictionary),
            "Ciao salve"
        );
    }

    #[test]
    fn prepare_tts_text_supports_case_insensitive_dictionary_entries() {
        let dictionary = vec![DictionaryEntry {
            original: "ciao".to_string(),
            replacement: "salve".to_string(),
            match_case: false,
            use_custom_voice: false,
            custom_voice_engine: None,
            custom_voice: None,
        }];

        assert_eq!(
            prepare_tts_text("Ciao cIAO ciao", false, &dictionary),
            "salve salve salve"
        );
    }

    #[test]
    fn pause_tags_render_as_edge_breaks_from_raw_or_escaped_text() {
        assert_eq!(
            render_edge_ssml_text_with_pause_tags("Ciao <pause ms=\"500\"/> dopo"),
            "Ciao <mstts:silence type=\"Sentenceboundary\" value=\"500ms\"/> dopo"
        );
        assert_eq!(
            render_edge_ssml_text_with_pause_tags("Ciao &lt;pause ms=&quot;1000&quot;/&gt; dopo"),
            "Ciao <mstts:silence type=\"Sentenceboundary\" value=\"1000ms\"/> dopo"
        );
    }

    #[test]
    fn pause_tags_render_as_sapi_silence() {
        assert_eq!(
            render_sapi_ssml_text_with_pause_tags("Ciao <pause ms=\"500\"/> dopo"),
            "Ciao <silence msec=\"500\"/> dopo"
        );
    }

    #[test]
    fn parse_sapi4_part_index_prefers_named_prefix() {
        assert_eq!(parse_sapi4_part_index("sub_0"), 0);
        assert_eq!(parse_sapi4_part_index("sub_12"), 12);
        assert_eq!(parse_sapi4_part_index("part_3"), 3);
        assert_eq!(parse_sapi4_part_index("part_20"), 20);
    }

    #[test]
    fn parse_sapi4_part_index_falls_back_to_trailing_digits() {
        assert_eq!(parse_sapi4_part_index("segment_0007"), 7);
        assert_eq!(parse_sapi4_part_index("chunk99"), 99);
        assert_eq!(parse_sapi4_part_index("audio"), 0);
        assert_eq!(parse_sapi4_part_index("sub_foo_5"), 5);
    }

    #[test]
    fn sapi4_part_sort_orders_by_index() {
        let mut paths = vec![
            std::path::PathBuf::from("sub_2.mp3"),
            std::path::PathBuf::from("sub_10.mp3"),
            std::path::PathBuf::from("sub_1.mp3"),
            std::path::PathBuf::from("part_3.mp3"),
            std::path::PathBuf::from("foo9.mp3"),
        ];
        paths.sort_by_key(|path| {
            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            parse_sapi4_part_index(stem)
        });
        let names: Vec<String> = paths
            .iter()
            .map(|p| {
                p.file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("")
                    .to_string()
            })
            .collect();
        assert_eq!(
            names,
            vec![
                "sub_1.mp3",
                "sub_2.mp3",
                "part_3.mp3",
                "foo9.mp3",
                "sub_10.mp3"
            ]
        );
    }

    #[test]
    fn sapi4_part_sort_handles_dirs_and_extensions() {
        let mut paths = vec![
            std::path::PathBuf::from("C:\\tmp\\sapi4\\sub_2.wav"),
            std::path::PathBuf::from("C:\\tmp\\sapi4\\part_12.mp3"),
            std::path::PathBuf::from("C:\\tmp\\sapi4\\foo9.m4a"),
            std::path::PathBuf::from("C:\\tmp\\sapi4\\sub_1.mp3"),
            std::path::PathBuf::from("C:\\tmp\\sapi4\\chunk_0003.wav"),
        ];
        paths.sort_by_key(|path| {
            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            parse_sapi4_part_index(stem)
        });
        let names: Vec<String> = paths
            .iter()
            .map(|p| {
                p.file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("")
                    .to_string()
            })
            .collect();
        assert_eq!(
            names,
            vec![
                "sub_1.mp3",
                "sub_2.wav",
                "chunk_0003.wav",
                "foo9.m4a",
                "part_12.mp3"
            ]
        );
    }

    #[test]
    fn split_voice_tag_spans_decodes_basic_xml_entities_in_text() {
        let input = r#"<voice engine="edge" voice="it-IT-DiegoNeural">&quot;Ciao &amp; mondo&quot;</voice>"#;
        let spans = split_voice_tag_spans(input, TtsEngine::Edge);
        assert_eq!(spans.len(), 1);
        assert_eq!(spans[0].0, "\"Ciao & mondo\"");
        assert!(spans[0].1.is_some());
    }

    #[test]
    fn parse_voice_tag_override_accepts_rate_pitch_volume() {
        let tag = r#"engine="edge" voice="it-IT-DiegoNeural" rate="-80" pitch="8" volume="140""#;
        let ov = parse_voice_tag_override(tag, TtsEngine::Edge).expect("override expected");
        assert_eq!(ov.rate, Some(-80));
        assert_eq!(ov.pitch, Some(8));
        assert_eq!(ov.volume, Some(140));
    }

    #[test]
    fn edge_normalizes_ascii_ellipsis_to_single_dot() {
        assert_eq!(sanitize_edge_text("Ciao... come va..."), "Ciao. come va.");
    }

    #[test]
    fn edge_sanitize_handles_weird_spaces_and_symbols() {
        assert_eq!(
            sanitize_edge_text("ciao\u{00A0}\u{200B}mondo\u{2409}!!!???"),
            "ciao mondo!?"
        );
    }

    #[test]
    fn edge_sanitize_normalizes_dot_quote_colon_sequences() {
        let text = "\"Addio Juve, secondo me andrà...\": parla chi gli ha cambiato la vita";
        let sanitized = sanitize_edge_text(text);
        assert!(!sanitized.contains(".\":"));
        assert!(sanitized.contains("andrà."));
    }

    #[test]
    fn edge_split_does_not_emit_quote_colon_only_chunk_after_ellipsis() {
        let chunks = split_into_tts_chunks(
            "\"Vlahovic distrugge reti e lascia! Addio Juve, secondo me andrà...\": parla chi gli ha cambiato la vita",
            true,
            &[],
            TtsEngine::Edge,
        );
        assert!(!chunks.is_empty());
        let has_bad_chunk = chunks.iter().any(|c| {
            let probe = c
                .text_to_read
                .replace("&quot;", "\"")
                .replace("&apos;", "'")
                .chars()
                .filter(|ch| !ch.is_whitespace())
                .collect::<String>();
            !probe.chars().any(|ch| ch.is_alphanumeric())
        });
        assert!(
            !has_bad_chunk,
            "Edge chunk split produced punctuation-only chunk"
        );
    }

    #[test]
    fn edge_text_usable_rejects_punctuation_only() {
        assert!(!is_edge_text_usable("...?!"));
        assert!(is_edge_text_usable("ciao."));
    }

    #[test]
    fn edge_split_into_chunks_does_not_create_dot_only_chunks_from_ellipsis() {
        let chunks = split_into_tts_chunks(
            "ciao io sono ambro... e sono un campione... e sono un grande atleta...",
            true,
            &[],
            TtsEngine::Edge,
        );
        assert!(!chunks.is_empty());
        assert!(chunks.iter().all(|c| c.text_to_read.trim() != "."));
    }

    #[test]
    fn split_sentences_keeps_number_groups_intact() {
        let text = "Sono chiamati alle urne 51.424.729 cittadini, tra cui 5.477.619 residenti all’estero. Si vota alle 12:30.";
        assert_eq!(
            split_sentences(text),
            vec![
                "Sono chiamati alle urne 51.424.729 cittadini, tra cui 5.477.619 residenti all’estero.",
                " Si vota alle 12:30."
            ]
        );
    }

    #[test]
    fn edge_split_idx_prefers_newline_or_space_and_stays_utf8_safe() {
        let text = "uno due\ntre quattro cinque";
        let split = find_edge_split_idx(text, 10);
        assert!(split > 0 && split <= 10);
        assert!(text.is_char_boundary(split));
    }

    #[test]
    fn edge_split_long_sentence_never_loops_and_preserves_order() {
        let text = "áááááááááááááááááááá";
        let parts = split_long_sentence_edge_with_limit(text, 3);
        assert!(!parts.is_empty());
        let joined: String = parts.join("");
        assert_eq!(joined, text);
    }

    #[test]
    fn edge_binary_audio_payload_parses_valid_audio_frame() {
        let header = b"Path:audio\r\nContent-Type:audio/mpeg";
        let payload = b"\x01\x02\x03";
        let mut frame = Vec::new();
        frame.extend_from_slice(&(header.len() as u16).to_be_bytes());
        frame.extend_from_slice(header);
        frame.extend_from_slice(payload);
        let out = parse_edge_binary_audio_payload(&frame).expect("valid frame should parse");
        assert_eq!(out, Some(payload.to_vec()));
    }

    #[test]
    fn edge_binary_audio_payload_rejects_non_audio_path() {
        let header = b"Path:turn.end\r\nContent-Type:audio/mpeg";
        let payload = b"\x01";
        let mut frame = Vec::new();
        frame.extend_from_slice(&(header.len() as u16).to_be_bytes());
        frame.extend_from_slice(header);
        frame.extend_from_slice(payload);
        let err = parse_edge_binary_audio_payload(&frame).expect_err("non-audio path must fail");
        assert!(err.contains("path is not audio"));
    }

    #[test]
    fn marker_split_with_chapter_and_m4b_merge_does_not_fail() {
        let text = "Chapter 1\nHello world.\n\nChapter 2\nMore text.\n";
        let (normalized, entries) = collect_marker_entries(text, "Chapter", true);
        assert_eq!(entries.len(), 2);
        let positions: Vec<usize> = entries.iter().map(|e| e.pos).collect();
        let parts =
            build_audiobook_parts_by_positions(&normalized, &positions, true, &[], TtsEngine::Edge)
                .expect("marker split should produce parts");
        let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
        assert!(!all_chunks.is_empty());
    }

    #[test]
    fn sapi4_audiobook_chunks_merge_in_order() {
        let mut produced_files = vec![
            std::path::PathBuf::from("sub_10.wav"),
            std::path::PathBuf::from("sub_2.wav"),
            std::path::PathBuf::from("sub_1.wav"),
            std::path::PathBuf::from("sub_3.wav"),
        ];
        produced_files.sort_by_key(|p| {
            let name = p.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            parse_sapi4_part_index(name)
        });
        let names: Vec<String> = produced_files
            .iter()
            .map(|p| {
                p.file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("")
                    .to_string()
            })
            .collect();
        assert_eq!(
            names,
            vec!["sub_1.wav", "sub_2.wav", "sub_3.wav", "sub_10.wav"]
        );
    }

    #[test]
    fn audiobook_liturgia_no_skipped_or_reordered_lines() {
        let path_owned = std::env::var("SONARPAD_TTS_TEST_PATH")
            .unwrap_or_else(|_| r"C:\Users\ambro\Downloads\05-02-2026.txt".to_string());
        let path = path_owned.as_str();
        let raw = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("Skipping test – cannot read {path}: {e}");
                return;
            }
        };

        let cleaned = strip_dashed_lines(&raw);
        let chunks = split_into_tts_chunks(&cleaned, true, &[], TtsEngine::Edge);

        assert!(
            !chunks.is_empty(),
            "split_into_tts_chunks produced 0 chunks"
        );

        // Reassemble all chunk text into a single string for searching
        let reassembled: String = chunks
            .iter()
            .map(|c| c.text_to_read.as_str())
            .collect::<Vec<_>>()
            .join(" ");

        // Normalize the cleaned text the same way: split_on_newline=true, no dictionary
        let normalized_input = normalize_for_tts(&cleaned, true);

        // Collect non-empty input lines with their 1-based line numbers
        let input_lines: Vec<(usize, &str)> = normalized_input
            .lines()
            .enumerate()
            .filter_map(|(i, line)| {
                let trimmed = line.trim();
                if trimmed.is_empty() {
                    None
                } else {
                    Some((i + 1, trimmed))
                }
            })
            .collect();

        assert!(!input_lines.is_empty(), "Input file has no non-empty lines");

        // For each input line, extract significant words (3+ chars) and check they
        // appear in the reassembled chunks. Pure punctuation or very short tokens
        // may be merged or transformed by XML escaping, so we check word presence.
        let mut missing_lines: Vec<(usize, String)> = Vec::new();

        for &(line_num, line_text) in &input_lines {
            // Extract words of 3+ alphabetic chars from this line
            let words: Vec<&str> = line_text
                .split(|c: char| !c.is_alphanumeric())
                .filter(|w| w.len() >= 3)
                .collect();

            if words.is_empty() {
                // Line is only short tokens / punctuation – skip
                continue;
            }

            // Check that at least half the significant words appear in the reassembled text
            let found_count = words.iter().filter(|w| reassembled.contains(**w)).count();

            if found_count < (words.len() + 1) / 2 {
                missing_lines.push((line_num, line_text.to_string()));
            }
        }

        if !missing_lines.is_empty() {
            let report: String = missing_lines
                .iter()
                .take(50)
                .map(|(num, text)| format!("  line {num}: {text}"))
                .collect::<Vec<_>>()
                .join("\n");
            panic!(
                "{} lines appear to be skipped by the TTS chunking pipeline:\n{report}",
                missing_lines.len()
            );
        }

        // Verify ordering: for each pair of consecutive input lines, check that
        // the first significant word of line N appears before the first significant
        // word of line N+1 in the reassembled text.
        let mut out_of_order: Vec<(usize, usize)> = Vec::new();
        let mut last_pos: usize = 0;

        for &(line_num, line_text) in &input_lines {
            let first_word = line_text
                .split(|c: char| !c.is_alphanumeric())
                .find(|w| w.len() >= 3);

            let Some(word) = first_word else {
                continue;
            };

            if let Some(pos) = reassembled[last_pos..].find(word) {
                last_pos += pos + word.len();
            } else if let Some(pos) = reassembled.find(word) {
                // Word found but before last_pos → out of order
                if pos < last_pos {
                    out_of_order.push((line_num, pos));
                }
            }
        }

        if !out_of_order.is_empty() {
            let report: String = out_of_order
                .iter()
                .take(50)
                .map(|(num, pos)| {
                    format!("  line {num} found at chunk position {pos} (before previous line)")
                })
                .collect::<Vec<_>>()
                .join("\n");
            panic!(
                "{} lines appear to be recorded out of order:\n{report}",
                out_of_order.len()
            );
        }

        eprintln!(
            "OK: {} non-empty lines processed into {} TTS chunks, all in order, none skipped.",
            input_lines.len(),
            chunks.len()
        );
    }

    #[test]
    fn audiobook_real_pipeline_chunk_count_for_env_file() {
        let path_owned = std::env::var("SONARPAD_TTS_TEST_PATH")
            .unwrap_or_else(|_| r"C:\Users\ambro\Downloads\05-02-2026.txt".to_string());
        let path = path_owned.as_str();
        let raw = match std::fs::read_to_string(path) {
            Ok(s) => s,
            Err(e) => {
                eprintln!("Skipping test – cannot read {path}: {e}");
                return;
            }
        };

        let cleaned = strip_dashed_lines(&raw);
        let prepared = prepare_tts_text(&cleaned, true, &[]);
        let chunks = split_text_for_engine(&prepared, TtsEngine::Edge);

        assert!(
            !chunks.is_empty(),
            "real audiobook pipeline produced 0 chunks"
        );

        let first_preview = preview_for_log(&chunks[0], 140);
        let last_preview = preview_for_log(&chunks[chunks.len() - 1], 140);
        eprintln!(
            "REAL PIPELINE: cleaned_chars={} prepared_chars={} chunks={} first=\"{}\" last=\"{}\"",
            cleaned.chars().count(),
            prepared.chars().count(),
            chunks.len(),
            first_preview,
            last_preview
        );
    }
}

fn merge_and_finalize_sapi4_audio(
    wav_files: &[PathBuf],
    output: &Path,
    language: Language,
    options: &AudiobookCommonOptions,
) -> Result<(), String> {
    if wav_files.is_empty() {
        return Ok(());
    }

    let extension = output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_mp3 = extension == "mp3";
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
    let final_wav = if is_mp3 || is_aac {
        output.with_extension("wav.tmp")
    } else {
        output.to_path_buf()
    };

    crate::audio_utils::join_wav_files(wav_files, &final_wav)
        .map_err(|e| format!("Failed to join audio parts: {}", e))?;

    if is_mp3 || is_aac {
        let res = if is_aac {
            let settings = crate::ffmpeg_export::ConvertAudioSettings {
                format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
                quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                    options.audiobook_bitrate_kbps,
                ),
            };
            let mut progress = |_p: u32| {};
            crate::ffmpeg_export::convert_audio_file(
                &final_wav,
                output,
                &settings,
                None,
                Some(&mut progress),
            )
        } else {
            let settings = crate::ffmpeg_export::ConvertAudioSettings {
                format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
                quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                    options.audiobook_bitrate_kbps,
                ),
            };
            let mut progress = |_p: u32| {};
            crate::ffmpeg_export::convert_audio_file(
                &final_wav,
                output,
                &settings,
                None,
                Some(&mut progress),
            )
        };
        std::fs::remove_file(&final_wav).ok();
        if let Err(e) = res {
            return Err(i18n::tr_f(language, "sapi5.mf_error", &[("err", &e)]));
        }
    }
    Ok(())
}

fn run_sapi5_parallel_part(
    chunks: &[String],
    global_progress: &mut usize,
    options: &AudiobookCommonOptions,
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }

    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    if extension != "mp3" {
        return Err("SAPI5 parallel mode supports MP3 output only".to_string());
    }

    let temp_dir = options
        .output
        .parent()
        .unwrap_or(Path::new("."))
        .join(format!(
            "sapi5_tmp_{}",
            options
                .output
                .file_name()
                .and_then(|s| s.to_str())
                .unwrap_or("part")
        ));
    std::fs::create_dir_all(&temp_dir).map_err(|e| e.to_string())?;

    let chunks_count = chunks.len();
    let worker_count = chunks_count.min(
        std::thread::available_parallelism()
            .map(|n| n.get())
            .unwrap_or(4)
            .clamp(2, 12),
    );
    let chunks_per_sub = chunks_count.div_ceil(worker_count.max(1));

    // Important: we split by chunk boundaries only; chunks are already sentence-safe upstream.
    let mut internal_parts: Vec<(Vec<String>, PathBuf)> = Vec::new();
    for i in 0..worker_count {
        let s = i * chunks_per_sub;
        let e = std::cmp::min(s + chunks_per_sub, chunks_count);
        if s >= e {
            break;
        }
        let sub_chunks = chunks[s..e].to_vec();
        let sub_output = temp_dir.join(format!("sub_{}.mp3", i));
        internal_parts.push((sub_chunks, sub_output));
    }
    let expected_parts = internal_parts.len();

    let progress_counter = Arc::new(std::sync::atomic::AtomicUsize::new(*global_progress));
    let (tx, rx) = std::sync::mpsc::channel::<Result<PathBuf, String>>();
    let parts_shared = Arc::new(Mutex::new(internal_parts.into_iter()));
    let mut handles = Vec::new();

    for _worker_idx in 0..worker_count {
        let tx = tx.clone();
        let parts_shared = parts_shared.clone();
        let progress_counter = progress_counter.clone();
        let cancel_token = options.cancel.clone();
        let progress_hwnd = options.progress_hwnd;
        let voice = options.voice.to_string();
        let language = options.language;
        let rate = options.rate;
        let pitch = options.pitch;
        let volume = options.volume;
        let bitrate = options.audiobook_bitrate_kbps;

        let handle = std::thread::spawn(move || {
            loop {
                let part = {
                    let mut guard = parts_shared.lock().unwrap_or_else(|e| e.into_inner());
                    guard.next()
                };
                let Some((sub_chunks, sub_output)) = part else {
                    break;
                };

                if cancel_token.load(Ordering::Relaxed) {
                    let _ignored = tx.send(Err("Cancelled".to_string()));
                    break;
                }

                let res = crate::sapi5_engine::speak_sapi_to_file(
                    crate::sapi5_engine::SapiExportOptions {
                        chunks: &sub_chunks,
                        voice_name: &voice,
                        output_path: &sub_output,
                        language,
                        rate,
                        pitch,
                        volume,
                        audiobook_bitrate_kbps: bitrate,
                        cancel: cancel_token.clone(),
                    },
                    |_| {
                        let current = progress_counter.fetch_add(1, Ordering::SeqCst) + 1;
                        if progress_hwnd.0 != 0 {
                            unsafe {
                                if let Err(e) = PostMessageW(
                                    progress_hwnd,
                                    crate::WM_UPDATE_PROGRESS,
                                    WPARAM(current),
                                    LPARAM(0),
                                ) {
                                    crate::log_debug(&format!(
                                        "Failed to post WM_UPDATE_PROGRESS: {}",
                                        e
                                    ));
                                }
                            }
                        }
                    },
                );

                match res {
                    Ok(()) => {
                        let _ignored = tx.send(Ok(sub_output));
                    }
                    Err(e) => {
                        if let Err(rem_err) = std::fs::remove_file(&sub_output) {
                            crate::log_debug(&format!(
                                "Failed to remove SAPI5 subpart after error: {}",
                                rem_err
                            ));
                        }
                        let _ignored = tx.send(Err(e));
                        break;
                    }
                }
            }
        });
        handles.push(handle);
    }

    {
        let _tx = tx;
    }

    let mut produced_files: Vec<PathBuf> = Vec::new();
    let mut error: Option<String> = None;
    for res in rx {
        match res {
            Ok(path) => produced_files.push(path),
            Err(e) => {
                if error.is_none() {
                    error = Some(e);
                }
            }
        }
    }
    for h in handles {
        if let Err(e) = h.join() {
            crate::log_debug(&format!("SAPI5 parallel worker join error: {:?}", e));
        }
    }

    if error.is_some() || options.cancel.load(Ordering::Relaxed) {
        if let Err(e) = std::fs::remove_dir_all(&temp_dir) {
            crate::log_debug(&format!("Failed to remove SAPI5 temp dir: {}", e));
        }
        let msg = if let Some(e) = error {
            e
        } else {
            "Cancelled".to_string()
        };
        return Err(if msg == "Cancelled" {
            cancelled_message(options.language)
        } else {
            msg
        });
    }

    if produced_files.len() != expected_parts {
        if let Err(e) = std::fs::remove_dir_all(&temp_dir) {
            crate::log_debug(&format!("Failed to remove SAPI5 temp dir: {}", e));
        }
        return Err(format!(
            "SAPI5 parallel integrity check failed: expected {} parts, got {}",
            expected_parts,
            produced_files.len()
        ));
    }

    produced_files.sort_by_key(|p: &PathBuf| {
        let name = p.file_stem().and_then(|s| s.to_str()).unwrap_or("");
        parse_sapi4_part_index(name)
    });

    for file in &produced_files {
        let size = std::fs::metadata(file).map_err(|e| e.to_string())?.len();
        if size == 0 {
            if let Err(e) = std::fs::remove_dir_all(&temp_dir) {
                crate::log_debug(&format!("Failed to remove SAPI5 temp dir: {}", e));
            }
            return Err(format!(
                "SAPI5 parallel integrity check failed: empty chunk output {:?}",
                file
            ));
        }
    }

    merge_and_finalize_sapi4_mp3(
        &produced_files,
        options.output,
        options.language,
        options.audiobook_bitrate_kbps,
        options.progress_hwnd,
    )?;
    if let Err(e) = std::fs::remove_dir_all(&temp_dir) {
        crate::log_debug(&format!("Failed to remove SAPI5 temp dir: {}", e));
    }
    *global_progress = progress_counter.load(Ordering::SeqCst);
    Ok(())
}

fn run_split_sapi_audiobook(
    chunks: &[String],
    split_parts: u32,
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let parts = if split_parts == 0 {
        1
    } else {
        split_parts as usize
    };
    let total_chunks = chunks.len();
    let parts = if total_chunks < parts {
        total_chunks
    } else {
        parts
    };
    let chunks_per_part = total_chunks.div_ceil(parts);
    let mut current_global_progress = 0;

    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }

        let part_chunks = &chunks[start_idx..end_idx];

        let part_output = if parts > 1 {
            split_part_output(options.output, part_idx, parts, options.part_naming_mode)
        } else {
            options.output.to_path_buf()
        };

        let extension = part_output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
        if extension == "mp3" {
            let part_options = AudiobookCommonOptions {
                voice: options.voice,
                output: &part_output,
                progress_hwnd: options.progress_hwnd,
                cancel: options.cancel.clone(),
                language: options.language,
                part_naming_mode: options.part_naming_mode,
                audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
                sapi4_threads: options.sapi4_threads,
            };
            run_sapi5_parallel_part(part_chunks, &mut current_global_progress, &part_options)?;
            continue;
        }
        let actual_output = if is_aac {
            part_output.with_extension("wav.tmp")
        } else {
            part_output.clone()
        };

        let progress_hwnd_clone = options.progress_hwnd;
        let cancel_clone = options.cancel.clone();

        crate::sapi5_engine::speak_sapi_to_file(
            crate::sapi5_engine::SapiExportOptions {
                chunks: part_chunks,
                voice_name: options.voice,
                output_path: &actual_output,
                language: options.language,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
                audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
                cancel: cancel_clone,
            },
            |_chunk_idx| {
                current_global_progress += 1;
                if progress_hwnd_clone.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            progress_hwnd_clone,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
            },
        )
        .map_err(|e| {
            if let Err(rem_err) = std::fs::remove_file(&actual_output) {
                crate::log_debug(&format!(
                    "Failed to remove part output after error {}: {}",
                    e, rem_err
                ));
            }
            e
        })?;

        if is_aac {
            let settings = crate::ffmpeg_export::ConvertAudioSettings {
                format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
                quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                    options.audiobook_bitrate_kbps,
                ),
            };
            let mut progress = |_p: u32| {};
            let res = crate::ffmpeg_export::convert_audio_file(
                &actual_output,
                &part_output,
                &settings,
                None,
                Some(&mut progress),
            );
            std::fs::remove_file(&actual_output).ok();
            res?;
        }
    }
    Ok(())
}

fn run_marker_split_sapi_audiobook(
    parts: &[Vec<String>],
    options: AudiobookCommonOptions,
) -> Result<(), String> {
    let parts_len = parts.len();
    let mut current_global_progress = 0;

    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if part_chunks.is_empty() {
            continue;
        }
        let part_output = if parts_len > 1 {
            split_part_output(
                options.output,
                part_idx,
                parts_len,
                options.part_naming_mode,
            )
        } else {
            options.output.to_path_buf()
        };

        let extension = part_output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
        if extension == "mp3" {
            let part_options = AudiobookCommonOptions {
                voice: options.voice,
                output: &part_output,
                progress_hwnd: options.progress_hwnd,
                cancel: options.cancel.clone(),
                language: options.language,
                part_naming_mode: options.part_naming_mode,
                audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
                sapi4_threads: options.sapi4_threads,
            };
            run_sapi5_parallel_part(part_chunks, &mut current_global_progress, &part_options)?;
            continue;
        }
        let actual_output = if is_aac {
            part_output.with_extension("wav.tmp")
        } else {
            part_output.clone()
        };

        let progress_hwnd_clone = options.progress_hwnd;
        let cancel_clone = options.cancel.clone();

        crate::sapi5_engine::speak_sapi_to_file(
            crate::sapi5_engine::SapiExportOptions {
                chunks: part_chunks,
                voice_name: options.voice,
                output_path: &actual_output,
                language: options.language,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
                audiobook_bitrate_kbps: options.audiobook_bitrate_kbps,
                cancel: cancel_clone,
            },
            |_chunk_idx| {
                current_global_progress += 1;
                if progress_hwnd_clone.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            progress_hwnd_clone,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
            },
        )
        .map_err(|e| {
            if let Err(rem_err) = std::fs::remove_file(&actual_output) {
                crate::log_debug(&format!(
                    "Failed to remove part output after error {}: {}",
                    e, rem_err
                ));
            }
            e
        })?;

        if is_aac {
            let settings = crate::ffmpeg_export::ConvertAudioSettings {
                format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
                quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                    options.audiobook_bitrate_kbps,
                ),
            };
            let mut progress = |_p: u32| {};
            let res = crate::ffmpeg_export::convert_audio_file(
                &actual_output,
                &part_output,
                &settings,
                None,
                Some(&mut progress),
            );
            std::fs::remove_file(&actual_output).ok();
            res?;
        }
    }
    Ok(())
}

pub(crate) fn run_tts_audiobook_part(
    chunks: &[String],
    current_global_progress: &mut usize,
    options: &AudiobookCommonOptions,
) -> Result<(), String> {
    let is_lithuanian_voice = options.voice.to_ascii_lowercase().starts_with("lt-");
    const LT_EDGE_MAX_BYTES: usize = 900;

    let chunk_texts: Vec<String> = {
        let mut out = Vec::new();
        for chunk in chunks {
            // Force strict sanitization for audiobook edge chunks by default.
            // Safe profile remains available for non-audiobook paths.
            let normalized =
                normalize_for_tts_with_profile(chunk, true, TtsSanitizeProfile::Strict);
            let trimmed = normalized.trim();
            if trimmed.is_empty() {
                continue;
            }
            if is_lithuanian_voice && normalized.len() > LT_EDGE_MAX_BYTES {
                out.extend(
                    split_text_edge_with_limit(&normalized, LT_EDGE_MAX_BYTES)
                        .into_iter()
                        .filter(|s| !s.trim().is_empty()),
                );
            } else {
                out.push(normalized);
            }
        }
        out
    };

    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|err| err.to_string())?;

    rt.block_on(async {
        if options.cancel.load(Ordering::Relaxed) {
            return Err(cancelled_message(options.language));
        }

        if chunk_texts.is_empty() {
            return Ok(());
        }

        let edge_chunks: Vec<TtsChunk> = chunk_texts
            .iter()
            .map(|chunk| TtsChunk {
                text_to_read: chunk.clone(),
                original_len: utf16_len(chunk),
                override_voice: None,
            })
            .collect();

        const WS_RETRY_MAX: usize = 3;
        let mut next_index = 0usize;

        let file = std::fs::File::create(options.output).map_err(|err| err.to_string())?;
        let mut writer = BufWriter::new(file);

        let stream_options = EdgeStreamOptions {
            voice: options.voice,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            language: options.language,
            cancel: options.cancel.as_ref(),
            progress_hwnd: options.progress_hwnd,
            allow_http_fallback: true,
        };

        let mut attempt = 1usize;
        let last_err = loop {
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }

            let base_progress = *current_global_progress;
            let res = download_edge_chunks_ws_parallel_to_writer(
                &edge_chunks,
                next_index,
                &stream_options,
                &mut writer,
                current_global_progress,
            )
            .await;

            match res {
                Ok(_) => break None,
                Err(err) => {
                    let completed = current_global_progress.saturating_sub(base_progress);
                    next_index = (next_index + completed).min(edge_chunks.len());
                    let err_str = err.as_str();
                    let retry_forever = is_edge_retry_forever_error(err_str);
                    crate::log_debug(&format!(
                        "Edge WS export failed (attempt {}/{}): {}",
                        attempt,
                        if retry_forever {
                            "inf".to_string()
                        } else {
                            WS_RETRY_MAX.to_string()
                        },
                        err_str
                    ));
                    if next_index >= edge_chunks.len() {
                        break Some(err);
                    }
                    if !retry_forever && attempt >= WS_RETRY_MAX {
                        break Some(err);
                    }
                    std::thread::sleep(std::time::Duration::from_millis(edge_retry_delay_ms(
                        err_str, attempt,
                    )));
                    attempt = attempt.saturating_add(1);
                }
            }
        };
        if let Some(err) = last_err {
            return Err(err);
        }
        Ok(())
    })?;

    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_mp3 = extension == "mp3";
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";

    if is_aac {
        set_audiobook_progress_phase(options.progress_hwnd, true);
        set_audiobook_progress_total(options.progress_hwnd, 200);
        post_audiobook_progress(options.progress_hwnd, 0);
        let source_mp3 = options.output.with_extension("edge_source.tmp.mp3");
        let source_wav = options.output.with_extension("edge_source.tmp.wav");
        if let Err(e) = std::fs::rename(options.output, &source_mp3) {
            return Err(format!("Failed to prepare AAC conversion: {}", e));
        }

        let wav_result = (|| -> Result<(), String> {
            let bytes = std::fs::read(&source_mp3)
                .map_err(|e| format!("Failed to read source MP3 for AAC conversion: {}", e))?;
            let (samples, src_rate, src_channels) = decode_mp3_to_pcm(&bytes)?;
            let target_rate = 48_000u32;
            let target_channels = 2u16;
            let resampled = resample_pcm(
                &samples,
                src_rate,
                src_channels,
                target_rate,
                target_channels,
            );
            write_wav_from_pcm(&source_wav, &resampled, target_rate, target_channels)
        })();
        post_audiobook_progress(options.progress_hwnd, 100);
        if let Err(err) = wav_result {
            crate::log_debug(&format!(
                "AAC post-process skipped: WAV conversion failed: {}",
                err
            ));
            if let Err(restore_err) = std::fs::rename(&source_mp3, options.output) {
                crate::log_debug(&format!(
                    "Failed to restore source MP3 after AAC conversion failure: {}",
                    restore_err
                ));
                return Err(format!(
                    "Failed to restore source MP3 after AAC conversion failure: {}",
                    restore_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_wav) {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after AAC conversion failure: {}",
                    remove_err
                ));
            }
            return Ok(());
        }

        let aac_settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut ffmpeg_progress = |p: u32| {
            post_finalization_progress_range(options.progress_hwnd, p, 100, 99, 200);
        };
        let convert_result = crate::ffmpeg_export::convert_audio_file(
            &source_wav,
            options.output,
            &aac_settings,
            None,
            Some(&mut ffmpeg_progress),
        );
        post_audiobook_progress(options.progress_hwnd, 199);
        if let Err(err) = convert_result {
            crate::log_if_err!(std::fs::remove_file(options.output));
            if let Err(restore_err) = std::fs::rename(&source_mp3, options.output) {
                crate::log_debug(&format!(
                    "Failed to restore source MP3 after AAC re-encode failure: {}",
                    restore_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_wav) {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after AAC re-encode failure: {}",
                    remove_err
                ));
            }
            return Err(err);
        }
        if !KEEP_EDGE_TEMP_AFTER_CONVERSION {
            if let Err(remove_err) = std::fs::remove_file(&source_wav) {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after AAC re-encode: {}",
                    remove_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_mp3) {
                crate::log_debug(&format!(
                    "Failed to remove temp source MP3 after AAC re-encode: {}",
                    remove_err
                ));
            }
        } else {
            crate::log_debug("Keeping AAC edge temp files (debug mode)");
        }
        post_audiobook_progress(options.progress_hwnd, 200);
    } else if is_mp3 {
        set_audiobook_progress_phase(options.progress_hwnd, true);
        set_audiobook_progress_total(options.progress_hwnd, 200);
        post_audiobook_progress(options.progress_hwnd, 0);
        let source_mp3 = options.output.with_extension("edge_source.tmp.mp3");
        let source_wav = options.output.with_extension("edge_source.tmp.wav");
        if let Err(e) = std::fs::rename(options.output, &source_mp3) {
            return Err(format!("Failed to prepare MP3 conversion: {}", e));
        }

        let wav_settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Wav,
            quality: crate::ffmpeg_export::ConvertAudioQuality::None,
        };
        const MP3_TO_WAV_MAX_ATTEMPTS: usize = 12;
        let mut wav_result = Err("MP3->WAV conversion not attempted".to_string());
        for attempt in 1..=MP3_TO_WAV_MAX_ATTEMPTS {
            let mut wav_progress = |p: u32| {
                post_finalization_progress_range(options.progress_hwnd, p, 0, 99, 200);
            };
            wav_result = crate::ffmpeg_export::convert_audio_file_with_channels(
                &source_mp3,
                &source_wav,
                &wav_settings,
                None,
                Some(&mut wav_progress),
                Some(2),
            );
            if wav_result.is_ok() {
                break;
            }
            if let Err(remove_err) = std::fs::remove_file(&source_wav)
                && remove_err.kind() != std::io::ErrorKind::NotFound
            {
                crate::log_debug(&format!(
                    "Failed to remove partial WAV output after MP3->WAV attempt {}: {}",
                    attempt, remove_err
                ));
            }
            if attempt < MP3_TO_WAV_MAX_ATTEMPTS {
                if let Err(err) = &wav_result {
                    crate::log_debug(&format!(
                        "MP3->WAV attempt {}/{} failed: {}. Retrying...",
                        attempt, MP3_TO_WAV_MAX_ATTEMPTS, err
                    ));
                }
                std::thread::sleep(std::time::Duration::from_millis(300));
            }
        }
        post_audiobook_progress(options.progress_hwnd, 100);
        if let Err(err) = wav_result {
            crate::log_if_err!(std::fs::remove_file(options.output));
            if let Err(restore_err) = std::fs::rename(&source_mp3, options.output) {
                crate::log_debug(&format!(
                    "Failed to restore original MP3 after WAV conversion failure: {}",
                    restore_err
                ));
                return Err(format!(
                    "MP3->WAV failed after {} attempts ({}) and restore failed: {}",
                    MP3_TO_WAV_MAX_ATTEMPTS, err, restore_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_wav)
                && remove_err.kind() != std::io::ErrorKind::NotFound
            {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after conversion failure: {}",
                    remove_err
                ));
            }
            return Err(format!(
                "MP3->WAV failed after {} attempts: {}",
                MP3_TO_WAV_MAX_ATTEMPTS, err
            ));
        }

        let mp3_settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut last_convert_pct = u32::MAX;
        let mut ffmpeg_progress = |p: u32| {
            let pct = (p / 100).min(100);
            if pct != last_convert_pct {
                last_convert_pct = pct;
                post_finalization_progress_range(options.progress_hwnd, p, 100, 99, 200);
            }
        };
        const MP3_REENCODE_MAX_ATTEMPTS: usize = 12;
        let mut reencode_result = Err("MP3 re-encode not attempted".to_string());
        for attempt in 1..=MP3_REENCODE_MAX_ATTEMPTS {
            reencode_result = crate::ffmpeg_export::convert_audio_file(
                &source_wav,
                options.output,
                &mp3_settings,
                None,
                Some(&mut ffmpeg_progress),
            );
            if reencode_result.is_ok() {
                break;
            }

            if let Err(remove_err) = std::fs::remove_file(options.output)
                && remove_err.kind() != std::io::ErrorKind::NotFound
            {
                crate::log_debug(&format!(
                    "Failed to remove partial MP3 output after re-encode attempt {}: {}",
                    attempt, remove_err
                ));
            }

            if attempt < MP3_REENCODE_MAX_ATTEMPTS {
                if let Err(err) = &reencode_result {
                    crate::log_debug(&format!(
                        "MP3 re-encode attempt {}/{} failed: {}. Retrying...",
                        attempt, MP3_REENCODE_MAX_ATTEMPTS, err
                    ));
                }
                std::thread::sleep(std::time::Duration::from_millis(300));
            }
        }
        post_audiobook_progress(options.progress_hwnd, 199);
        if let Err(err) = reencode_result {
            crate::log_if_err!(std::fs::remove_file(options.output));
            if let Err(restore_err) = std::fs::rename(&source_mp3, options.output) {
                crate::log_debug(&format!(
                    "Failed to restore original MP3 after re-encode failure: {}",
                    restore_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_wav)
                && remove_err.kind() != std::io::ErrorKind::NotFound
            {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after re-encode failure: {}",
                    remove_err
                ));
            }
            return Err(format!(
                "MP3 re-encode failed after {} attempts: {}",
                MP3_REENCODE_MAX_ATTEMPTS, err
            ));
        }
        if !KEEP_EDGE_TEMP_AFTER_CONVERSION {
            if let Err(remove_err) = std::fs::remove_file(&source_wav)
                && remove_err.kind() != std::io::ErrorKind::NotFound
            {
                crate::log_debug(&format!(
                    "Failed to remove temp WAV after MP3 re-encode: {}",
                    remove_err
                ));
            }
            if let Err(remove_err) = std::fs::remove_file(&source_mp3) {
                crate::log_debug(&format!(
                    "Failed to remove source MP3 after re-encode: {}",
                    remove_err
                ));
            }
        } else {
            crate::log_debug("Keeping MP3 edge temp files (debug mode)");
        }
        post_audiobook_progress(options.progress_hwnd, 200);
    }

    Ok(())
}

pub(crate) struct MixedAudiobookConfig {
    pub(crate) main_engine: TtsEngine,
}

fn mixed_chunk_has_usable_content(text: &str) -> bool {
    let decoded = decode_basic_xml_entities(text);
    let trimmed = decoded.trim();
    if trimmed.is_empty() {
        return false;
    }
    trimmed.chars().any(|ch| ch.is_alphanumeric())
}

fn resolve_mixed_chunk_synth<'a>(
    chunk: &'a TtsChunk,
    options: &'a AudiobookCommonOptions,
    config: &'a MixedAudiobookConfig,
) -> (TtsEngine, &'a str, i32, i32, i32) {
    if let Some(ov) = &chunk.override_voice {
        (
            ov.engine,
            ov.voice.as_str(),
            ov.rate.unwrap_or(0),
            ov.pitch.unwrap_or(0),
            ov.volume.unwrap_or(100),
        )
    } else {
        (
            config.main_engine,
            options.voice,
            options.rate,
            options.pitch,
            options.volume,
        )
    }
}

async fn synthesize_mixed_chunk_with_retry(
    chunk_idx: usize,
    chunk: &TtsChunk,
    synth: &SynthesisConfig,
    target: TargetAudio,
) -> Option<PathBuf> {
    const MIXED_SEGMENT_MAX_ATTEMPTS: usize = 5;

    crate::log_debug(&format!(
        "Mixed audiobook synth: chunk={} engine={} voice={:?} rate={} pitch={} volume={} override={} text_preview={:?}",
        chunk_idx,
        match synth.engine {
            TtsEngine::Edge => "edge",
            TtsEngine::Sapi5 => "sapi5",
            TtsEngine::Sapi4 => "sapi4",
        },
        synth.voice,
        synth.rate,
        synth.pitch,
        synth.volume,
        chunk.override_voice.is_some(),
        preview_for_log(&chunk.text_to_read, 120)
    ));

    if !mixed_chunk_has_usable_content(&chunk.text_to_read) {
        crate::log_debug(&format!(
            "Mixed audiobook: dropping non-usable chunk={} engine={} voice={:?} text_preview={:?}",
            chunk_idx,
            match synth.engine {
                TtsEngine::Edge => "edge",
                TtsEngine::Sapi5 => "sapi5",
                TtsEngine::Sapi4 => "sapi4",
            },
            synth.voice,
            preview_for_log(&chunk.text_to_read, 120)
        ));
        return None;
    }

    let mut last_err: Option<String> = None;
    for attempt in 1..=MIXED_SEGMENT_MAX_ATTEMPTS {
        match synthesize_segment_to_wav(&chunk.text_to_read, synth, target).await {
            Ok(wav_path) => return Some(wav_path),
            Err(err) => {
                crate::log_debug(&format!(
                    "Mixed audiobook: synth attempt {}/{} failed for chunk={} engine={} voice={:?}: {}",
                    attempt,
                    MIXED_SEGMENT_MAX_ATTEMPTS,
                    chunk_idx,
                    match synth.engine {
                        TtsEngine::Edge => "edge",
                        TtsEngine::Sapi5 => "sapi5",
                        TtsEngine::Sapi4 => "sapi4",
                    },
                    synth.voice,
                    err
                ));
                last_err = Some(err);
                if attempt < MIXED_SEGMENT_MAX_ATTEMPTS {
                    tokio::time::sleep(Duration::from_millis(250 * attempt as u64)).await;
                }
            }
        }
    }

    crate::log_debug(&format!(
        "Mixed audiobook: dropping chunk after retries chunk={} engine={} voice={:?} text_preview={:?} last_error={}",
        chunk_idx,
        match synth.engine {
            TtsEngine::Edge => "edge",
            TtsEngine::Sapi5 => "sapi5",
            TtsEngine::Sapi4 => "sapi4",
        },
        synth.voice,
        preview_for_log(&chunk.text_to_read, 120),
        last_err.unwrap_or_else(|| "unknown error".to_string())
    ));
    None
}

pub(crate) fn render_mixed_audiobook_part(
    chunks: &[TtsChunk],
    current_global_progress: &mut usize,
    output: &Path,
    options: &AudiobookCommonOptions,
    config: &MixedAudiobookConfig,
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }
    let rt = tokio::runtime::Runtime::new().map_err(|e| e.to_string())?;
    let extension = output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
    let is_mp3 = extension == "mp3";
    let actual_output = if is_aac || is_mp3 {
        output.with_extension("wav.tmp")
    } else {
        output.to_path_buf()
    };

    let target = TargetAudio {
        sample_rate: 44_100,
        channels: 1,
    };
    let mut temp_wavs: Vec<PathBuf> = Vec::new();
    let edge_only = chunks.iter().all(|chunk| {
        let (engine, _, _, _, _) = resolve_mixed_chunk_synth(chunk, options, config);
        engine == TtsEngine::Edge
    });

    if edge_only {
        const MIXED_EDGE_PARALLELISM: usize = 8;
        crate::log_debug(&format!(
            "Mixed audiobook: edge-only mode, parallel synthesis enabled with concurrency={}",
            MIXED_EDGE_PARALLELISM
        ));
        for batch_start in (0..chunks.len()).step_by(MIXED_EDGE_PARALLELISM) {
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }
            let batch_end = std::cmp::min(batch_start + MIXED_EDGE_PARALLELISM, chunks.len());
            let batch = &chunks[batch_start..batch_end];
            let results = rt.block_on(async {
                let futures = batch.iter().enumerate().map(|(offset, chunk)| {
                    let idx = batch_start + offset;
                    let (engine, voice, rate, pitch, volume) =
                        resolve_mixed_chunk_synth(chunk, options, config);
                    let synth = SynthesisConfig {
                        engine,
                        voice: voice.to_string(),
                        rate,
                        pitch,
                        volume,
                        language: options.language,
                    };
                    async move {
                        (
                            idx,
                            synthesize_mixed_chunk_with_retry(idx, chunk, &synth, target).await,
                        )
                    }
                });
                join_all(futures).await
            });

            for (_idx, wav) in results {
                if let Some(wav_path) = wav {
                    temp_wavs.push(wav_path);
                }
                *current_global_progress += 1;
                if options.progress_hwnd.0 != 0 {
                    unsafe {
                        if let Err(e) = PostMessageW(
                            options.progress_hwnd,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(*current_global_progress),
                            LPARAM(0),
                        ) {
                            crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                        }
                    }
                }
            }
        }
    } else {
        for (chunk_idx, chunk) in chunks.iter().enumerate() {
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }
            let (engine, voice, rate, pitch, volume) =
                resolve_mixed_chunk_synth(chunk, options, config);
            let synth = SynthesisConfig {
                engine,
                voice: voice.to_string(),
                rate,
                pitch,
                volume,
                language: options.language,
            };
            let wav = rt.block_on(synthesize_mixed_chunk_with_retry(
                chunk_idx, chunk, &synth, target,
            ));
            if let Some(wav_path) = wav {
                temp_wavs.push(wav_path);
            }
            *current_global_progress += 1;
            if options.progress_hwnd.0 != 0 {
                unsafe {
                    if let Err(e) = PostMessageW(
                        options.progress_hwnd,
                        crate::WM_UPDATE_PROGRESS,
                        WPARAM(*current_global_progress),
                        LPARAM(0),
                    ) {
                        crate::log_debug(&format!("Failed to post WM_UPDATE_PROGRESS: {}", e));
                    }
                }
            }
        }
    }

    if temp_wavs.is_empty() {
        return Err("No valid audio segments were produced.".to_string());
    }
    crate::audio_utils::join_wav_files(&temp_wavs, &actual_output).map_err(|e| e.to_string())?;
    for path in temp_wavs {
        if let Err(e) = std::fs::remove_file(&path) {
            crate::log_debug(&format!("Failed to remove temp wav {:?}: {}", path, e));
        }
    }

    if is_aac {
        let settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Aac,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut progress = |_p: u32| {};
        let res = crate::ffmpeg_export::convert_audio_file(
            &actual_output,
            output,
            &settings,
            None,
            Some(&mut progress),
        );
        std::fs::remove_file(&actual_output).ok();
        res?;
    } else if is_mp3 {
        let settings = crate::ffmpeg_export::ConvertAudioSettings {
            format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
            quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                options.audiobook_bitrate_kbps,
            ),
        };
        let mut progress = |_p: u32| {};
        let res = crate::ffmpeg_export::convert_audio_file(
            &actual_output,
            output,
            &settings,
            None,
            Some(&mut progress),
        );
        std::fs::remove_file(&actual_output).ok();
        res?;
    }

    Ok(())
}
