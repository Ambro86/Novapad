use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings;
use crate::settings::{
    AudiobookResult, DictionaryEntry, Language, TRUSTED_CLIENT_TOKEN, TtsEngine,
};
use crate::{get_active_edit, log_debug, save_audio_dialog, show_error, with_state};
use chrono::Local;
use futures_util::{SinkExt, StreamExt};
use rand::Rng;
use rodio::{Decoder, OutputStreamBuilder, Sink};
use sha2::{Digest, Sha256};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::Duration;
use tokio::sync::mpsc;
use tokio_tungstenite::{
    connect_async, tungstenite, tungstenite::client::IntoClientRequest,
    tungstenite::http::HeaderValue, tungstenite::protocol::Message,
};
use url::Url;
use uuid::Uuid;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Power::{ES_CONTINUOUS, ES_SYSTEM_REQUIRED, SetThreadExecutionState};
use windows::Win32::UI::WindowsAndMessaging::{PostMessageW, SendMessageW, WM_APP};

pub const WSS_URL_BASE: &str =
    "wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1";
pub const MAX_TTS_TEXT_LEN: usize = 3000;
pub const MAX_TTS_TEXT_LEN_LONG: usize = 2000;
pub const MAX_TTS_FIRST_CHUNK_LEN_LONG: usize = 800;
pub const TTS_LONG_TEXT_THRESHOLD: usize = MAX_TTS_TEXT_LEN;

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

/// Quotation mark characters used to identify dialogues.
/// Includes straight quotes, curly quotes, guillemets, and various international styles.
#[allow(dead_code)]
pub const DIALOGUE_QUOTES: &[char] = &[
    '"', '“', '”', '«', '»', '‹', '›', '„', '‚', '\'', '‘', '’', '‟',
];

/// Opening quotes that start a dialogue span.
const OPENING_QUOTES: &[char] = &['"', '“', '«', '‹', '„', '\'', '‘', '‟'];

/// Closing quotes that end a dialogue span.
const CLOSING_QUOTES: &[char] = &['"', '”', '»', '›', '"', '\'', '’', '"'];

#[derive(Clone)]
pub struct TtsChunk {
    pub text_to_read: String,
    pub original_len: usize,
    pub is_dialogue: bool,
}

pub struct TtsPlaybackOptions {
    pub hwnd: HWND,
    pub cleaned: String,
    pub voice: String,
    pub chunks: Vec<TtsChunk>,
    pub initial_caret_pos: i32,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub dialogue_voice: Option<String>,
}

pub struct AudiobookCommonOptions<'a> {
    pub voice: &'a str,
    pub output: &'a Path,
    pub progress_hwnd: HWND,
    pub cancel: Arc<AtomicBool>,
    pub language: Language,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub dialogue_voice: Option<String>,
    pub sapi4_threads: Option<u32>,
}

pub struct DownloadChunkOptions<'a> {
    pub text: &'a str,
    pub voice: &'a str,
    pub request_id: &'a str,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub language: Language,
    pub cancel: &'a AtomicBool,
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
    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return;
    };
    let (
        language,
        split_on_newline,
        tts_engine,
        dictionary,
        tts_rate,
        tts_pitch,
        tts_volume,
        use_dialogue_voice,
        dialogue_voice,
    ) = unsafe {
        with_state(hwnd, |state| {
            (
                state.settings.language,
                state.settings.split_on_newline,
                state.settings.tts_engine,
                state.settings.dictionary.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
                state.settings.use_dialogue_voice,
                state.settings.dialogue_voice.clone(),
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
        false,
        String::new(),
    ));

    let (text, initial_caret_pos) = unsafe { get_text_from_caret(hwnd_edit) };
    if text.trim().is_empty() {
        unsafe {
            show_error(hwnd, language, &settings::tts_no_text_message(language));
        }
        return;
    }
    let voice = unsafe {
        with_state(hwnd, |state| state.settings.tts_voice.clone())
            .unwrap_or_else(|| "it-IT-IsabellaNeural".to_string())
    };
    let chunks = split_into_tts_chunks(&text, split_on_newline, &dictionary);

    // Determine dialogue voice - use it only if enabled and set
    let dialogue_voice_option = if use_dialogue_voice && !dialogue_voice.is_empty() {
        Some(dialogue_voice)
    } else {
        None
    };

    match tts_engine {
        TtsEngine::Edge => start_tts_playback_with_chunks(TtsPlaybackOptions {
            hwnd,
            cleaned: text,
            voice,
            chunks,
            initial_caret_pos,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
            dialogue_voice: dialogue_voice_option,
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
            if unsafe {
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
            if unsafe {
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

pub fn toggle_tts_pause(hwnd: HWND) {
    if unsafe {
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
    prevent_sleep(false);
    if unsafe {
        with_state(hwnd, |state| {
            if let Some(session) = &state.tts_session {
                session.cancel.store(true, Ordering::SeqCst);
                if let Err(e) = session.command_tx.send(TtsCommand::Stop) {
                    crate::log_debug(&format!("Failed to send Stop command: {}", e));
                }
            }
            state.tts_session = None;
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

pub fn start_tts_playback_with_chunks(options: TtsPlaybackOptions) {
    stop_tts_playback(options.hwnd);
    prevent_sleep(true);
    if options.chunks.is_empty() {
        return;
    }

    let language =
        unsafe { with_state(options.hwnd, |state| state.settings.language) }.unwrap_or_default();
    let (tx, mut rx) = mpsc::unbounded_channel::<TtsCommand>();
    let cancel = Arc::new(AtomicBool::new(false));
    let cancel_flag = cancel.clone();
    let session_id = unsafe {
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
    let dialogue_voice = options.dialogue_voice;

    std::thread::spawn(move || {
        log_debug(&format!(
            "TTS start: voice={voice} chunks={} text_len={} dialogue_voice={:?}",
            chunks.len(),
            cleaned.len(),
            dialogue_voice
        ));
        let stream_handle = match OutputStreamBuilder::open_default_stream() {
            Ok(handle) => handle,
            Err(_) => {
                post_tts_error(
                    hwnd_copy,
                    session_id,
                    "Audio output device not available.".to_string(),
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
                return;
            }
        };

        let (audio_tx, mut audio_rx) = mpsc::channel::<Result<(Vec<u8>, usize), String>>(10);
        let cancel_downloader = cancel_flag.clone();
        let chunks_downloader = chunks.clone();
        let voice_downloader = voice.clone();
        let dialogue_voice_downloader = dialogue_voice.clone();
        let rate = tts_rate;
        let pitch = tts_pitch;
        let volume = tts_volume;

        rt.spawn(async move {
            for chunk_obj in chunks_downloader {
                if cancel_downloader.load(Ordering::SeqCst) {
                    break;
                }
                let request_id = Uuid::new_v4().simple().to_string();
                // Use dialogue voice if chunk is dialogue and dialogue voice is set
                let chunk_voice = if chunk_obj.is_dialogue {
                    dialogue_voice_downloader
                        .as_ref()
                        .unwrap_or(&voice_downloader)
                } else {
                    &voice_downloader
                };
                match download_audio_chunk(
                    &chunk_obj.text_to_read,
                    chunk_voice,
                    &request_id,
                    rate,
                    pitch,
                    volume,
                    language,
                )
                .await
                {
                    Ok(data) => {
                        if audio_tx
                            .send(Ok((data, chunk_obj.original_len)))
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

        loop {
            if cancel_flag.load(Ordering::SeqCst) {
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
                break;
            };
            let (audio, orig_len) = match res {
                Ok(data) => data,
                Err(e) => {
                    post_tts_error(hwnd_copy, session_id, e);
                    break;
                }
            };

            if audio.is_empty() {
                continue;
            }

            unsafe {
                if let Err(e) = PostMessageW(
                    hwnd_copy,
                    WM_TTS_CHUNK_START,
                    WPARAM(session_id as usize),
                    LPARAM(current_offset as isize),
                ) {
                    crate::log_debug(&format!("Failed to post WM_TTS_CHUNK_START: {}", e));
                }
            }

            let cursor = std::io::Cursor::new(audio);
            let source = match Decoder::new(cursor) {
                Ok(source) => source,
                Err(_) => {
                    post_tts_error(hwnd_copy, session_id, "Failed to decode audio.".to_string());
                    break;
                }
            };

            sink.append(source);
            appended_any = true;
            while !sink.empty() {
                if cancel_flag.load(Ordering::SeqCst) {
                    sink.stop();
                    break;
                }
                while let Ok(cmd) = rx.try_recv() {
                    if handle_tts_command(cmd, &sink, cancel_flag.as_ref(), &mut paused) {
                        break;
                    }
                }
                if cancel_flag.load(Ordering::SeqCst) {
                    sink.stop();
                    break;
                }
                std::thread::sleep(Duration::from_millis(50));
            }
            if cancel_flag.load(Ordering::SeqCst) {
                break;
            }
            current_offset += orig_len;
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
    });
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

pub async fn download_audio_chunk(
    text: &str,
    voice: &str,
    request_id: &str,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    language: Language,
) -> Result<Vec<u8>, String> {
    let max_retries = 40;
    let mut last_error = String::new();

    for attempt in 1..=max_retries {
        match download_audio_chunk_attempt(text, voice, request_id, tts_rate, tts_pitch, tts_volume)
            .await
        {
            Ok(data) => return Ok(data),
            Err(e) => {
                last_error = e;
                let msg = i18n::tr_f(
                    language,
                    "tts.chunk_download_retry",
                    &[
                        ("attempt", &attempt.to_string()),
                        ("max", &max_retries.to_string()),
                        ("err", &last_error),
                    ],
                );
                log_debug(&msg);
                if attempt < max_retries {
                    tokio::time::sleep(Duration::from_millis(500 * attempt as u64)).await;
                }
            }
        }
    }
    Err(i18n::tr_f(
        language,
        "tts.chunk_download_error",
        &[("err", &last_error)],
    ))
}

async fn wait_or_cancel(duration: Duration, cancel: &AtomicBool) -> bool {
    tokio::select! {
        _ = tokio::time::sleep(duration) => false,
        _ = async {
            while !cancel.load(Ordering::Relaxed) {
                tokio::time::sleep(Duration::from_millis(50)).await;
            }
        } => true,
    }
}

pub async fn download_audio_chunk_cancel(
    options: DownloadChunkOptions<'_>,
) -> Result<Vec<u8>, String> {
    let max_retries = 40;
    let mut last_error = String::new();

    for attempt in 1..=max_retries {
        if options.cancel.load(Ordering::Relaxed) {
            return Err(cancelled_message(options.language));
        }
        match download_audio_chunk_attempt(
            options.text,
            options.voice,
            options.request_id,
            options.rate,
            options.pitch,
            options.volume,
        )
        .await
        {
            Ok(data) => return Ok(data),
            Err(e) => {
                last_error = e;
                let msg = i18n::tr_f(
                    options.language,
                    "tts.chunk_download_retry",
                    &[
                        ("attempt", &attempt.to_string()),
                        ("max", &max_retries.to_string()),
                        ("err", &last_error),
                    ],
                );
                log_debug(&msg);
                if attempt < max_retries
                    && wait_or_cancel(Duration::from_millis(500 * attempt as u64), options.cancel)
                        .await
                {
                    return Err(cancelled_message(options.language));
                }
            }
        }
    }
    Err(i18n::tr_f(
        options.language,
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

    let connect_timeout = Duration::from_secs(10);
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

    let ssml = mkssml(text, voice, tts_rate, tts_pitch, tts_volume);
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
            Message::Text(text) => {
                if text.contains("Path:turn.end") {
                    break;
                }
            }
            Message::Binary(data) => {
                if data.len() < 2 {
                    continue;
                }
                let be_len = u16::from_be_bytes([data[0], data[1]]) as usize;
                let le_len = u16::from_le_bytes([data[0], data[1]]) as usize;
                let mut parsed = false;
                for header_len in [be_len, le_len] {
                    if header_len == 0 || data.len() < header_len + 2 {
                        continue;
                    }
                    let headers_bytes = &data[2..2 + header_len];
                    let headers_str = String::from_utf8_lossy(headers_bytes);
                    if headers_str.contains("Path:audio") {
                        audio_data.extend_from_slice(&data[2 + header_len..]);
                        parsed = true;
                        break;
                    }
                }
                if parsed {
                    continue;
                }
            }
            Message::Close(_) => break,
            _ => {}
        }
    }
    Ok(audio_data)
}

unsafe fn get_text_from_caret(hwnd_edit: HWND) -> (String, i32) {
    let mut start: i32 = 0;
    let mut end: i32 = 0;
    SendMessageW(
        hwnd_edit,
        crate::accessibility::EM_GETSEL,
        WPARAM(&mut start as *mut _ as usize),
        LPARAM(&mut end as *mut _ as isize),
    );
    let caret_pos = start.min(end).max(0) as usize;
    let full_text = get_edit_text(hwnd_edit);

    // Se siamo all'inizio, o se siamo alla fine del testo, leggi tutto dall'inizio
    if caret_pos == 0 {
        return (full_text, 0);
    }

    let normalized = full_text.replace("\r\n", "\n");
    let wide: Vec<u16> = normalized.encode_utf16().collect();

    // Se la posizione del cursore Š oltre la lunghezza del testo (fine file),
    // ricomincia a leggere dall'inizio come richiesto.
    if caret_pos >= wide.len() {
        return (full_text, 0);
    }

    let adjusted_pos = adjust_tts_caret_pos(&normalized, caret_pos as i32);
    let adjusted_pos = adjusted_pos.max(0) as usize;
    if adjusted_pos >= wide.len() {
        return (full_text, 0);
    }
    (
        String::from_utf16_lossy(&wide[adjusted_pos..]),
        adjusted_pos as i32,
    )
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
    format!(
        "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='{}'><voice name='{}'><prosody pitch='{}' rate='{}' volume='{}'>{}</prosody></voice></speak>",
        lang,
        voice,
        format_pitch(tts_pitch),
        format_rate(tts_rate),
        format_volume(tts_volume),
        text
    )
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

pub fn normalize_for_tts(text: &str, split_on_newline: bool) -> String {
    let normalized = if split_on_newline {
        text.to_string()
    } else {
        text.replace('\n', " ").replace('\r', "")
    };
    normalized.replace(['«', '»'], "")
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
        out = out.replace(&entry.original, &entry.replacement);
    }
    out
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

pub(crate) fn build_audiobook_parts_by_positions(
    text: &str,
    positions: &[usize],
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
) -> Option<Vec<Vec<String>>> {
    let parts_text = split_text_by_positions(text, positions)?;
    let mut parts_chunks = Vec::new();

    for part_text in parts_text {
        let prepared = prepare_tts_text(&part_text, split_on_newline, dictionary);
        let chunks = split_text(&prepared);
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

pub fn split_text(text: &str) -> Vec<String> {
    let mut chunks = Vec::new();
    // Trim input to avoid processing leading/trailing whitespace as separate empty chunks
    let text = text.trim();
    if text.is_empty() {
        return chunks;
    }

    let char_indices: Vec<(usize, char)> = text.char_indices().collect();
    let char_len = char_indices.len();
    let mut current_char = 0;

    let is_long = char_len > TTS_LONG_TEXT_THRESHOLD;
    let max_len = if is_long {
        MAX_TTS_TEXT_LEN_LONG
    } else {
        MAX_TTS_TEXT_LEN
    };
    // First chunk can be smaller if text is very long to start playback faster?
    // But for recording, consistency is better.
    let target_chunk_len = if is_long {
        MAX_TTS_FIRST_CHUNK_LEN_LONG
    } else {
        max_len
    };

    let byte_index_at = |char_idx: usize| -> usize {
        if char_idx >= char_len {
            text.len()
        } else {
            char_indices[char_idx].0
        }
    };

    while current_char < char_len {
        // Calculate potential end
        let effective_target = if chunks.is_empty() {
            target_chunk_len
        } else {
            max_len
        };

        let mut split_char = current_char + effective_target;

        if split_char >= char_len {
            // Take the rest
            let chunk = text[byte_index_at(current_char)..].trim().to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            break;
        }

        // Search backwards for a good split point
        let search_end = split_char;
        let search_start = current_char; // Don't go back past start
        let mut split_found = None;

        // Priority 1: Sentence endings
        for idx in (search_start..search_end).rev() {
            let c = char_indices[idx].1;
            if matches!(c, '.' | '!' | '?') {
                let next_idx = idx + 1;
                // Split if next char is end of string or whitespace
                if next_idx >= char_len || char_indices[next_idx].1.is_whitespace() {
                    split_found = Some(next_idx);
                    break;
                }
            }
        }

        // Priority 2: Newlines or major punctuation
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() {
                let c = char_indices[idx].1;
                if c == '\n' {
                    // Try to keep newlines together?
                    if idx + 1 < char_len && char_indices[idx + 1].1 == '\n' {
                        split_found = Some(idx + 2);
                        break;
                    } else {
                        split_found = Some(idx + 1);
                        break;
                    }
                } else if matches!(c, ';' | ':') {
                    split_found = Some(idx + 1);
                    break;
                }
            }
        }

        // Priority 3: Spaces
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() {
                if char_indices[idx].1.is_whitespace() {
                    split_found = Some(idx + 1);
                    break;
                }
            }
        }

        // Apply split
        if let Some(split_at) = split_found {
            // Ensure we advance
            if split_at > current_char {
                split_char = split_at;
            }
        }

        // Extract chunk
        let end_byte = byte_index_at(split_char);
        let start_byte = byte_index_at(current_char);

        if end_byte > start_byte {
            let chunk = text[start_byte..end_byte].trim().to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            current_char = split_char;
        } else {
            // Should not happen if logic is correct, but prevent infinite loop
            // Force advance by 1 char or target if something is wrong
            let next_char = std::cmp::min(current_char + effective_target, char_len);
            if next_char == current_char {
                // Force break
                break;
            }
            let chunk = text[start_byte..byte_index_at(next_char)]
                .trim()
                .to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            current_char = next_char;
        }
    }

    chunks
}

/// Returns true if this character is an opening quote.
fn is_opening_quote(ch: char) -> bool {
    OPENING_QUOTES.contains(&ch)
}

/// Returns true if this character is a closing quote.
fn is_closing_quote(ch: char) -> bool {
    CLOSING_QUOTES.contains(&ch)
}

/// Detects if text is predominantly inside a dialogue (quoted text).
/// Uses a simple heuristic: tracks open/close quote balance.
pub fn detect_dialogue_in_text(text: &str) -> bool {
    let mut inside_dialogue = false;
    let mut dialogue_char_count = 0usize;
    let mut total_char_count = 0usize;

    for ch in text.chars() {
        if !ch.is_whitespace() {
            total_char_count += 1;
            if inside_dialogue {
                dialogue_char_count += 1;
            }
        }

        if is_opening_quote(ch) {
            inside_dialogue = true;
        } else if is_closing_quote(ch) {
            inside_dialogue = false;
        }
    }

    // If more than half the text is inside quotes, consider it a dialogue chunk
    if total_char_count == 0 {
        return false;
    }
    dialogue_char_count * 2 > total_char_count
}

pub fn split_into_tts_chunks(
    text: &str,
    split_on_newline: bool,
    dictionary: &[DictionaryEntry],
) -> Vec<TtsChunk> {
    let mut sentences = Vec::new();
    let mut current_sentence = String::new();
    let mut current_orig_len = 0usize;
    let chars: Vec<char> = text.chars().collect();
    for (idx, ch) in chars.iter().copied().enumerate() {
        current_sentence.push(ch);
        current_orig_len += 1;
        let next_ch = chars.get(idx + 1).copied();
        let punct_before_newline = split_on_newline
            && matches!(
                ch,
                '.' | '!'
                    | '?'
                    | ','
                    | ';'
                    | ':'
                    | ')'
                    | ']'
                    | '}'
                    | '"'
                    | '\''
                    | '-'
                    | '…'
                    | '—'
                    | '«'
                    | '»'
                    | '“'
                    | '”'
                    | '‘'
                    | '’'
                    | '‐'
                    | '‑'
                    | '–'
                    | '·'
            )
            && matches!(next_ch, Some('\n') | Some('\r'));
        let is_terminal = !punct_before_newline
            && (matches!(ch, '.' | '!' | '?') || (split_on_newline && ch == '\n'));
        if is_terminal {
            if !current_sentence.trim().is_empty() {
                sentences.push((current_sentence.clone(), current_orig_len));
            }
            current_sentence.clear();
            current_orig_len = 0;
        }
    }
    if !current_sentence.trim().is_empty() {
        sentences.push((current_sentence, current_orig_len));
    }
    let mut chunks = Vec::new();
    let mut current_chunk_text = String::new();
    let mut current_chunk_orig_len = 0usize;
    let max_chars = 150;
    for (s_text, s_len) in sentences {
        let potential_new_len = current_chunk_text.chars().count() + s_text.chars().count();
        if !current_chunk_text.is_empty() && potential_new_len > max_chars {
            let cleaned = strip_dashed_lines(&current_chunk_text);
            let prepared = prepare_tts_text(&cleaned, split_on_newline, dictionary);
            let is_dialogue = detect_dialogue_in_text(&current_chunk_text);
            chunks.push(TtsChunk {
                text_to_read: prepared,
                original_len: current_chunk_orig_len,
                is_dialogue,
            });
            current_chunk_text.clear();
            current_chunk_orig_len = 0;
        }
        current_chunk_text.push_str(&s_text);
        current_chunk_orig_len += s_len;
    }
    if !current_chunk_text.is_empty() {
        let cleaned = strip_dashed_lines(&current_chunk_text);
        let prepared = prepare_tts_text(&cleaned, split_on_newline, dictionary);
        let is_dialogue = detect_dialogue_in_text(&current_chunk_text);
        chunks.push(TtsChunk {
            text_to_read: prepared,
            original_len: current_chunk_orig_len,
            is_dialogue,
        });
    }
    chunks
}

pub fn start_audiobook(hwnd: HWND) {
    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return;
    };
    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let text = unsafe { get_edit_text(hwnd_edit) };
    if text.trim().is_empty() {
        unsafe {
            show_error(hwnd, language, &settings::tts_no_text_message(language));
        }
        return;
    }
    let suggested_name = unsafe {
        with_state(hwnd, |state| {
            state.docs.get(state.current).map(|doc| {
                let p = Path::new(&doc.title);
                p.file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or(&doc.title)
                    .to_string()
            })
        })
    }
    .flatten();

    let Some(output) = (unsafe { save_audio_dialog(hwnd, suggested_name.as_deref()) }) else {
        return;
    };
    let voice = unsafe {
        with_state(hwnd, |state| state.settings.tts_voice.clone())
            .unwrap_or_else(|| "it-IT-IsabellaNeural".to_string())
    };

    let (
        split_on_newline,
        audiobook_split,
        audiobook_split_by_text,
        audiobook_split_text,
        audiobook_split_text_requires_newline,
        tts_engine,
        dictionary,
        tts_rate,
        tts_pitch,
        tts_volume,
        use_dialogue_voice,
        dialogue_voice,
    ) = unsafe {
        with_state(hwnd, |state| {
            (
                state.settings.split_on_newline,
                state.settings.audiobook_split,
                state.settings.audiobook_split_by_text,
                state.settings.audiobook_split_text.clone(),
                state.settings.audiobook_split_text_requires_newline,
                state.settings.tts_engine,
                state.settings.dictionary.clone(),
                state.settings.tts_rate,
                state.settings.tts_pitch,
                state.settings.tts_volume,
                state.settings.use_dialogue_voice,
                state.settings.dialogue_voice.clone(),
            )
        })
    }
    .unwrap_or((
        true,
        0,
        false,
        String::new(),
        true,
        TtsEngine::Edge,
        Vec::new(),
        0,
        0,
        100,
        false,
        String::new(),
    ));

    let dialogue_voice_option = if use_dialogue_voice && !dialogue_voice.is_empty() {
        Some(dialogue_voice)
    } else {
        None
    };

    let cleaned = strip_dashed_lines(&text);
    let mut split_parts = audiobook_split;
    let mut marker_parts: Option<Vec<Vec<String>>> = None;
    let mut sapi4_threads: Option<u32> = None;

    if tts_engine == TtsEngine::Sapi4 {
        let title = i18n::tr(language, "audiobook.sapi4_threads_title");

        let body = i18n::tr(language, "audiobook.sapi4_threads_body");

        if let Some(val_str) = unsafe {
            crate::app_windows::prompt_window::prompt_user(hwnd, &title, &body, "30", language)
        } {
            if let Ok(val) = val_str.parse::<u32>() {
                sapi4_threads = Some(val.clamp(1, 100));
            }
        } else {
            // User cancelled the prompt, abort the whole process

            return;
        }
    }

    if audiobook_split_by_text {
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
            if positions.is_empty() {
                split_parts = 0;
            } else {
                marker_parts = build_audiobook_parts_by_positions(
                    &normalized,
                    &positions,
                    split_on_newline,
                    &dictionary,
                );
                if marker_parts.is_none() {
                    split_parts = 0;
                }
            }
        }
    }

    let prepared = if marker_parts.is_some() {
        String::new()
    } else {
        prepare_tts_text(&cleaned, split_on_newline, &dictionary)
    };

    let chunks = if marker_parts.is_some() {
        Vec::new()
    } else {
        split_text(&prepared)
    };
    let chunks_len = if let Some(parts) = &marker_parts {
        parts.iter().map(|part| part.len()).sum()
    } else {
        chunks.len()
    };

    let cancel_token = Arc::new(AtomicBool::new(false));
    let progress_hwnd = unsafe {
        let h = crate::app_windows::audiobook_window::open(hwnd, chunks_len);
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
        let extension = output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";

        let options = AudiobookCommonOptions {
            voice: &voice,
            output: &output,
            progress_hwnd,
            cancel: cancel_clone,
            language,
            rate: tts_rate,
            pitch: tts_pitch,
            volume: tts_volume,
            dialogue_voice: dialogue_voice_option,
            sapi4_threads,
        };

        let result = match tts_engine {
            TtsEngine::Edge => {
                if let Some(ref parts) = marker_parts {
                    if is_aac {
                        // Merge all marker parts into one for AAC
                        let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                        run_split_audiobook(&all_chunks, 0, options)
                    } else {
                        run_marker_split_audiobook(parts, options)
                    }
                } else {
                    run_split_audiobook(&chunks, if is_aac { 0 } else { split_parts }, options)
                }
            }
            TtsEngine::Sapi4 => {
                let voice_idx = parse_sapi4_voice_index(&voice);
                if let Some(ref parts) = marker_parts {
                    if is_aac {
                        let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                        run_split_sapi4_audiobook(&all_chunks, voice_idx, 0, options)
                    } else {
                        run_marker_split_sapi4_audiobook(parts, voice_idx, options)
                    }
                } else {
                    run_split_sapi4_audiobook(
                        &chunks,
                        voice_idx,
                        if is_aac { 0 } else { split_parts },
                        options,
                    )
                }
            }
            TtsEngine::Sapi5 => {
                if let Some(ref parts) = marker_parts {
                    if is_aac {
                        let all_chunks: Vec<String> = parts.iter().flatten().cloned().collect();
                        run_split_sapi_audiobook(&all_chunks, 0, options)
                    } else {
                        run_marker_split_sapi_audiobook(parts, options)
                    }
                } else {
                    run_split_sapi_audiobook(&chunks, if is_aac { 0 } else { split_parts }, options)
                }
            }
        };
        let success = result.is_ok();

        if success && is_aac {
            // Set metadata for the single M4B file
            let file_title = output
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
                &output,
                Some(file_title),
                Some("Novapad"),
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
            Err(err) => err,
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

fn parse_sapi4_voice_index(voice: &str) -> i32 {
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
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
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            dialogue_voice: options.dialogue_voice.clone(),
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            dialogue_voice: options.dialogue_voice.clone(),
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            dialogue_voice: options.dialogue_voice.clone(),
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            options.output.to_path_buf()
        };

        let part_options = AudiobookCommonOptions {
            voice: options.voice,
            output: &part_output,
            progress_hwnd: options.progress_hwnd,
            cancel: options.cancel.clone(),
            language: options.language,
            rate: options.rate,
            pitch: options.pitch,
            volume: options.volume,
            dialogue_voice: options.dialogue_voice.clone(),
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

fn run_sapi4_parallel_part(
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
                        let _com_guard = crate::com_guard::ComGuard::new_mta();
                        if let Err(e) =
                            crate::mf_encoder::encode_wav_to_audio(&sub_output, &encoded_sub, 128)
                        {
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

    std::mem::drop(tx);
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
        name.strip_prefix("sub_")
            .and_then(|s| s.strip_prefix("part_")) // support both naming conventions if any
            .and_then(|s| s.parse::<usize>().ok())
            .unwrap_or(0)
    });

    let result = if is_mp3 {
        // Fast binary join for MP3
        merge_and_finalize_sapi4_mp3(&produced_files, options.output, options.language)
    } else {
        // This covers M4B (AAC) and standard WAV.
        // For M4B, it joins WAV chunks and then encodes to AAC in one pass.
        merge_and_finalize_sapi4_audio(&produced_files, options.output, options.language)
    };

    std::fs::remove_dir_all(&temp_dir).ok();
    *global_progress = progress_counter.load(Ordering::SeqCst);
    result
}

fn merge_and_finalize_sapi4_mp3(
    mp3_files: &[PathBuf],
    output: &Path,
    _language: Language,
) -> Result<(), String> {
    if mp3_files.is_empty() {
        return Ok(());
    }

    let mut out_file = std::fs::File::create(output).map_err(|e| e.to_string())?;
    for path in mp3_files {
        let mut in_file = std::fs::File::open(path).map_err(|e| e.to_string())?;
        std::io::copy(&mut in_file, &mut out_file).map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn merge_and_finalize_sapi4_audio(
    wav_files: &[PathBuf],
    output: &Path,
    language: Language,
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
            crate::mf_encoder::encode_wav_to_m4b(&final_wav, output)
        } else {
            crate::mf_encoder::encode_wav_to_mp3(&final_wav, output)
        };
        std::fs::remove_file(&final_wav).ok();
        res.map_err(|e| {
            if e.contains("Media Foundation") {
                i18n::tr(language, "sapi5.mf_not_available")
            } else {
                i18n::tr_f(language, "sapi5.mf_error", &[("err", &e)])
            }
        })?;
    }
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            options.output.to_path_buf()
        };

        let extension = part_output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
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
                dialogue_voice: options.dialogue_voice.clone(),
                output_path: &actual_output,
                language: options.language,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
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
            let res = crate::mf_encoder::encode_wav_to_m4b(&actual_output, &part_output);
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
            let stem = options
                .output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = options
                .output
                .extension()
                .and_then(|s| s.to_str())
                .unwrap_or("mp3");
            options
                .output
                .with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            options.output.to_path_buf()
        };

        let extension = part_output
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";
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
                dialogue_voice: options.dialogue_voice.clone(),
                output_path: &actual_output,
                language: options.language,
                rate: options.rate,
                pitch: options.pitch,
                volume: options.volume,
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
            let res = crate::mf_encoder::encode_wav_to_m4b(&actual_output, &part_output);
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
    let file = std::fs::File::create(options.output).map_err(|err| err.to_string())?;
    let mut writer = BufWriter::new(file);
    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|err| err.to_string())?;

    rt.block_on(async {
        let tasks = chunks.iter().enumerate().map(|(i, chunk)| {
            let chunk = chunk.clone();
            let voice = options.voice.to_string();
            let dialogue_voice = options.dialogue_voice.clone();
            let cancel = options.cancel.clone();
            let lang = options.language;
            let rate = options.rate;
            let pitch = options.pitch;
            let volume = options.volume;
            async move {
                let request_id = Uuid::new_v4().simple().to_string();
                // Detect if this specific chunk is a dialogue
                let is_dialogue = detect_dialogue_in_text(&chunk);
                let chunk_voice = if is_dialogue && dialogue_voice.is_some() {
                    dialogue_voice.as_deref().unwrap_or(&voice)
                } else {
                    &voice
                };

                loop {
                    if cancel.load(Ordering::Relaxed) {
                        return Err("Cancelled".to_string());
                    }
                    match download_audio_chunk_cancel(DownloadChunkOptions {
                        text: &chunk,
                        voice: chunk_voice,
                        request_id: &request_id,
                        rate,
                        pitch,
                        volume,
                        language: lang,
                        cancel: cancel.as_ref(),
                    })
                    .await
                    {
                        Ok(data) => return Ok::<Vec<u8>, String>(data),
                        Err(err) => {
                            if cancel.load(Ordering::Relaxed) {
                                return Err("Cancelled".to_string());
                            }
                            let msg = i18n::tr_f(
                                lang,
                                "tts.chunk_download_retry_wait",
                                &[("index", &(i + 1).to_string()), ("err", &err)],
                            );
                            log_debug(&msg);

                            tokio::select! {
                                _ = tokio::time::sleep(Duration::from_secs(5)) => {},
                                _ = async {
                                    while !cancel.load(Ordering::Relaxed) {
                                        tokio::time::sleep(Duration::from_millis(100)).await;
                                    }
                                } => {
                                    return Err("Cancelled".to_string());
                                }
                            }
                        }
                    }
                }
            }
        });

        let mut stream = futures_util::stream::iter(tasks).buffered(30);

        while let Some(result) = stream.next().await {
            if options.cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(options.language));
            }
            let audio = match result {
                Ok(data) => data,
                Err(e) if e == "Cancelled" => return Err(cancelled_message(options.language)),
                Err(e) => return Err(e),
            };

            writer.write_all(&audio).map_err(|err| err.to_string())?;
            *current_global_progress += 1;
            if options.progress_hwnd.0 != 0 {
                if options.cancel.load(Ordering::Relaxed) {
                    return Err(cancelled_message(options.language));
                }
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
        writer.flush().map_err(|err| err.to_string())?;
        Ok(())
    })?;

    let extension = options
        .output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let is_aac = extension == "m4b" || extension == "m4a" || extension == "mp4";

    if is_aac {
        let wav_path = options.output.with_extension("wav.tmp");
        if let Err(e) = std::fs::rename(options.output, &wav_path) {
            return Err(format!("Failed to rename for M4B conversion: {}", e));
        }
        let res = crate::mf_encoder::encode_wav_to_m4b(&wav_path, options.output);
        std::fs::remove_file(&wav_path).ok();
        res?;
    }

    Ok(())
}
