#![allow(
    clippy::collapsible_if,
    clippy::manual_div_ceil,
    clippy::manual_inspect,
    clippy::manual_repeat_n,
    clippy::collapsible_str_replace,
    clippy::too_many_arguments
)]
use crate::editor_manager::get_edit_text;
use crate::settings;
use crate::settings::{
    AudiobookResult, DictionaryEntry, Language, TRUSTED_CLIENT_TOKEN, TtsEngine,
};
use crate::{get_active_edit, log_debug, save_audio_dialog, show_error, with_state};
use chrono::Local;
use futures_util::{SinkExt, StreamExt};
use rand::Rng;
use rodio::{Decoder, OutputStream, Sink};
use sha2::{Digest, Sha256};
use std::io::{BufWriter, Write};
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use tokio::sync::mpsc;
use tokio_tungstenite::{
    connect_async, tungstenite::client::IntoClientRequest, tungstenite::http::HeaderValue,
    tungstenite::protocol::Message,
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

#[derive(Clone)]
pub struct TtsChunk {
    pub text_to_read: String,
    pub original_len: usize,
}

fn cancelled_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Operazione annullata.",
        Language::English => "Operation cancelled.",
    }
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
    let _ = unsafe {
        PostMessageW(
            hwnd,
            WM_TTS_PLAYBACK_ERROR,
            WPARAM(session_id as usize),
            LPARAM(Box::into_raw(payload) as isize),
        )
    };
}

pub fn start_tts_from_caret(hwnd: HWND) {
    let Some(hwnd_edit) = (unsafe { get_active_edit(hwnd) }) else {
        return;
    };
    let (language, split_on_newline, tts_engine, dictionary, tts_rate, tts_pitch, tts_volume) =
        unsafe {
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

    let (text, initial_caret_pos) = unsafe { get_text_from_caret(hwnd_edit) };
    if text.trim().is_empty() {
        unsafe {
            show_error(hwnd, language, settings::tts_no_text_message(language));
        }
        return;
    }
    let voice = unsafe {
        with_state(hwnd, |state| state.settings.tts_voice.clone())
            .unwrap_or_else(|| "it-IT-IsabellaNeural".to_string())
    };
    let chunks = split_into_tts_chunks(&text, split_on_newline, &dictionary);

    match tts_engine {
        TtsEngine::Edge => start_tts_playback_with_chunks(
            hwnd,
            text,
            voice,
            chunks,
            initial_caret_pos,
            tts_rate,
            tts_pitch,
            tts_volume,
        ),
        TtsEngine::Sapi5 => {
            // Stop any existing playback
            stop_tts_playback(hwnd);
            let cancel = Arc::new(AtomicBool::new(false));
            let (command_tx, command_rx) = mpsc::unbounded_channel();
            let _ = unsafe {
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
            };

            let chunk_strings: Vec<String> = chunks.into_iter().map(|c| c.text_to_read).collect();
            let _ = crate::sapi5_engine::play_sapi(
                chunk_strings,
                voice,
                tts_rate,
                tts_pitch,
                tts_volume,
                cancel,
                command_rx,
            );
        }
    }
}

pub fn toggle_tts_pause(hwnd: HWND) {
    let _ = unsafe {
        with_state(hwnd, |state| {
            let Some(session) = &mut state.tts_session else {
                return;
            };
            if session.paused {
                prevent_sleep(true);
                let _ = session.command_tx.send(TtsCommand::Resume);
                session.paused = false;
            } else {
                prevent_sleep(false);
                let _ = session.command_tx.send(TtsCommand::Pause);
                session.paused = true;
            }
        })
    };
}

pub fn stop_tts_playback(hwnd: HWND) {
    prevent_sleep(false);
    let _ = unsafe {
        with_state(hwnd, |state| {
            if let Some(session) = &state.tts_session {
                session.cancel.store(true, Ordering::SeqCst);
                let _ = session.command_tx.send(TtsCommand::Stop);
            }
            state.tts_session = None;
        })
    };
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

pub fn start_tts_playback_with_chunks(
    hwnd: HWND,
    cleaned: String,
    voice: String,
    chunks: Vec<TtsChunk>,
    initial_caret_pos: i32,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
) {
    stop_tts_playback(hwnd);
    prevent_sleep(true);
    if chunks.is_empty() {
        return;
    }

    let language = unsafe { with_state(hwnd, |state| state.settings.language) }.unwrap_or_default();
    let (tx, mut rx) = mpsc::unbounded_channel::<TtsCommand>();
    let cancel = Arc::new(AtomicBool::new(false));
    let cancel_flag = cancel.clone();
    let session_id = unsafe {
        with_state(hwnd, |state| {
            let id = state.tts_next_session_id;
            state.tts_next_session_id = state.tts_next_session_id.saturating_add(1);
            state.tts_session = Some(TtsSession {
                id,
                command_tx: tx.clone(),
                cancel: cancel.clone(),
                paused: false,
                initial_caret_pos,
            });
            id
        })
        .unwrap_or(0)
    };
    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        log_debug(&format!(
            "TTS start: voice={voice} chunks={} text_len={}",
            chunks.len(),
            cleaned.len()
        ));
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(values) => values,
            Err(_) => {
                post_tts_error(
                    hwnd_copy,
                    session_id,
                    "Audio output device not available.".to_string(),
                );
                return;
            }
        };
        let sink = match Sink::try_new(&handle) {
            Ok(sink) => sink,
            Err(_) => {
                post_tts_error(
                    hwnd_copy,
                    session_id,
                    "Failed to create audio sink.".to_string(),
                );
                return;
            }
        };
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
        let rate = tts_rate;
        let pitch = tts_pitch;
        let volume = tts_volume;

        rt.spawn(async move {
            for chunk_obj in chunks_downloader {
                if cancel_downloader.load(Ordering::SeqCst) {
                    break;
                }
                let request_id = Uuid::new_v4().simple().to_string();
                match download_audio_chunk(
                    &chunk_obj.text_to_read,
                    &voice_downloader,
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
                        let _ = audio_tx.send(Err(e)).await;
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
                    if cancel_flag.load(Ordering::SeqCst) { return None; }
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
                            if let Some(cmd) = cmd_opt {
                                if handle_tts_command(cmd, &sink, cancel_flag.as_ref(), &mut paused) {
                                    return None;
                                }
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

            let _ = unsafe {
                PostMessageW(
                    hwnd_copy,
                    WM_TTS_CHUNK_START,
                    WPARAM(session_id as usize),
                    LPARAM(current_offset as isize),
                )
            };

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
            let _ = unsafe {
                PostMessageW(
                    hwnd_copy,
                    WM_TTS_PLAYBACK_DONE,
                    WPARAM(session_id as usize),
                    LPARAM(0),
                )
            };
        }
    });
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
    let max_retries = 5;
    let mut last_error = String::new();

    for attempt in 1..=max_retries {
        log_debug(&format!(
            "Edge TTS: chunk start attempt={}/{} voice={} text_len={} request_id={}",
            attempt,
            max_retries,
            voice,
            text.chars().count(),
            request_id
        ));
        match download_audio_chunk_attempt(text, voice, request_id, tts_rate, tts_pitch, tts_volume)
            .await
        {
            Ok(data) => return Ok(data),
            Err(e) => {
                last_error = e;
                let msg = match language {
                    Language::Italian => format!(
                        "Errore download chunk (tentativo {}/{}) : {}. Riprovo...",
                        attempt, max_retries, last_error
                    ),
                    Language::English => format!(
                        "Chunk download error (attempt {}/{}): {}. Retrying...",
                        attempt, max_retries, last_error
                    ),
                };
                log_debug(&msg);
                if attempt < max_retries {
                    tokio::time::sleep(Duration::from_millis(500 * attempt as u64)).await;
                }
            }
        }
    }
    Err(match language {
        Language::Italian => format!(
            "Falliti {} tentativi. Ultimo errore: {}",
            max_retries, last_error
        ),
        Language::English => format!(
            "Failed {} attempts. Last error: {}",
            max_retries, last_error
        ),
    })
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
    let sec_ms_gec_version = "1-130.0.2849.68";

    let url_str = format!(
        "{}?TrustedClientToken={}&ConnectionId={}&Sec-MS-GEC={}&Sec-MS-GEC-Version={}",
        WSS_URL_BASE, TRUSTED_CLIENT_TOKEN, request_id, sec_ms_gec, sec_ms_gec_version
    );
    let url = Url::parse(&url_str).map_err(|err| err.to_string())?;

    let mut request = url.into_client_request().map_err(|err| err.to_string())?;
    let headers = request.headers_mut();
    headers.insert("Pragma", HeaderValue::from_static("no-cache"));
    headers.insert("Cache-Control", HeaderValue::from_static("no-cache"));
    headers.insert(
        "Origin",
        HeaderValue::from_static("chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold"),
    );
    headers.insert("User-Agent", HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"));
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

    log_debug(&format!(
        "Edge TTS: ws connect start voice={} request_id={}",
        voice, request_id
    ));
    let connect_timeout = Duration::from_secs(3);
    let (ws_stream, _) = match tokio::time::timeout(connect_timeout, connect_async(request)).await {
        Ok(res) => res.map_err(|err| err.to_string())?,
        Err(_) => {
            log_debug(&format!(
                "Edge TTS: ws connect timeout {}s request_id={}",
                connect_timeout.as_secs(),
                request_id
            ));
            return Err("WebSocket connect timeout".to_string());
        }
    };
    log_debug(&format!(
        "Edge TTS: ws connect ok voice={} request_id={}",
        voice, request_id
    ));
    let (mut write, mut read) = ws_stream.split();

    let config_msg = format!(
        "X-Timestamp:{}\r\nContent-Type:application/json; charset=utf-8\r\nPath:speech.config\r\n\r\n{{\"context\":{{\"synthesis\":{{\"audio\":{{\"metadataoptions\":{{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"}},\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}}}}}",
        get_date_string()
    );
    write
        .send(Message::Text(config_msg))
        .await
        .map_err(|err| err.to_string())?;
    log_debug(&format!(
        "Edge TTS: ws config sent request_id={}",
        request_id
    ));

    let ssml = mkssml(text, voice, tts_rate, tts_pitch, tts_volume);
    let ssml_msg = format!(
        "X-RequestId:{}\r\nContent-Type:application/ssml+xml\r\nX-Timestamp:{}Z\r\nPath:ssml\r\n\r\n{}",
        request_id,
        get_date_string(),
        ssml
    );
    write
        .send(Message::Text(ssml_msg))
        .await
        .map_err(|err| err.to_string())?;
    log_debug(&format!(
        "Edge TTS: ws ssml sent request_id={}",
        request_id
    ));

    let mut audio_data = Vec::new();
    let mut audio_frames = 0usize;
    let mut saw_turn_end = false;
    while let Some(msg) = read.next().await {
        let msg = msg.map_err(|err| err.to_string())?;
        match msg {
            Message::Text(text) => {
                if text.contains("Path:turn.end") {
                    saw_turn_end = true;
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
                        audio_frames += 1;
                        parsed = true;
                        break;
                    }
                }
                if parsed {
                    continue;
                }
            }
            Message::Close(_) => {
                log_debug(&format!(
                    "Edge TTS: ws closed request_id={} audio_bytes={} audio_frames={}",
                    request_id,
                    audio_data.len(),
                    audio_frames
                ));
                break;
            }
            _ => {}
        }
    }
    log_debug(&format!(
        "Edge TTS: ws done request_id={} audio_bytes={} audio_frames={} turn_end={}",
        request_id,
        audio_data.len(),
        audio_frames,
        saw_turn_end
    ));
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
    if prev_is_word && next_is_word {
        let mut idx = prev.unwrap();
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
    let ticks = (ticks - (ticks % 300)) as f64 * 1e7;
    let str_to_hash = format!("{:.0}{}", ticks, TRUSTED_CLIENT_TOKEN);
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
                out.extend(std::iter::repeat('-').take(dash_run));
            }
            dash_run = 0;
        }
        out.push(ch);
    }
    if dash_run > 0 && dash_run < 3 {
        out.extend(std::iter::repeat('-').take(dash_run));
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
    normalized.replace('«', "").replace('»', "")
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

fn prepare_tts_text(text: &str, split_on_newline: bool, dictionary: &[DictionaryEntry]) -> String {
    let normalized = normalize_for_tts(text, split_on_newline);
    apply_dictionary(&normalized, dictionary)
}

fn normalize_newlines(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}

struct MarkerEntry {
    pos: usize,
    label: String,
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

fn collect_marker_entries(
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
        if require_newline {
            if idx != 0 && !normalized[..idx].ends_with('\n') {
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
        if *pos > start && *pos <= text.len() {
            parts.push(text[start..*pos].to_string());
            start = *pos;
        }
    }
    if start <= text.len() {
        parts.push(text[start..].to_string());
    }
    Some(parts)
}

fn build_audiobook_parts_by_positions(
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
        if !chunks.is_empty() {
            parts_chunks.push(chunks);
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
    let char_indices: Vec<(usize, char)> = text.char_indices().collect();
    let char_len = char_indices.len();
    let mut current_char = 0;
    let is_long = char_len > TTS_LONG_TEXT_THRESHOLD;
    let max_len = if is_long {
        MAX_TTS_TEXT_LEN_LONG
    } else {
        MAX_TTS_TEXT_LEN
    };
    let first_len = if is_long {
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
        let target_len = if chunks.is_empty() {
            first_len
        } else {
            max_len
        };
        let mut split_char = current_char + target_len;
        if split_char >= char_len {
            let chunk = text[byte_index_at(current_char)..].trim().to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            break;
        }
        let search_end = split_char;
        let search_start = current_char;
        let mut split_found = None;
        for idx in (search_start..search_end).rev() {
            let c = char_indices[idx].1;
            if c == '.' || c == '!' || c == '?' {
                let next_idx = idx + 1;
                if next_idx >= char_len || char_indices[next_idx].1.is_whitespace() {
                    split_found = Some(next_idx);
                    break;
                }
            }
        }
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() {
                let c = char_indices[idx].1;
                if c == '\n' {
                    if idx + 1 < char_len && char_indices[idx + 1].1 == '\n' {
                        split_found = Some(idx + 2);
                        break;
                    }
                } else if c == ';' || c == ':' {
                    split_found = Some(idx + 1);
                    break;
                }
            }
        }
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() {
                if char_indices[idx].1 == ' ' {
                    split_found = Some(idx + 1);
                    break;
                }
            }
        }
        if let Some(split_at) = split_found {
            split_char = split_at;
        }
        if split_char > current_char {
            let chunk = text[byte_index_at(current_char)..byte_index_at(split_char)]
                .trim()
                .to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            current_char = split_char;
        } else {
            let hard_limit = std::cmp::min(current_char + target_len, char_len);
            let chunk = text[byte_index_at(current_char)..byte_index_at(hard_limit)]
                .trim()
                .to_string();
            if !chunk.is_empty() {
                chunks.push(chunk);
            }
            current_char = hard_limit;
        }
    }
    chunks
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
            chunks.push(TtsChunk {
                text_to_read: prepared,
                original_len: current_chunk_orig_len,
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
        chunks.push(TtsChunk {
            text_to_read: prepared,
            original_len: current_chunk_orig_len,
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
            show_error(hwnd, language, settings::tts_no_text_message(language));
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
    ));

    let cleaned = strip_dashed_lines(&text);
    let mut split_parts = audiobook_split;
    let mut marker_parts: Option<Vec<Vec<String>>> = None;
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
        let _ = with_state(hwnd, |state| {
            state.audiobook_progress = h;
            state.audiobook_cancel = Some(cancel_token.clone());
        });
        h
    };

    let cancel_clone = cancel_token.clone();
    std::thread::spawn(move || {
        let result = match tts_engine {
            TtsEngine::Edge => {
                if let Some(parts) = marker_parts {
                    run_marker_split_audiobook(
                        &parts,
                        &voice,
                        &output,
                        progress_hwnd,
                        cancel_clone,
                        language,
                        tts_rate,
                        tts_pitch,
                        tts_volume,
                    )
                } else {
                    run_split_audiobook(
                        &chunks,
                        &voice,
                        &output,
                        split_parts,
                        progress_hwnd,
                        cancel_clone,
                        language,
                        tts_rate,
                        tts_pitch,
                        tts_volume,
                    )
                }
            }
            TtsEngine::Sapi5 => {
                if let Some(parts) = marker_parts {
                    run_marker_split_sapi_audiobook(
                        &parts,
                        &voice,
                        &output,
                        progress_hwnd,
                        cancel_clone,
                        language,
                        tts_rate,
                        tts_pitch,
                        tts_volume,
                    )
                } else {
                    run_split_sapi_audiobook(
                        &chunks,
                        &voice,
                        &output,
                        split_parts,
                        progress_hwnd,
                        cancel_clone,
                        language,
                        tts_rate,
                        tts_pitch,
                        tts_volume,
                    )
                }
            }
        };
        let success = result.is_ok();
        let message = match result {
            Ok(()) => match language {
                Language::Italian => "Audiolibro salvato con successo.".to_string(),
                Language::English => "Audiobook saved successfully.".to_string(),
            },
            Err(err) => err,
        };
        let payload = Box::new(AudiobookResult { success, message });
        let _ = unsafe {
            PostMessageW(
                hwnd,
                crate::WM_TTS_AUDIOBOOK_DONE,
                WPARAM(0),
                LPARAM(Box::into_raw(payload) as isize),
            )
        };
    });
}

fn run_split_audiobook(
    chunks: &[String],
    voice: &str,
    output: &Path,
    split_parts: u32,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
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
    let chunks_per_part = (total_chunks + parts - 1) / parts; // Ceiling division

    let mut current_global_progress = 0;

    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }

        let part_chunks = &chunks[start_idx..end_idx];

        let part_output = if parts > 1 {
            let stem = output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = output.extension().and_then(|s| s.to_str()).unwrap_or("mp3");
            output.with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            output.to_path_buf()
        };

        run_tts_audiobook_part(
            part_chunks,
            voice,
            &part_output,
            progress_hwnd,
            cancel.clone(),
            &mut current_global_progress,
            language,
            tts_rate,
            tts_pitch,
            tts_volume,
        )?;
    }
    Ok(())
}

fn run_marker_split_audiobook(
    parts: &[Vec<String>],
    voice: &str,
    output: &Path,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
) -> Result<(), String> {
    let parts_len = parts.len();
    let mut current_global_progress = 0;

    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if part_chunks.is_empty() {
            continue;
        }
        let part_output = if parts_len > 1 {
            let stem = output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = output.extension().and_then(|s| s.to_str()).unwrap_or("mp3");
            output.with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            output.to_path_buf()
        };

        run_tts_audiobook_part(
            part_chunks,
            voice,
            &part_output,
            progress_hwnd,
            cancel.clone(),
            &mut current_global_progress,
            language,
            tts_rate,
            tts_pitch,
            tts_volume,
        )?;
    }
    Ok(())
}

fn run_split_sapi_audiobook(
    chunks: &[String],
    voice: &str,
    output: &Path,
    split_parts: u32,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
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
    let chunks_per_part = (total_chunks + parts - 1) / parts;
    let mut current_global_progress = 0;

    for part_idx in 0..parts {
        let start_idx = part_idx * chunks_per_part;
        let end_idx = std::cmp::min(start_idx + chunks_per_part, total_chunks);
        if start_idx >= end_idx {
            break;
        }

        let part_chunks = &chunks[start_idx..end_idx];

        let part_output = if parts > 1 {
            let stem = output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = output.extension().and_then(|s| s.to_str()).unwrap_or("mp3");
            output.with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            output.to_path_buf()
        };

        let progress_hwnd_clone = progress_hwnd;
        let cancel_clone = cancel.clone();

        crate::sapi5_engine::speak_sapi_to_file(
            part_chunks,
            voice,
            &part_output,
            language,
            tts_rate,
            tts_pitch,
            tts_volume,
            cancel_clone,
            |_chunk_idx| {
                current_global_progress += 1;
                if progress_hwnd_clone.0 != 0 {
                    unsafe {
                        let _ = PostMessageW(
                            progress_hwnd_clone,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        );
                    }
                }
            },
        )
        .map_err(|e| {
            let _ = std::fs::remove_file(&part_output);
            e
        })?;
    }
    Ok(())
}

fn run_marker_split_sapi_audiobook(
    parts: &[Vec<String>],
    voice: &str,
    output: &Path,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
) -> Result<(), String> {
    let parts_len = parts.len();
    let mut current_global_progress = 0;

    for (part_idx, part_chunks) in parts.iter().enumerate() {
        if part_chunks.is_empty() {
            continue;
        }
        let part_output = if parts_len > 1 {
            let stem = output
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("audiobook");
            let ext = output.extension().and_then(|s| s.to_str()).unwrap_or("mp3");
            output.with_file_name(format!("{}_part{}.{}", stem, part_idx + 1, ext))
        } else {
            output.to_path_buf()
        };

        let progress_hwnd_clone = progress_hwnd;
        let cancel_clone = cancel.clone();

        crate::sapi5_engine::speak_sapi_to_file(
            part_chunks,
            voice,
            &part_output,
            language,
            tts_rate,
            tts_pitch,
            tts_volume,
            cancel_clone,
            |_chunk_idx| {
                current_global_progress += 1;
                if progress_hwnd_clone.0 != 0 {
                    unsafe {
                        let _ = PostMessageW(
                            progress_hwnd_clone,
                            crate::WM_UPDATE_PROGRESS,
                            WPARAM(current_global_progress),
                            LPARAM(0),
                        );
                    }
                }
            },
        )
        .map_err(|e| {
            let _ = std::fs::remove_file(&part_output);
            e
        })?;
    }
    Ok(())
}

fn run_tts_audiobook_part(
    chunks: &[String],
    voice: &str,
    output: &Path,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
    current_global_progress: &mut usize,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
) -> Result<(), String> {
    let file = std::fs::File::create(output).map_err(|err| err.to_string())?;
    let mut writer = BufWriter::new(file);
    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|err| err.to_string())?;

    rt.block_on(async {
        let tasks = chunks.iter().enumerate().map(|(i, chunk)| {
            let chunk = chunk.clone();
            let voice = voice.to_string();
            let cancel = cancel.clone();
            async move {
                let request_id = Uuid::new_v4().simple().to_string();
                loop {
                    if cancel.load(Ordering::Relaxed) {
                        return Err("Cancelled".to_string());
                    }
                    match download_audio_chunk(
                        &chunk,
                        &voice,
                        &request_id,
                        tts_rate,
                        tts_pitch,
                        tts_volume,
                        language,
                    )
                    .await
                    {
                        Ok(data) => return Ok::<Vec<u8>, String>(data),
                        Err(err) => {
                            if cancel.load(Ordering::Relaxed) {
                                return Err("Cancelled".to_string());
                            }
                            let msg = match language {
                                Language::Italian => format!(
                                    "Errore download chunk {}: {}. Riprovo tra 5 secondi...",
                                    i + 1,
                                    err
                                ),
                                Language::English => format!(
                                    "Chunk download error {}: {}. Retrying in 5 seconds...",
                                    i + 1,
                                    err
                                ),
                            };
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
            if cancel.load(Ordering::Relaxed) {
                return Err(cancelled_message(language).to_string());
            }
            let audio = match result {
                Ok(data) => data,
                Err(e) if e == "Cancelled" => return Err(cancelled_message(language).to_string()),
                Err(e) => return Err(e),
            };

            writer.write_all(&audio).map_err(|err| err.to_string())?;
            *current_global_progress += 1;
            if progress_hwnd.0 != 0 {
                if cancel.load(Ordering::Relaxed) {
                    return Err(cancelled_message(language).to_string());
                }
                unsafe {
                    let _ = PostMessageW(
                        progress_hwnd,
                        crate::WM_UPDATE_PROGRESS,
                        WPARAM(*current_global_progress),
                        LPARAM(0),
                    );
                }
            }
        }
        writer.flush().map_err(|err| err.to_string())?;
        Ok(())
    })
    .map_err(|e| {
        if e == cancelled_message(language) {
            let _ = std::fs::remove_file(output);
        }
        e
    })
}
