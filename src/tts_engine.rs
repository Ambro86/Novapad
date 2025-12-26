use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use std::path::Path;
use std::io::{BufWriter, Write};
use chrono::Local;
use futures_util::{SinkExt, StreamExt};
use rand::Rng;
use rodio::{Decoder, OutputStream, Sink};
use sha2::{Digest, Sha256};
use tokio::sync::mpsc;
use tokio_tungstenite::{
    connect_async,
    tungstenite::client::IntoClientRequest,
    tungstenite::http::HeaderValue,
    tungstenite::protocol::Message,
};
use url::Url;
use uuid::Uuid;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{PostMessageW, SendMessageW, WM_APP};
use windows::Win32::System::Power::{SetThreadExecutionState, ES_CONTINUOUS, ES_SYSTEM_REQUIRED};
use crate::{with_state, get_active_edit, log_debug, show_error, save_audio_dialog};
use crate::settings;
use crate::editor_manager::get_edit_text;
use crate::settings::{Language, AudiobookResult, TRUSTED_CLIENT_TOKEN};
use crate::accessibility::{EM_GETSEL};

pub const WSS_URL_BASE: &str = "wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1";
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
    let (language, split_on_newline) = unsafe {
        with_state(hwnd, |state| {
            (state.settings.language, state.settings.split_on_newline)
        })
    }
    .unwrap_or_default();
    let text = unsafe { get_text_from_caret(hwnd_edit) };
    if text.trim().is_empty() {
        unsafe {
            show_error(hwnd, language, settings::tts_no_text_message(language));
        }
        return;
    }
    let voice = unsafe {
        with_state(hwnd, |state| state.settings.tts_voice.clone()).unwrap_or_else(|| {
            "it-IT-IsabellaNeural".to_string()
        })
    };
    let mut start: i32 = 0;
    let mut end: i32 = 0;
    unsafe { SendMessageW(hwnd_edit, EM_GETSEL, WPARAM(&mut start as *mut _ as usize), LPARAM(&mut end as *mut _ as isize)); }
    let initial_caret_pos = start.min(end);

    let chunks = split_into_tts_chunks(&text, split_on_newline);
    start_tts_playback_with_chunks(hwnd, text, voice, chunks, initial_caret_pos);
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

pub fn start_tts_playback_with_chunks(hwnd: HWND, cleaned: String, voice: String, chunks: Vec<TtsChunk>, initial_caret_pos: i32) {
    stop_tts_playback(hwnd);
    prevent_sleep(true);
    if chunks.is_empty() {
        return;
    }

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
                post_tts_error(hwnd_copy, session_id, "Audio output device not available.".to_string());
                return;
            }
        };
        let sink = match Sink::try_new(&handle) {
            Ok(sink) => sink,
            Err(_) => {
                post_tts_error(hwnd_copy, session_id, "Failed to create audio sink.".to_string());
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

        rt.spawn(async move {
            for chunk_obj in chunks_downloader {
                if cancel_downloader.load(Ordering::SeqCst) { break; }
                let request_id = Uuid::new_v4().simple().to_string();
                match download_audio_chunk(&chunk_obj.text_to_read, &voice_downloader, &request_id).await {
                    Ok(data) => {
                        if audio_tx.send(Ok((data, chunk_obj.original_len))).await.is_err() { break; }
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
            if cancel_flag.load(Ordering::SeqCst) { break; }

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

            let Some(res) = packet else { break; };
            let (audio, orig_len) = match res {
                Ok(data) => data,
                Err(e) => {
                    post_tts_error(hwnd_copy, session_id, e);
                    break;
                }
            };

            if audio.is_empty() { continue; }

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
            if cancel_flag.load(Ordering::SeqCst) { break; }
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
) -> Result<Vec<u8>, String> {
    let max_retries = 5;
    let mut last_error = String::new();

    for attempt in 1..=max_retries {
        match download_audio_chunk_attempt(text, voice, request_id).await {
            Ok(data) => return Ok(data),
            Err(e) => {
                last_error = e;
                log_debug(&format!(
                    "Errore download chunk (tentativo {}/{}) : {}. Riprovo...",
                    attempt, max_retries, last_error
                ));
                if attempt < max_retries {
                    tokio::time::sleep(Duration::from_millis(500 * attempt as u64)).await;
                }
            }
        }
    }
    Err(format!("Falliti {} tentativi. Ultimo errore: {}", max_retries, last_error))
}

async fn download_audio_chunk_attempt(
    text: &str,
    voice: &str,
    request_id: &str,
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
    headers.insert("Origin", HeaderValue::from_static("chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold"));
    headers.insert("User-Agent", HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"));
    headers.insert("Accept-Encoding", HeaderValue::from_static("gzip, deflate, br"));
    headers.insert("Accept-Language", HeaderValue::from_static("en-US,en;q=0.9"));
    let cookie = format!("muid={};", generate_muid());
    headers.insert("Cookie", HeaderValue::from_str(&cookie).map_err(|err| err.to_string())?);

    let (ws_stream, _) = connect_async(request).await.map_err(|err| err.to_string())?;
    let (mut write, mut read) = ws_stream.split();

    let config_msg = format!(
        "X-Timestamp:{}\r\nContent-Type:application/json; charset=utf-8\r\nPath:speech.config\r\n\r\n{{\"context\":{{\"synthesis\":{{\"audio\":{{\"metadataoptions\":{{\"sentenceBoundaryEnabled\":\"false\",\"wordBoundaryEnabled\":\"false\"}},\"outputFormat\":\"audio-24khz-48kbitrate-mono-mp3\"}}}}}}}}",
        get_date_string()
    );
    write.send(Message::Text(config_msg)).await.map_err(|err| err.to_string())?;

    let ssml = mkssml(text, voice);
    let ssml_msg = format!(
        "X-RequestId:{}\r\nContent-Type:application/ssml+xml\r\nX-Timestamp:{}Z\r\nPath:ssml\r\n\r\n{}",
        request_id,
        get_date_string(),
        ssml
    );
    write.send(Message::Text(ssml_msg)).await.map_err(|err| err.to_string())?;

    let mut audio_data = Vec::new();
    while let Some(msg) = read.next().await {
        let msg = msg.map_err(|err| err.to_string())?;
        match msg {
            Message::Text(text) => { if text.contains("Path:turn.end") { break; } }
            Message::Binary(data) => {
                if data.len() < 2 { continue; }
                let be_len = u16::from_be_bytes([data[0], data[1]]) as usize;
                let le_len = u16::from_le_bytes([data[0], data[1]]) as usize;
                let mut parsed = false;
                for header_len in [be_len, le_len] {
                    if header_len == 0 || data.len() < header_len + 2 { continue; }
                    let headers_bytes = &data[2..2 + header_len];
                    let headers_str = String::from_utf8_lossy(headers_bytes);
                    if headers_str.contains("Path:audio") {
                        audio_data.extend_from_slice(&data[2 + header_len..]);
                        parsed = true;
                        break;
                    }
                }
                if parsed { continue; }
            }
            Message::Close(_) => break,
            _ => {}
        }
    }
    Ok(audio_data)
}

unsafe fn get_text_from_caret(hwnd_edit: HWND) -> String {
    let mut start: i32 = 0;
    let mut end: i32 = 0;
    SendMessageW(hwnd_edit, crate::accessibility::EM_GETSEL, WPARAM(&mut start as *mut _ as usize), LPARAM(&mut end as *mut _ as isize));
    let caret_pos = start.min(end).max(0) as usize;
    let full_text = get_edit_text(hwnd_edit);
    
    // Se siamo all'inizio, o se siamo alla fine del testo, leggi tutto dall'inizio
    if caret_pos == 0 {
        return full_text;
    }
    
    let wide: Vec<u16> = full_text.replace("\r\n", "\n").encode_utf16().collect();
    
    // Se la posizione del cursore è oltre la lunghezza del testo (fine file), 
    // ricomincia a leggere dall'inizio come richiesto.
    if caret_pos >= wide.len() {
        return full_text;
    }
    
    String::from_utf16_lossy(&wide[caret_pos..])
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
    Local::now().format("%a %b %d %Y %H:%M:%S GMT+0000 (Coordinated Universal Time)").to_string()
}

fn mkssml(text: &str, voice: &str) -> String {
    let lang = voice.split('-').collect::<Vec<_>>();
    let lang = if lang.len() >= 3 { lang[0..2].join("-") } else { "en-US".to_string() };
    format!("<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='{}'><voice name='{}'><prosody pitch='+0Hz' rate='+0%' volume='+0%'>{}</prosody></voice></speak>", lang, voice, text)
}

pub fn remove_long_dash_runs(line: &str) -> String {
    let mut out = String::with_capacity(line.len());
    let mut dash_run = 0;
    for ch in line.chars() {
        if ch == '-' { dash_run += 1; continue; }
        if dash_run > 0 { if dash_run < 3 { out.extend(std::iter::repeat('-').take(dash_run)); } dash_run = 0; }
        out.push(ch);
    }
    if dash_run > 0 && dash_run < 3 { out.extend(std::iter::repeat('-').take(dash_run)); }
    out
}

pub fn strip_dashed_lines(text: &str) -> String {
    text.lines().filter_map(|line| {
        let trimmed = line.trim();
        if trimmed.is_empty() { return Some(String::new()); }
        let cleaned = remove_long_dash_runs(line);
        if cleaned.trim().is_empty() { None } else { Some(cleaned) }
    }).collect::<Vec<_>>().join("\n")
}

pub fn normalize_for_tts(text: &str, split_on_newline: bool) -> String {
    if split_on_newline { text.to_string() } else { text.replace('\n', " ").replace('\r', "") }
}

pub fn split_text(text: &str) -> Vec<String> {
    let mut chunks = Vec::new();
    let char_indices: Vec<(usize, char)> = text.char_indices().collect();
    let char_len = char_indices.len();
    let mut current_char = 0;
    let is_long = char_len > TTS_LONG_TEXT_THRESHOLD;
    let max_len = if is_long { MAX_TTS_TEXT_LEN_LONG } else { MAX_TTS_TEXT_LEN };
    let first_len = if is_long { MAX_TTS_FIRST_CHUNK_LEN_LONG } else { max_len };
    let byte_index_at = |char_idx: usize| -> usize { if char_idx >= char_len { text.len() } else { char_indices[char_idx].0 } };
    while current_char < char_len {
        let target_len = if chunks.is_empty() { first_len } else { max_len };
        let mut split_char = current_char + target_len;
        if split_char >= char_len {
            let chunk = text[byte_index_at(current_char)..].trim().to_string();
            if !chunk.is_empty() { chunks.push(chunk); }
            break;
        }
        let search_end = split_char;
        let search_start = current_char;
        let mut split_found = None;
        for idx in (search_start..search_end).rev() {
            let c = char_indices[idx].1;
            if c == '.' || c == '!' || c == '?' {
                let next_idx = idx + 1;
                if next_idx >= char_len || char_indices[next_idx].1.is_whitespace() { split_found = Some(next_idx); break; }
            }
        }
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() {
                let c = char_indices[idx].1;
                if c == '\n' { if idx + 1 < char_len && char_indices[idx + 1].1 == '\n' { split_found = Some(idx + 2); break; } }
                else if c == ';' || c == ':' { split_found = Some(idx + 1); break; }
            }
        }
        if split_found.is_none() {
            for idx in (search_start..search_end).rev() { if char_indices[idx].1 == ' ' { split_found = Some(idx + 1); break; } }
        }
        if let Some(split_at) = split_found { split_char = split_at; }
        if split_char > current_char {
            let chunk = text[byte_index_at(current_char)..byte_index_at(split_char)].trim().to_string();
            if !chunk.is_empty() { chunks.push(chunk); }
            current_char = split_char;
        } else {
            let hard_limit = std::cmp::min(current_char + target_len, char_len);
            let chunk = text[byte_index_at(current_char)..byte_index_at(hard_limit)].trim().to_string();
            if !chunk.is_empty() { chunks.push(chunk); }
            current_char = hard_limit;
        }
    }
    chunks
}

pub fn split_into_tts_chunks(text: &str, split_on_newline: bool) -> Vec<TtsChunk> {
    let mut sentences = Vec::new();
    let mut current_sentence = String::new();
    let mut current_orig_len = 0usize;
    for ch in text.chars() {
        current_sentence.push(ch);
        current_orig_len += 1;
        let is_terminal = matches!(ch, '.' | '!' | '?') || (split_on_newline && ch == '\n');
        if is_terminal {
            if !current_sentence.trim().is_empty() { sentences.push((current_sentence.clone(), current_orig_len)); }
            current_sentence.clear();
            current_orig_len = 0;
        }
    }
    if !current_sentence.trim().is_empty() { sentences.push((current_sentence, current_orig_len)); }
    let mut chunks = Vec::new();
    let mut current_chunk_text = String::new();
    let mut current_chunk_orig_len = 0usize;
    let max_chars = 150;
    for (s_text, s_len) in sentences {
        let potential_new_len = current_chunk_text.chars().count() + s_text.chars().count();
        if !current_chunk_text.is_empty() && potential_new_len > max_chars {
            let cleaned = strip_dashed_lines(&current_chunk_text);
            let prepared = normalize_for_tts(&cleaned, split_on_newline);
            chunks.push(TtsChunk { text_to_read: prepared, original_len: current_chunk_orig_len });
            current_chunk_text.clear();
            current_chunk_orig_len = 0;
        }
        current_chunk_text.push_str(&s_text);
        current_chunk_orig_len += s_len;
    }
    if !current_chunk_text.is_empty() {
        let cleaned = strip_dashed_lines(&current_chunk_text);
        let prepared = normalize_for_tts(&cleaned, split_on_newline);
        chunks.push(TtsChunk { text_to_read: prepared, original_len: current_chunk_orig_len });
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
                p.file_stem().and_then(|s| s.to_str()).unwrap_or(&doc.title).to_string()
            })
        })
    }.flatten();

    let Some(output) = (unsafe { save_audio_dialog(hwnd, suggested_name.as_deref()) }) else {
        return;
    };
    let voice = unsafe {
        with_state(hwnd, |state| state.settings.tts_voice.clone()).unwrap_or_else(|| {
            "it-IT-IsabellaNeural".to_string()
        })
    };

    let split_on_newline = unsafe { with_state(hwnd, |state| state.settings.split_on_newline) }
        .unwrap_or(true);

    let cleaned = strip_dashed_lines(&text);
    let prepared = normalize_for_tts(&cleaned, split_on_newline);

    let chunks = split_text(&prepared);
    let chunks_len = chunks.len();
    
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
        let result = run_tts_audiobook(&chunks, &voice, &output, progress_hwnd, cancel_clone);
        let success = result.is_ok();
        let message = match result {
            Ok(()) => match language {
                Language::Italian => "Audiolibro salvato con successo.".to_string(),
                Language::English => "Audiobook saved successfully.".to_string(),
            },
            Err(err) => err,
        };
        let payload = Box::new(AudiobookResult {
            success,
            message,
        });
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

pub fn run_tts_audiobook(
    chunks: &[String],
    voice: &str,
    output: &Path,
    progress_hwnd: HWND,
    cancel: Arc<AtomicBool>,
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
                    match download_audio_chunk(&chunk, &voice, &request_id).await {
                        Ok(data) => return Ok::<Vec<u8>, String>(data),
                        Err(err) => {
                            if cancel.load(Ordering::Relaxed) {
                                return Err("Cancelled".to_string());
                            }
                            let msg = format!("Errore download chunk {}: {}. Riprovo tra 5 secondi...", i + 1, err);
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
        let mut completed = 0;
        while let Some(result) = stream.next().await {
            if cancel.load(Ordering::Relaxed) {
                return Err("Operazione annullata.".to_string());
            }
            let audio = match result {
                Ok(data) => data,
                Err(e) if e == "Cancelled" => return Err("Operazione annullata.".to_string()),
                Err(e) => return Err(e),
            };
            
            writer.write_all(&audio).map_err(|err| err.to_string())?;
            completed += 1;
            if progress_hwnd.0 != 0 {
                if cancel.load(Ordering::Relaxed) { return Err("Operazione annullata.".to_string()); }
                unsafe { let _ = PostMessageW(progress_hwnd, crate::WM_UPDATE_PROGRESS, WPARAM(completed), LPARAM(0)); }
            }
        }
        writer.flush().map_err(|err| err.to_string())?;
        Ok(())
    }).map_err(|e| {
        if e == "Operazione annullata." {
            let _ = std::fs::remove_file(output);
        }
        e
    })
}