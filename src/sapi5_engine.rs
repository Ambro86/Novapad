#![allow(
    clippy::collapsible_if,
    clippy::unnecessary_map_or,
    clippy::too_many_arguments
)]
use crate::accessibility::to_wide;
use crate::settings::{Language, VoiceInfo};
use crate::tts_engine::TtsCommand;
use std::collections::HashSet;
use std::collections::VecDeque;
use std::io::{Read, Seek, Write};
use std::path::Path;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use tokio::sync::mpsc;
use windows::Win32::Media::Audio::{WAVE_FORMAT_PCM, WAVEFORMATEX};
use windows::Win32::Media::Speech::{
    IEnumSpObjectTokens, ISpObjectToken, ISpObjectTokenCategory, ISpStream, ISpVoice, SPF_ASYNC,
    SPF_IS_XML, SPF_PURGEBEFORESPEAK, SPFM_CREATE_ALWAYS, SPRS_DONE, SPVOICESTATUS, SpFileStream,
    SpObjectTokenCategory, SpVoice,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoTaskMemFree,
    CoUninitialize,
};
use windows::core::{GUID, PCWSTR, w};

// SPDFID_WaveFormatEx: {C31ADBAE-527F-4ff5-A230-F62BB61FF70C}
const SPDFID_WAVEFORMATEX: GUID = GUID::from_values(
    0xC31ADBAE,
    0x527F,
    0x4ff5,
    [0xA2, 0x30, 0xF6, 0x2B, 0xB6, 0x1F, 0xF7, 0x0C],
);
const SAPI_VOICES_PATH: PCWSTR = w!(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices");
const ONECORE_VOICES_PATH: PCWSTR =
    w!(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices");

unsafe fn wcslen(ptr: *const u16) -> usize {
    let mut len = 0;
    while *ptr.add(len) != 0 {
        len += 1;
    }
    len
}

unsafe fn collect_voice_descriptions(category_id: PCWSTR) -> Result<Vec<String>, String> {
    let category: ISpObjectTokenCategory =
        CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL)
            .map_err(|e| format!("CoCreateInstance(Category) failed: {}", e))?;
    category
        .SetId(category_id, false)
        .map_err(|e| format!("SetId failed: {}", e))?;

    let enum_tokens: IEnumSpObjectTokens = category
        .EnumTokens(None, None)
        .map_err(|e| format!("EnumTokens failed: {}", e))?;

    let mut count = 0;
    enum_tokens.GetCount(&mut count).ok();

    let mut voices = Vec::new();
    for i in 0..count {
        if let Ok(token) = enum_tokens.Item(i) {
            if let Ok(desc_ptr) = token.GetStringValue(PCWSTR::null()) {
                let description = if !desc_ptr.is_null() {
                    let s = String::from_utf16_lossy(std::slice::from_raw_parts(
                        desc_ptr.as_ptr(),
                        wcslen(desc_ptr.as_ptr()),
                    ));
                    CoTaskMemFree(Some(desc_ptr.as_ptr() as *const _));
                    s
                } else {
                    "Unknown Voice".to_string()
                };
                voices.push(description);
            }
        }
    }
    Ok(voices)
}

unsafe fn find_voice_token(voice_name: &str) -> Option<ISpObjectToken> {
    for category_id in [SAPI_VOICES_PATH, ONECORE_VOICES_PATH] {
        let category: windows::core::Result<ISpObjectTokenCategory> =
            CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL);
        if let Ok(cat) = category {
            let _ = cat.SetId(category_id, false);
            if let Ok(enum_tokens) = cat.EnumTokens(None, None) {
                let mut count = 0;
                if enum_tokens.GetCount(&mut count).is_ok() {
                    for i in 0..count {
                        if let Ok(tok) = enum_tokens.Item(i) {
                            if let Ok(desc_ptr) = tok.GetStringValue(PCWSTR::null()) {
                                if !desc_ptr.is_null() {
                                    let description =
                                        String::from_utf16_lossy(std::slice::from_raw_parts(
                                            desc_ptr.as_ptr(),
                                            wcslen(desc_ptr.as_ptr()),
                                        ));
                                    CoTaskMemFree(Some(desc_ptr.as_ptr() as *const _));
                                    if description == voice_name {
                                        return Some(tok);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    None
}

pub fn list_sapi_voices() -> Result<Vec<VoiceInfo>, String> {
    unsafe {
        // Use APARTMENTTHREADED for better compatibility with SAPI5
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

        let mut names = Vec::new();
        if let Ok(list) = collect_voice_descriptions(SAPI_VOICES_PATH) {
            names.extend(list);
        }
        if let Ok(list) = collect_voice_descriptions(ONECORE_VOICES_PATH) {
            names.extend(list);
        }

        let mut seen = HashSet::new();
        let mut voices = Vec::new();
        for name in names {
            if seen.insert(name.clone()) {
                voices.push(VoiceInfo {
                    short_name: name,
                    locale: "SAPI5".to_string(),
                    is_multilingual: false,
                });
            }
        }
        Ok(voices)
    }
}

pub fn play_sapi(
    chunks: Vec<String>,
    voice_name: String,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    cancel: Arc<AtomicBool>,
    mut command_rx: mpsc::UnboundedReceiver<TtsCommand>,
) -> Result<(), String> {
    std::thread::spawn(move || {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_err() {
                crate::log_debug(&format!("SAPI playback: CoInitializeEx failed: {:?}", hr));
                return;
            }

            let voice_res: windows::core::Result<ISpVoice> =
                CoCreateInstance(&SpVoice, None, CLSCTX_ALL);
            let voice = match voice_res {
                Ok(v) => v,
                Err(e) => {
                    crate::log_debug(&format!("SAPI playback: Failed to create SpVoice: {}", e));
                    return;
                }
            };

            if let Some(token) = find_voice_token(&voice_name) {
                let _ = voice.SetVoice(&token);
            }
            let _ = voice.SetRate(map_sapi_rate(tts_rate));
            let _ = voice.SetVolume(map_sapi_volume(tts_volume));

            let mut paused = false;
            let mut pending: VecDeque<String> = VecDeque::from(chunks);

            while let Some(chunk) = pending.pop_front() {
                // Wait here if a pause was requested between chunks.
                while paused {
                    if cancel.load(Ordering::Relaxed) {
                        let _ = voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None);
                        return;
                    }
                    while let Ok(cmd) = command_rx.try_recv() {
                        match cmd {
                            TtsCommand::Resume => {
                                paused = false;
                            }
                            TtsCommand::Stop => {
                                cancel.store(true, Ordering::SeqCst);
                                let _ = voice.Speak(
                                    PCWSTR::null(),
                                    SPF_PURGEBEFORESPEAK.0 as u32,
                                    None,
                                );
                                return;
                            }
                            TtsCommand::Pause => {}
                        }
                    }
                    std::thread::sleep(Duration::from_millis(50));
                }

                if cancel.load(Ordering::Relaxed) {
                    break;
                }

                let current_chunk = chunk;
                let ssml = mk_sapi_ssml(&current_chunk, tts_rate, tts_pitch, tts_volume);
                let chunk_wide = to_wide(&ssml);
                let _ = voice.Speak(
                    PCWSTR(chunk_wide.as_ptr()),
                    (SPF_ASYNC.0 | SPF_IS_XML.0) as u32,
                    None,
                );

                loop {
                    if cancel.load(Ordering::Relaxed) {
                        let _ = voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None);
                        return;
                    }
                    while let Ok(cmd) = command_rx.try_recv() {
                        match cmd {
                            TtsCommand::Pause => {
                                let mut status = SPVOICESTATUS::default();
                                let mut remainder: Option<String> = None;
                                if voice.GetStatus(&mut status, std::ptr::null_mut()).is_ok() {
                                    let pos = status.ulInputWordPos as usize;
                                    let wide: Vec<u16> = current_chunk.encode_utf16().collect();
                                    let start = pos.min(wide.len());
                                    if start < wide.len() {
                                        let tail = String::from_utf16_lossy(&wide[start..]);
                                        if !tail.trim().is_empty() {
                                            remainder = Some(tail);
                                        }
                                    }
                                }
                                let _ = voice.Speak(
                                    PCWSTR::null(),
                                    SPF_PURGEBEFORESPEAK.0 as u32,
                                    None,
                                );
                                if let Some(rem) = remainder {
                                    pending.push_front(rem);
                                }
                                paused = true;
                                break;
                            }
                            TtsCommand::Resume => {
                                paused = false;
                            }
                            TtsCommand::Stop => {
                                cancel.store(true, Ordering::SeqCst);
                                let _ = voice.Speak(
                                    PCWSTR::null(),
                                    SPF_PURGEBEFORESPEAK.0 as u32,
                                    None,
                                );
                                return;
                            }
                        }
                    }
                    if paused {
                        break;
                    }

                    let mut status = SPVOICESTATUS::default();
                    if voice.GetStatus(&mut status, std::ptr::null_mut()).is_ok() {
                        if status.dwRunningState == SPRS_DONE.0 as u32 {
                            break;
                        }
                    }
                    let _ = voice.WaitUntilDone(50);
                }
            }
        }
    });
    Ok(())
}

pub fn speak_sapi_to_file(
    chunks: &[String],
    voice_name: &str,
    output_path: &Path,
    language: Language,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    cancel: Arc<AtomicBool>,
    mut progress_callback: impl FnMut(usize),
) -> Result<(), String> {
    unsafe {
        let com_initialized = CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok();
        struct ComGuard(bool);
        impl Drop for ComGuard {
            fn drop(&mut self) {
                if self.0 {
                    unsafe {
                        CoUninitialize();
                    }
                }
            }
        }
        let _com_guard = ComGuard(com_initialized);

        let is_mp3 = output_path
            .extension()
            .map_or(false, |e| e.eq_ignore_ascii_case("mp3"));
        crate::log_debug(&format!(
            "SAPI: is_mp3={}, output_path={:?}",
            is_mp3, output_path
        ));
        let wav_path = if is_mp3 {
            output_path.with_extension("wav.tmp")
        } else {
            output_path.to_path_buf()
        };
        crate::log_debug(&format!("SAPI: Target wav_path={:?}", wav_path));

        {
            let voice: ISpVoice = CoCreateInstance(&SpVoice, None, CLSCTX_ALL)
                .map_err(|e| format!("Failed to create SpVoice: {}", e))?;

            let voice_token = find_voice_token(voice_name).ok_or_else(|| {
                "Selected SAPI voice not found. Please select a voice in Options.".to_string()
            })?;
            voice
                .SetVoice(&voice_token)
                .map_err(|e| format!("SetVoice failed: {}", e))?;
            let _ = voice.SetRate(map_sapi_rate(tts_rate));
            let _ = voice.SetVolume(map_sapi_volume(tts_volume));

            let stream: ISpStream = CoCreateInstance(&SpFileStream, None, CLSCTX_ALL)
                .map_err(|e| format!("Failed to create SpFileStream: {}", e))?;

            let path_wide = to_wide(wav_path.to_str().ok_or("Invalid path")?);

            let mut wfx = WAVEFORMATEX::default();
            wfx.wFormatTag = WAVE_FORMAT_PCM as u16;
            wfx.nChannels = 1;
            wfx.nSamplesPerSec = 44100;
            wfx.wBitsPerSample = 16;
            wfx.nBlockAlign = wfx.nChannels * (wfx.wBitsPerSample / 8);
            wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * (wfx.nBlockAlign as u32);
            wfx.cbSize = 0;

            stream
                .BindToFile(
                    PCWSTR(path_wide.as_ptr()),
                    SPFM_CREATE_ALWAYS,
                    Some(&SPDFID_WAVEFORMATEX),
                    Some(&wfx),
                    0,
                )
                .map_err(|e| format!("BindToFile failed: {}", e))?;

            voice
                .SetOutput(&stream, true)
                .map_err(|e| format!("SetOutput failed: {}", e))?;

            for (i, chunk) in chunks.iter().enumerate() {
                if cancel.load(Ordering::Relaxed) {
                    let _ = stream.Close();
                    let _ = std::fs::remove_file(&wav_path);
                    return Err("Cancelled".to_string());
                }
                let ssml = mk_sapi_ssml(chunk, tts_rate, tts_pitch, tts_volume);
                let chunk_wide = to_wide(&ssml);
                voice
                    .Speak(PCWSTR(chunk_wide.as_ptr()), SPF_IS_XML.0 as u32, None)
                    .map_err(|e| format!("Speak failed: {}", e))?;

                progress_callback(i + 1);
            }

            let _ = voice.WaitUntilDone(u32::MAX);
            let _ = voice.SetOutput(None, false);
            let _ = stream.Close();
        }

        if is_mp3 {
            if let Ok(data_size) = wav_data_size(&wav_path) {
                if data_size == 0 {
                    let sample_rate = 44100u32;
                    let channels = 1u16;
                    let bits_per_sample = 16u16;
                    let duration_ms = 500u32;
                    let _ = write_silence_wav(
                        &wav_path,
                        sample_rate,
                        channels,
                        bits_per_sample,
                        duration_ms,
                    );
                    crate::log_debug(
                        "SAPI: WAV had no data; wrote 500ms silence for MP3 encoding.",
                    );
                }
            }
            match crate::mf_encoder::encode_wav_to_mp3(&wav_path, output_path) {
                Ok(()) => {
                    let _ = std::fs::remove_file(&wav_path);
                }
                Err(e) => {
                    let dest_wav = output_path.with_extension("wav");
                    let _ = std::fs::rename(&wav_path, &dest_wav);
                    let msg = if e.contains("Media Foundation not available") {
                        mf_not_available_message(language)
                    } else {
                        mf_error_message(language, &e)
                    };
                    return Err(msg);
                }
            }
        }

        Ok(())
    }
}

fn mf_not_available_message(language: Language) -> String {
    match language {
        Language::Italian => {
            "Media Foundation non disponibile (Windows N/KN). Installa Media Feature Pack. Salvato in WAV.".to_string()
        }
        Language::English => {
            "Media Foundation not available (Windows N/KN). Install Media Feature Pack. Saved as WAV.".to_string()
        }
    }
}

fn mf_error_message(language: Language, err: &str) -> String {
    match language {
        Language::Italian => format!("Errore MP3 Media Foundation: {}. Salvato in WAV.", err),
        Language::English => format!("Media Foundation MP3 error: {}. Saved as WAV.", err),
    }
}

fn map_sapi_rate(rate_percent: i32) -> i32 {
    (rate_percent / 10).clamp(-10, 10)
}

fn map_sapi_volume(volume: i32) -> u16 {
    let vol = volume.clamp(0, 100);
    vol as u16
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

fn mk_sapi_ssml(text: &str, rate: i32, pitch: i32, volume: i32) -> String {
    let escaped = escape_xml(text);
    format!(
        "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis'><prosody pitch='{}' rate='{}' volume='{}'>{}</prosody></speak>",
        format_pitch(pitch),
        format_rate(rate),
        format_volume(volume),
        escaped
    )
}

fn wav_data_size(path: &Path) -> Result<u32, String> {
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut riff_header = [0u8; 12];
    file.read_exact(&mut riff_header)
        .map_err(|e| e.to_string())?;
    if &riff_header[0..4] != b"RIFF" || &riff_header[8..12] != b"WAVE" {
        return Err("Invalid WAV header".to_string());
    }

    loop {
        let mut chunk_header = [0u8; 8];
        if file.read_exact(&mut chunk_header).is_err() {
            break;
        }
        let chunk_id = &chunk_header[0..4];
        let chunk_size = u32::from_le_bytes(chunk_header[4..8].try_into().unwrap());
        if chunk_id == b"data" {
            return Ok(chunk_size);
        }
        file.seek(std::io::SeekFrom::Current(chunk_size as i64))
            .map_err(|e| e.to_string())?;
        if chunk_size % 2 == 1 {
            file.seek(std::io::SeekFrom::Current(1))
                .map_err(|e| e.to_string())?;
        }
    }
    Err("WAV data chunk not found".to_string())
}

fn write_silence_wav(
    path: &Path,
    sample_rate: u32,
    channels: u16,
    bits_per_sample: u16,
    duration_ms: u32,
) -> Result<(), String> {
    let bytes_per_sample = (bits_per_sample / 8) as u32;
    let samples = sample_rate.saturating_mul(duration_ms) / 1000;
    let data_size = samples
        .saturating_mul(channels as u32)
        .saturating_mul(bytes_per_sample);
    let riff_size = 36u32.saturating_add(data_size);

    let mut file = std::fs::File::create(path).map_err(|e| e.to_string())?;
    file.write_all(b"RIFF").map_err(|e| e.to_string())?;
    file.write_all(&riff_size.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(b"WAVE").map_err(|e| e.to_string())?;

    file.write_all(b"fmt ").map_err(|e| e.to_string())?;
    file.write_all(&16u32.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(&1u16.to_le_bytes())
        .map_err(|e| e.to_string())?; // PCM
    file.write_all(&channels.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(&sample_rate.to_le_bytes())
        .map_err(|e| e.to_string())?;
    let byte_rate = sample_rate
        .saturating_mul(channels as u32)
        .saturating_mul(bytes_per_sample);
    let block_align = (channels as u32 * bytes_per_sample) as u16;
    file.write_all(&byte_rate.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(&block_align.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(&bits_per_sample.to_le_bytes())
        .map_err(|e| e.to_string())?;

    file.write_all(b"data").map_err(|e| e.to_string())?;
    file.write_all(&data_size.to_le_bytes())
        .map_err(|e| e.to_string())?;
    let zeros = vec![0u8; 4096];
    let mut remaining = data_size as usize;
    while remaining > 0 {
        let chunk = remaining.min(zeros.len());
        file.write_all(&zeros[..chunk]).map_err(|e| e.to_string())?;
        remaining -= chunk;
    }
    Ok(())
}
