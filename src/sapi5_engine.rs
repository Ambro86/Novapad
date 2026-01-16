#![allow(
    clippy::collapsible_if,
    clippy::unnecessary_map_or,
    clippy::too_many_arguments
)]
use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::{Language, VoiceInfo};
use crate::tts_engine::TtsCommand;
use std::collections::HashSet;
use std::collections::VecDeque;
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

fn pwstr_to_string(ptr: PCWSTR) -> String {
    if ptr.is_null() {
        return "Unknown Voice".to_string();
    }
    unsafe {
        ptr.to_string()
            .unwrap_or_else(|_| "Unknown Voice".to_string())
    }
}

struct ComRuntime;

impl ComRuntime {
    fn new() -> Result<Self, windows::core::Error> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok()?;
        Ok(Self)
    }
}

impl Drop for ComRuntime {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

fn collect_voice_descriptions(category_id: PCWSTR) -> Result<Vec<String>, String> {
    let _com = ComRuntime::new().map_err(|e| format!("CoInitializeEx failed: {}", e))?;

    let category: ISpObjectTokenCategory = unsafe {
        CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL)
            .map_err(|e| format!("CoCreateInstance(Category) failed: {}", e))?
    };

    unsafe { category.SetId(category_id, false) }.map_err(|e| format!("SetId failed: {}", e))?;

    let enum_tokens: IEnumSpObjectTokens = unsafe { category.EnumTokens(None, None) }
        .map_err(|e| format!("EnumTokens failed: {}", e))?;

    let mut count = 0;
    if unsafe { enum_tokens.GetCount(&mut count) }.is_err() {
        return Ok(Vec::new());
    }

    let mut voices = Vec::new();
    for i in 0..count {
        // Safe wrapper around token operations
        if let Ok(token) = unsafe { enum_tokens.Item(i) } {
            if let Ok(desc_ptr) = unsafe { token.GetStringValue(PCWSTR::null()) } {
                let description = pwstr_to_string(PCWSTR(desc_ptr.0));
                unsafe {
                    CoTaskMemFree(Some(desc_ptr.0 as *const _));
                }
                voices.push(description);
            }
        }
    }
    Ok(voices)
}

fn find_voice_token(voice_name: &str) -> Option<ISpObjectToken> {
    for category_id in [SAPI_VOICES_PATH, ONECORE_VOICES_PATH] {
        let category: windows::core::Result<ISpObjectTokenCategory> =
            unsafe { CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL) };
        if let Ok(cat) = category {
            let _ = unsafe { cat.SetId(category_id, false) };
            if let Ok(enum_tokens) = unsafe { cat.EnumTokens(None, None) } {
                let mut count = 0;
                if unsafe { enum_tokens.GetCount(&mut count) }.is_ok() {
                    for i in 0..count {
                        if let Ok(tok) = unsafe { enum_tokens.Item(i) } {
                            if let Ok(desc_ptr) = unsafe { tok.GetStringValue(PCWSTR::null()) } {
                                let description = pwstr_to_string(PCWSTR(desc_ptr.0));
                                unsafe {
                                    CoTaskMemFree(Some(desc_ptr.0 as *const _));
                                }
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
    None
}

pub fn list_sapi_voices() -> Result<Vec<VoiceInfo>, String> {
    let _com = ComRuntime::new().map_err(|e| format!("CoInitializeEx failed: {}", e))?;

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
            if let Ok(data_size) = crate::audio_utils::get_wav_data_size(&wav_path) {
                if data_size == 0 {
                    let sample_rate = 44100u32;
                    let channels = 1u16;
                    let bits_per_sample = 16u16;
                    let duration_ms = 500u32;
                    let _ = crate::audio_utils::write_silence_file(
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
    i18n::tr(language, "sapi5.mf_not_available")
}

fn mf_error_message(language: Language, err: &str) -> String {
    i18n::tr_f(language, "sapi5.mf_error", &[("err", err)])
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
