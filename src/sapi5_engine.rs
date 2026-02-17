use crate::accessibility::to_wide;
use crate::com_guard::ComGuard;
use crate::i18n;
use crate::settings::{Language, VoiceInfo};
use crate::tts_engine::TtsCommand;
use std::collections::HashSet;
use std::collections::VecDeque;
use std::io::{Read, Seek, SeekFrom, Write};
use std::os::windows::process::CommandExt;
use std::path::Path;
use std::path::PathBuf;
use std::process::{Command, Stdio};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;
use tokio::sync::mpsc;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::Media::Audio::{WAVE_FORMAT_PCM, WAVEFORMATEX};
use windows::Win32::Media::Speech::{
    ISpEventSource, ISpMMSysAudio, ISpObjectToken, ISpObjectTokenCategory, ISpStream, ISpVoice,
    ISpeechObjectToken, ISpeechObjectTokenCategory, ISpeechVoice, SPAS_PAUSE, SPAS_RUN, SPAS_STOP,
    SPEI_WORD_BOUNDARY, SPEVENT, SPF_ASYNC, SPF_IS_XML, SPF_PURGEBEFORESPEAK, SPFM_CREATE_ALWAYS,
    SPRS_DONE, SPVOICESTATUS, SpFileStream, SpMMAudioOut, SpObjectTokenCategory, SpVoice,
};
use windows::Win32::System::Com::{CLSCTX_ALL, CoCreateInstance, CoTaskMemFree};
use windows::core::{BSTR, GUID, Interface, PCWSTR, w};

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
const SAPI_VOICES_PATH_STR: &str = r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices";
const ONECORE_VOICES_PATH_STR: &str =
    r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices";
const SAPI_VOICES_PATH_HKCU_STR: &str = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\Voices";
const ONECORE_VOICES_PATH_HKCU_STR: &str =
    r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech_OneCore\Voices";
static SAPI5_TMP_SEQ: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(1);

fn repair_wav_header_sizes_if_needed(path: &Path) -> Result<bool, String> {
    let file_size = std::fs::metadata(path).map_err(|e| e.to_string())?.len();
    if file_size < 44 {
        return Ok(false);
    }

    let mut file = std::fs::OpenOptions::new()
        .read(true)
        .write(true)
        .open(path)
        .map_err(|e| e.to_string())?;
    let mut header = [0u8; 44];
    file.read_exact(&mut header).map_err(|e| e.to_string())?;

    // Keep this repair strictly for canonical PCM WAV header layout.
    if &header[0..4] != b"RIFF" || &header[8..12] != b"WAVE" || &header[36..40] != b"data" {
        return Ok(false);
    }

    let data_size = u32::from_le_bytes([header[40], header[41], header[42], header[43]]);
    if data_size != 0 {
        return Ok(false);
    }

    let data_size_u32 = file_size.saturating_sub(44).min(u32::MAX as u64) as u32;
    let riff_size_u32 = file_size.saturating_sub(8).min(u32::MAX as u64) as u32;

    file.seek(SeekFrom::Start(4)).map_err(|e| e.to_string())?;
    file.write_all(&riff_size_u32.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.seek(SeekFrom::Start(40)).map_err(|e| e.to_string())?;
    file.write_all(&data_size_u32.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.flush().map_err(|e| e.to_string())?;
    Ok(true)
}

fn select_sapi_bridge_exe() -> Result<PathBuf, String> {
    let exe_path = std::env::current_exe().map_err(|e| e.to_string())?;
    let dir = exe_path.parent().ok_or("Missing exe dir")?;
    let settings_dir = crate::settings::settings_dir();
    let candidates = [
        settings_dir.join("sapi4_bridge_32.exe"),
        settings_dir.join("sapi4_bridge_x86.exe"),
        settings_dir.join("sapi4_bridge.exe"),
        dir.join("sapi4_bridge_32.exe"),
        dir.join("sapi4_bridge_x86.exe"),
        dir.join("sapi4_bridge.exe"),
    ];
    for path in candidates {
        if path.exists() {
            return Ok(path);
        }
    }
    Err("SAPI bridge executable not found".to_string())
}

fn collect_voice_descriptions_from_sapi5_bridge() -> Vec<String> {
    let exe_path = match select_sapi_bridge_exe() {
        Ok(path) => path,
        Err(_) => return Vec::new(),
    };
    let out = match Command::new(exe_path)
        .arg("--sapi5-list")
        .creation_flags(0x08000000)
        .output()
    {
        Ok(o) => o,
        Err(_) => return Vec::new(),
    };
    let stdout = String::from_utf8_lossy(&out.stdout);
    stdout
        .lines()
        .filter_map(|line| line.strip_prefix("SAPI5VOICE:"))
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect()
}

fn pwstr_to_string(ptr: PCWSTR) -> String {
    if ptr.is_null() {
        return "Unknown Voice".to_string();
    }
    unsafe {
        ptr.to_string()
            .unwrap_or_else(|_| "Unknown Voice".to_string())
    }
}

fn token_display_name(token: &ISpObjectToken) -> Option<String> {
    if let Ok(value_ptr) = unsafe { token.GetStringValue(PCWSTR::null()) } {
        let value = pwstr_to_string(PCWSTR(value_ptr.0));
        unsafe {
            CoTaskMemFree(Some(value_ptr.0 as *const _));
        }
        let trimmed = value.trim();
        if !trimmed.is_empty() {
            return Some(trimmed.to_string());
        }
    }

    if let Ok(name_ptr) = unsafe { token.GetStringValue(w!("Name")) } {
        let name = pwstr_to_string(PCWSTR(name_ptr.0));
        unsafe {
            CoTaskMemFree(Some(name_ptr.0 as *const _));
        }
        let trimmed = name.trim();
        if !trimmed.is_empty() {
            return Some(trimmed.to_string());
        }
    }

    None
}

fn speech_token_display_name(token: &ISpeechObjectToken) -> Option<String> {
    if let Ok(desc) = unsafe { token.GetDescription(0) } {
        let d = desc.to_string();
        if !d.trim().is_empty() {
            return Some(d);
        }
    }
    if let Ok(id) = unsafe { token.Id() } {
        let s = id.to_string();
        if !s.trim().is_empty() {
            return Some(s);
        }
    }
    None
}

fn collect_voice_descriptions_from_speech_voice() -> Result<Vec<String>, String> {
    let _com = ComGuard::new_sta().map_err(|e| format!("CoInitializeEx failed: {}", e))?;
    let voice: ISpeechVoice = unsafe {
        CoCreateInstance(&SpVoice, None, CLSCTX_ALL)
            .map_err(|e| format!("CoCreateInstance(ISpeechVoice) failed: {}", e))?
    };
    let required = BSTR::from("");
    let optional = BSTR::from("");
    let tokens = unsafe { voice.GetVoices(&required, &optional) }
        .map_err(|e| format!("ISpeechVoice.GetVoices failed: {}", e))?;
    let count = unsafe { tokens.Count() }.map_err(|e| format!("tokens.Count failed: {}", e))?;
    let mut voices = Vec::new();
    for i in 0..count {
        if let Ok(token) = unsafe { tokens.Item(i) }
            && let Some(name) = speech_token_display_name(&token)
        {
            voices.push(name);
        }
    }
    Ok(voices)
}

fn collect_voice_descriptions_from_speech_category(
    category_id: &str,
) -> Result<Vec<String>, String> {
    let _com = ComGuard::new_sta().map_err(|e| format!("CoInitializeEx failed: {}", e))?;
    let category: ISpeechObjectTokenCategory = unsafe {
        CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL)
            .map_err(|e| format!("CoCreateInstance(ISpeechObjectTokenCategory) failed: {}", e))?
    };
    let category_id = BSTR::from(category_id);
    unsafe { category.SetId(&category_id, VARIANT_BOOL(0)) }
        .map_err(|e| format!("ISpeechObjectTokenCategory.SetId failed: {}", e))?;
    let required = BSTR::from("");
    let optional = BSTR::from("");
    let tokens = unsafe { category.EnumerateTokens(&required, &optional) }
        .map_err(|e| format!("ISpeechObjectTokenCategory.EnumerateTokens failed: {}", e))?;
    let count = unsafe { tokens.Count() }.map_err(|e| format!("tokens.Count failed: {}", e))?;
    let mut voices = Vec::new();
    for i in 0..count {
        if let Ok(token) = unsafe { tokens.Item(i) }
            && let Some(name) = speech_token_display_name(&token)
        {
            voices.push(name);
        }
    }
    Ok(voices)
}

fn find_voice_token_in_speech_category(
    category_id: &str,
    voice_name: &str,
) -> Option<ISpObjectToken> {
    let category: ISpeechObjectTokenCategory =
        unsafe { CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL).ok()? };
    let category_id = BSTR::from(category_id);
    unsafe { category.SetId(&category_id, VARIANT_BOOL(0)) }.ok()?;
    let required = BSTR::from("");
    let optional = BSTR::from("");
    let tokens = unsafe { category.EnumerateTokens(&required, &optional) }.ok()?;
    let count = unsafe { tokens.Count() }.ok()?;
    for i in 0..count {
        if let Ok(token) = unsafe { tokens.Item(i) }
            && let Some(description) = speech_token_display_name(&token)
            && description == voice_name
            && let Ok(sp_token) = token.cast::<ISpObjectToken>()
        {
            return Some(sp_token);
        }
    }
    None
}

fn find_voice_token(voice_name: &str) -> Option<ISpObjectToken> {
    for category_id in [
        SAPI_VOICES_PATH_STR,
        SAPI_VOICES_PATH_HKCU_STR,
        ONECORE_VOICES_PATH_STR,
        ONECORE_VOICES_PATH_HKCU_STR,
    ] {
        if let Some(token) = find_voice_token_in_speech_category(category_id, voice_name) {
            return Some(token);
        }
    }

    for category_id in [SAPI_VOICES_PATH, ONECORE_VOICES_PATH] {
        let category: windows::core::Result<ISpObjectTokenCategory> =
            unsafe { CoCreateInstance(&SpObjectTokenCategory, None, CLSCTX_ALL) };
        if let Ok(cat) = category {
            if let Err(e) = unsafe { cat.SetId(category_id, false) } {
                crate::log_debug(&format!("Failed to set SAPI5 category: {:?}", e));
            }
            if let Ok(enum_tokens) = unsafe { cat.EnumTokens(None, None) } {
                let mut count = 0;
                if unsafe { enum_tokens.GetCount(&mut count) }.is_ok() {
                    for i in 0..count {
                        if let Ok(tok) = unsafe { enum_tokens.Item(i) }
                            && let Some(description) = token_display_name(&tok)
                            && description == voice_name
                        {
                            return Some(tok);
                        }
                    }
                }
            }
        }
    }
    None
}

fn has_sapi5_bridge_voice(voice_name: &str) -> bool {
    collect_voice_descriptions_from_sapi5_bridge()
        .iter()
        .any(|name| name == voice_name)
}

fn has_native_sapi5_voice(voice_name: &str) -> bool {
    let mut names = Vec::new();
    if let Ok(list) = collect_voice_descriptions_from_speech_voice() {
        names.extend(list);
    }
    for category_id in [
        SAPI_VOICES_PATH_STR,
        SAPI_VOICES_PATH_HKCU_STR,
        ONECORE_VOICES_PATH_STR,
        ONECORE_VOICES_PATH_HKCU_STR,
    ] {
        if let Ok(list) = collect_voice_descriptions_from_speech_category(category_id) {
            names.extend(list);
        }
    }
    names.into_iter().any(|n| n == voice_name)
}

pub fn list_sapi_voices() -> Result<Vec<VoiceInfo>, String> {
    let _com = ComGuard::new_sta().map_err(|e| format!("CoInitializeEx failed: {}", e))?;

    let mut names = Vec::new();
    if let Ok(list) = collect_voice_descriptions_from_speech_voice() {
        names.extend(list);
    }
    for category_id in [
        SAPI_VOICES_PATH_STR,
        SAPI_VOICES_PATH_HKCU_STR,
        ONECORE_VOICES_PATH_STR,
        ONECORE_VOICES_PATH_HKCU_STR,
    ] {
        if let Ok(list) = collect_voice_descriptions_from_speech_category(category_id) {
            names.extend(list);
        }
    }
    names.extend(collect_voice_descriptions_from_sapi5_bridge());

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

fn send_bridge_line(stdin: &mut std::process::ChildStdin, line: &str) -> std::io::Result<()> {
    stdin.write_all(line.as_bytes())?;
    stdin.write_all(b"\n")?;
    stdin.flush()
}

fn send_bridge_speak(stdin: &mut std::process::ChildStdin, text: &str) -> std::io::Result<()> {
    let bytes = text.as_bytes();
    stdin.write_all(format!("SPEAK {}\n", bytes.len()).as_bytes())?;
    stdin.write_all(bytes)?;
    stdin.flush()
}

fn play_sapi_via_bridge(
    chunks: Vec<String>,
    voice_name: String,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    cancel: Arc<AtomicBool>,
    mut command_rx: mpsc::UnboundedReceiver<TtsCommand>,
) -> Result<(), String> {
    let exe_path = select_sapi_bridge_exe()?;
    std::thread::spawn(move || {
        let mut child = match Command::new(&exe_path)
            .arg("--sapi5-server")
            .arg("--voice-name")
            .arg(&voice_name)
            .arg("--rate")
            .arg(tts_rate.to_string())
            .arg("--pitch")
            .arg(tts_pitch.to_string())
            .arg("--volume")
            .arg(tts_volume.to_string())
            .stdin(Stdio::piped())
            .creation_flags(0x08000000)
            .spawn()
        {
            Ok(child) => child,
            Err(err) => {
                crate::log_debug(&format!("SAPI5 bridge spawn failed: {}", err));
                return;
            }
        };
        let mut stdin = match child.stdin.take() {
            Some(stdin) => stdin,
            None => {
                crate::log_debug("SAPI5 bridge stdin unavailable");
                crate::log_if_err!(child.wait());
                return;
            }
        };
        let text = chunks.join("\n");
        if let Err(err) = send_bridge_speak(&mut stdin, &text) {
            crate::log_debug(&format!("SAPI5 bridge SPEAK failed: {}", err));
            crate::log_if_err!(child.wait());
            return;
        }

        loop {
            if cancel.load(Ordering::Relaxed) {
                crate::log_if_err!(send_bridge_line(&mut stdin, "STOP"));
                break;
            }
            match command_rx.try_recv() {
                Ok(TtsCommand::Pause) => {
                    crate::log_if_err!(send_bridge_line(&mut stdin, "PAUSE"));
                }
                Ok(TtsCommand::Resume) => {
                    crate::log_if_err!(send_bridge_line(&mut stdin, "RESUME"));
                }
                Ok(TtsCommand::Stop) => {
                    crate::log_if_err!(send_bridge_line(&mut stdin, "STOP"));
                    break;
                }
                Err(tokio::sync::mpsc::error::TryRecvError::Empty) => {
                    std::thread::sleep(Duration::from_millis(5));
                }
                Err(tokio::sync::mpsc::error::TryRecvError::Disconnected) => {
                    crate::log_if_err!(send_bridge_line(&mut stdin, "STOP"));
                    break;
                }
            }
        }
        crate::log_if_err!(child.wait());
    });
    Ok(())
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
    if has_sapi5_bridge_voice(&voice_name) && !has_native_sapi5_voice(&voice_name) {
        crate::log_debug(&format!(
            "SAPI5: using 32-bit bridge voice fallback for '{}'",
            voice_name
        ));
        return play_sapi_via_bridge(
            chunks, voice_name, tts_rate, tts_pitch, tts_volume, cancel, command_rx,
        );
    }

    const COMMAND_POLL_MS: u32 = 5;
    const PAUSED_POLL_MS: u32 = 2;
    const HARD_PAUSE_IMMEDIATE: bool = false;

    std::thread::spawn(move || {
        let _com = match ComGuard::new_sta() {
            Ok(g) => g,
            Err(e) => {
                crate::log_debug(&format!("SAPI playback: CoInitializeEx failed: {:?}", e));
                return;
            }
        };

        unsafe {
            let voice_res: windows::core::Result<ISpVoice> =
                CoCreateInstance(&SpVoice, None, CLSCTX_ALL);
            let voice = match voice_res {
                Ok(v) => v,
                Err(e) => {
                    crate::log_debug(&format!("SAPI playback: Failed to create SpVoice: {}", e));
                    return;
                }
            };

            if let Some(token) = find_voice_token(&voice_name)
                && let Err(e) = voice.SetVoice(&token)
            {
                crate::log_debug(&format!("Failed to set SAPI5 voice: {}", e));
            }
            if let Err(e) = voice.SetRate(map_sapi_rate(tts_rate)) {
                crate::log_debug(&format!("Failed to set SAPI5 rate: {}", e));
            }
            if let Err(e) = voice.SetVolume(map_sapi_volume(tts_volume)) {
                crate::log_debug(&format!("Failed to set SAPI5 volume: {}", e));
            }
            let audio_output: Option<ISpMMSysAudio> =
                match CoCreateInstance::<_, ISpMMSysAudio>(&SpMMAudioOut, None, CLSCTX_ALL) {
                    Ok(audio) => {
                        if let Err(e) = voice.SetOutput(&audio, true) {
                            crate::log_debug(&format!("Failed to set SAPI5 audio output: {}", e));
                            None
                        } else {
                            Some(audio)
                        }
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Failed to create SAPI5 MM audio output: {}", e));
                        None
                    }
                };
            let event_source = match voice.cast::<ISpEventSource>() {
                Ok(source) => {
                    let mask = word_boundary_interest_mask();
                    if mask != 0 {
                        match source.SetInterest(mask, mask) {
                            Ok(()) => {}
                            Err(e) => {
                                crate::log_debug(&format!(
                                    "Failed to set SAPI5 event interest: {}",
                                    e
                                ));
                            }
                        }
                    }
                    Some(source)
                }
                Err(e) => {
                    crate::log_debug(&format!(
                        "Failed to cast SAPI5 voice to event source: {}",
                        e
                    ));
                    None
                }
            };

            let mut paused = false;
            let mut paused_by_purge = false;
            let mut pending: VecDeque<String> = VecDeque::from(chunks);

            while let Some(chunk) = pending.pop_front() {
                // Wait here if a pause was requested between chunks.
                while paused {
                    if cancel.load(Ordering::Relaxed) {
                        if let Err(e) =
                            voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None)
                        {
                            crate::log_debug(&format!("SAPI5 Speak failed: {}", e));
                        }
                        return;
                    }
                    while let Ok(cmd) = command_rx.try_recv() {
                        match cmd {
                            TtsCommand::Resume => {
                                if let Some(audio) = &audio_output {
                                    if let Err(e) = audio.SetState(SPAS_RUN, 0) {
                                        crate::log_debug(&format!(
                                            "SAPI5 audio resume failed: {}",
                                            e
                                        ));
                                    }
                                } else if !paused_by_purge && let Err(e) = voice.Resume() {
                                    crate::log_debug(&format!("SAPI5 Resume failed: {}", e));
                                }
                                paused_by_purge = false;
                                paused = false;
                            }
                            TtsCommand::Stop => {
                                cancel.store(true, Ordering::SeqCst);
                                if let Some(audio) = &audio_output
                                    && let Err(e) = audio.SetState(SPAS_STOP, 0)
                                {
                                    crate::log_debug(&format!("SAPI5 audio stop failed: {}", e));
                                }
                                if let Err(e) =
                                    voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None)
                                {
                                    crate::log_debug(&format!(
                                        "SAPI5 Speak purge failed (stop): {}",
                                        e
                                    ));
                                }
                                return;
                            }
                            TtsCommand::Pause => {}
                        }
                    }
                    std::thread::sleep(Duration::from_millis(PAUSED_POLL_MS as u64));
                }

                if cancel.load(Ordering::Relaxed) {
                    break;
                }

                let current_chunk = chunk;
                let ssml = mk_sapi_ssml(&current_chunk, tts_rate, tts_pitch, tts_volume);
                let mut last_word_stream_pos_utf16 = 0usize;
                let mut has_word_boundary = false;
                let chunk_wide = to_wide(&ssml);
                if let Err(e) = voice.Speak(
                    PCWSTR(chunk_wide.as_ptr()),
                    (SPF_ASYNC.0 | SPF_IS_XML.0) as u32,
                    None,
                ) {
                    crate::log_debug(&format!("SAPI5 chunk Speak failed: {}", e));
                }

                loop {
                    if let Some(source) = &event_source
                        && drain_word_boundary_events(source, &mut last_word_stream_pos_utf16)
                    {
                        has_word_boundary = true;
                    }
                    if cancel.load(Ordering::Relaxed) {
                        if let Err(e) =
                            voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None)
                        {
                            crate::log_debug(&format!("SAPI5 Speak failed: {}", e));
                        }
                        return;
                    }
                    while let Ok(cmd) = command_rx.try_recv() {
                        match cmd {
                            TtsCommand::Pause => {
                                if let Some(audio) = &audio_output {
                                    if let Err(e) = audio.SetState(SPAS_PAUSE, 0) {
                                        crate::log_debug(&format!(
                                            "SAPI5 audio pause failed: {}",
                                            e
                                        ));
                                    }
                                    paused_by_purge = false;
                                } else if HARD_PAUSE_IMMEDIATE {
                                    let mut remainder: Option<String> = None;
                                    let text_wide: Vec<u16> =
                                        current_chunk.encode_utf16().collect();
                                    let mut start = if has_word_boundary {
                                        resolve_stream_pos_to_text_pos(
                                            &current_chunk,
                                            &ssml,
                                            last_word_stream_pos_utf16,
                                        )
                                    } else {
                                        0
                                    };
                                    if start == 0 {
                                        let mut status = SPVOICESTATUS::default();
                                        if voice
                                            .GetStatus(&mut status, std::ptr::null_mut())
                                            .is_ok()
                                        {
                                            let stream_pos = status.ulInputWordPos as usize;
                                            start = resolve_stream_pos_to_text_pos(
                                                &current_chunk,
                                                &ssml,
                                                stream_pos,
                                            );
                                        }
                                    }
                                    let start = start.min(text_wide.len());
                                    if start < text_wide.len() {
                                        let tail = String::from_utf16_lossy(&text_wide[start..]);
                                        if !tail.trim().is_empty() {
                                            remainder = Some(tail);
                                        }
                                    }
                                    if let Err(e) = voice.Speak(
                                        PCWSTR::null(),
                                        SPF_PURGEBEFORESPEAK.0 as u32,
                                        None,
                                    ) {
                                        crate::log_debug(&format!(
                                            "SAPI5 Speak purge failed (pause): {}",
                                            e
                                        ));
                                    }
                                    if let Some(rem) = remainder {
                                        pending.push_front(rem);
                                    } else {
                                        // Preserve audible behavior when offset is unavailable.
                                        pending.push_front(current_chunk.clone());
                                    }
                                    paused_by_purge = true;
                                } else if let Err(e) = voice.Pause() {
                                    crate::log_debug(&format!("SAPI5 Pause failed: {}", e));
                                    paused_by_purge = false;
                                } else {
                                    paused_by_purge = false;
                                }
                                paused = true;
                                break;
                            }
                            TtsCommand::Resume => {
                                if let Some(audio) = &audio_output {
                                    if let Err(e) = audio.SetState(SPAS_RUN, 0) {
                                        crate::log_debug(&format!(
                                            "SAPI5 audio resume failed: {}",
                                            e
                                        ));
                                    }
                                } else if !paused_by_purge && let Err(e) = voice.Resume() {
                                    crate::log_debug(&format!("SAPI5 Resume failed: {}", e));
                                }
                                paused_by_purge = false;
                                paused = false;
                            }
                            TtsCommand::Stop => {
                                cancel.store(true, Ordering::SeqCst);
                                if let Some(audio) = &audio_output
                                    && let Err(e) = audio.SetState(SPAS_STOP, 0)
                                {
                                    crate::log_debug(&format!("SAPI5 audio stop failed: {}", e));
                                }
                                if let Err(e) =
                                    voice.Speak(PCWSTR::null(), SPF_PURGEBEFORESPEAK.0 as u32, None)
                                {
                                    crate::log_debug(&format!(
                                        "SAPI5 Speak purge failed (stop): {}",
                                        e
                                    ));
                                }
                                return;
                            }
                        }
                    }
                    if paused {
                        std::thread::sleep(Duration::from_millis(PAUSED_POLL_MS as u64));
                        continue;
                    }

                    let mut status = SPVOICESTATUS::default();
                    if voice.GetStatus(&mut status, std::ptr::null_mut()).is_ok()
                        && status.dwRunningState == SPRS_DONE.0 as u32
                    {
                        break;
                    }
                    std::thread::sleep(Duration::from_millis(COMMAND_POLL_MS as u64));
                }
            }
        }
    });
    Ok(())
}

pub struct SapiExportOptions<'a> {
    pub chunks: &'a [String],
    pub voice_name: &'a str,
    pub output_path: &'a Path,
    pub language: Language,
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub audiobook_bitrate_kbps: u32,
    pub cancel: Arc<AtomicBool>,
}

pub struct SapiVoice {
    voice: ISpVoice,
}

impl SapiVoice {
    pub fn new() -> Result<Self, String> {
        let voice: ISpVoice = unsafe { CoCreateInstance(&SpVoice, None, CLSCTX_ALL) }
            .map_err(|e| format!("Failed to create SpVoice: {}", e))?;
        Ok(Self { voice })
    }
}

fn configure_sapi_voice(voice: &ISpVoice, options: &SapiExportOptions) -> Result<(), String> {
    let voice_token = find_voice_token(options.voice_name).ok_or_else(|| {
        "Selected SAPI voice not found. Please select a voice in Options.".to_string()
    })?;
    unsafe {
        voice
            .SetVoice(&voice_token)
            .map_err(|e| format!("SetVoice failed: {}", e))?;
    }
    crate::log_debug(&format!(
        "SAPI: voice set. voice_name={}",
        options.voice_name
    ));
    if let Err(e) = unsafe { voice.SetRate(map_sapi_rate(options.rate)) } {
        crate::log_debug(&format!("Failed to set SAPI5 rate: {}", e));
    }
    if let Err(e) = unsafe { voice.SetVolume(map_sapi_volume(options.volume)) } {
        crate::log_debug(&format!("Failed to set SAPI5 volume: {}", e));
    }
    Ok(())
}

fn speak_sapi5_to_file_via_bridge(
    options: &SapiExportOptions,
    mut progress_callback: impl FnMut(usize),
) -> Result<(), String> {
    let exe_path = select_sapi_bridge_exe()?;
    let mut child = Command::new(exe_path)
        .arg("--sapi5-output")
        .arg(options.output_path)
        .arg("--voice-name")
        .arg(options.voice_name)
        .arg("--rate")
        .arg(options.rate.to_string())
        .arg("--pitch")
        .arg(options.pitch.to_string())
        .arg("--volume")
        .arg(options.volume.to_string())
        .stdin(Stdio::piped())
        .creation_flags(0x08000000)
        .spawn()
        .map_err(|e| format!("Failed to spawn SAPI5 bridge: {}", e))?;

    if let Some(mut stdin) = child.stdin.take() {
        let text = options.chunks.join("\n");
        stdin
            .write_all(text.as_bytes())
            .map_err(|e| format!("Failed to send text to SAPI5 bridge: {}", e))?;
    }

    loop {
        if options.cancel.load(Ordering::SeqCst) {
            crate::log_if_err!(child.kill());
            return Err("Cancelled".to_string());
        }
        match child.try_wait() {
            Ok(Some(status)) => {
                if !status.success() {
                    return Err(format!("SAPI5 bridge failed with status: {}", status));
                }
                break;
            }
            Ok(None) => std::thread::sleep(Duration::from_millis(25)),
            Err(e) => return Err(format!("SAPI5 bridge wait failed: {}", e)),
        }
    }

    for _ in options.chunks {
        progress_callback(0);
    }
    Ok(())
}

fn speak_sapi_to_file_with_voice(
    voice: &ISpVoice,
    options: SapiExportOptions,
    mut progress_callback: impl FnMut(usize),
) -> Result<(), String> {
    if has_sapi5_bridge_voice(options.voice_name) && !has_native_sapi5_voice(options.voice_name) {
        crate::log_debug(&format!(
            "SAPI5 export: using 32-bit bridge voice fallback for '{}'",
            options.voice_name
        ));
        return speak_sapi5_to_file_via_bridge(&options, progress_callback);
    }

    unsafe {
        let is_mp3 = options
            .output_path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("mp3"));
        crate::log_debug(&format!(
            "SAPI: is_mp3={}, output_path={:?}",
            is_mp3, options.output_path
        ));
        configure_sapi_voice(voice, &options)?;

        let mut wfx = WAVEFORMATEX::default();
        wfx.wFormatTag = WAVE_FORMAT_PCM as u16;
        wfx.nChannels = 1;
        wfx.nSamplesPerSec = 44100;
        wfx.wBitsPerSample = 16;
        wfx.nBlockAlign = wfx.nChannels * (wfx.wBitsPerSample / 8);
        wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * (wfx.nBlockAlign as u32);
        wfx.cbSize = 0;

        if is_mp3 {
            let non_empty_chunks: Vec<(usize, &String)> = options
                .chunks
                .iter()
                .enumerate()
                .filter(|(_, c)| !c.trim().is_empty())
                .collect();

            if non_empty_chunks.is_empty() {
                let _unused_file =
                    std::fs::File::create(options.output_path).map_err(|e| e.to_string())?;
                return Ok(());
            }

            let temp_root = std::env::temp_dir().join(format!(
                "sonarpad_sapi5_mp3_{}_{}_{}",
                std::process::id(),
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_millis())
                    .unwrap_or(0),
                SAPI5_TMP_SEQ.fetch_add(1, Ordering::Relaxed)
            ));
            std::fs::create_dir_all(&temp_root).map_err(|e| e.to_string())?;

            let ff_settings = crate::ffmpeg_export::ConvertAudioSettings {
                format: crate::ffmpeg_export::ConvertAudioFormat::Mp3,
                quality: crate::ffmpeg_export::ConvertAudioQuality::BitrateKbps(
                    options.audiobook_bitrate_kbps,
                ),
            };
            let total_parts = non_empty_chunks.len() as u32;
            let mut converted_parts: Vec<PathBuf> = Vec::with_capacity(non_empty_chunks.len());
            let mut last_convert_pct = u32::MAX;

            for (part_idx, (chunk_idx, chunk)) in non_empty_chunks.iter().enumerate() {
                if options.cancel.load(Ordering::Relaxed) {
                    for p in &converted_parts {
                        if let Err(e) = std::fs::remove_file(p) {
                            crate::log_debug(&format!("Failed to remove temp MP3 part: {}", e));
                        }
                    }
                    if let Err(e) = std::fs::remove_dir_all(&temp_root) {
                        crate::log_debug(&format!("Failed to remove temp SAPI5 folder: {}", e));
                    }
                    return Err("Cancelled".to_string());
                }

                let part_no = part_idx + 1;
                let part_wav = temp_root.join(format!("part_{:06}.wav", part_no));
                let part_mp3 = temp_root.join(format!("part_{:06}.mp3", part_no));

                let stream: ISpStream = CoCreateInstance(&SpFileStream, None, CLSCTX_ALL)
                    .map_err(|e| format!("Failed to create SpFileStream: {}", e))?;
                let path_wide = to_wide(part_wav.to_str().ok_or("Invalid path")?);
                let mut bind_ok = false;
                let mut bind_err = String::new();
                for attempt in 1..=3 {
                    match stream.BindToFile(
                        PCWSTR(path_wide.as_ptr()),
                        SPFM_CREATE_ALWAYS,
                        Some(&SPDFID_WAVEFORMATEX),
                        Some(&wfx),
                        0,
                    ) {
                        Ok(()) => {
                            bind_ok = true;
                            break;
                        }
                        Err(e) => {
                            bind_err = e.to_string();
                            crate::log_debug(&format!(
                                "SAPI: BindToFile retry {}/3 failed: {}",
                                attempt, bind_err
                            ));
                            std::thread::sleep(Duration::from_millis(25));
                        }
                    }
                }
                if !bind_ok {
                    return Err(format!("BindToFile failed: {}", bind_err));
                }

                voice
                    .SetOutput(&stream, true)
                    .map_err(|e| format!("SetOutput failed: {}", e))?;

                crate::log_debug(&format!(
                    "SAPI: chunk start. idx={} len={}",
                    *chunk_idx + 1,
                    chunk.len()
                ));
                let ssml = mk_sapi_ssml(chunk, options.rate, options.pitch, options.volume);
                let chunk_wide = to_wide(&ssml);
                voice
                    .Speak(PCWSTR(chunk_wide.as_ptr()), SPF_IS_XML.0 as u32, None)
                    .map_err(|e| format!("Speak failed: {}", e))?;
                if let Err(e) = voice.WaitUntilDone(u32::MAX) {
                    crate::log_debug(&format!("Failed to wait for SAPI5: {}", e));
                }
                if let Err(e) = stream.Close() {
                    crate::log_debug(&format!("Failed to close SAPI5 stream: {}", e));
                }
                crate::log_debug(&format!("SAPI: chunk done. idx={}", *chunk_idx + 1));

                match crate::audio_utils::get_wav_data_size(&part_wav) {
                    Ok(0) => {
                        crate::log_debug("SAPI: WAV data size is 0, repairing WAV header sizes.");
                        match repair_wav_header_sizes_if_needed(&part_wav) {
                            Ok(true) => crate::log_debug("SAPI: WAV header repaired."),
                            Ok(false) => crate::log_debug(
                                "SAPI: WAV header repair skipped (layout mismatch).",
                            ),
                            Err(e) => {
                                crate::log_debug(&format!("SAPI: WAV header repair failed: {}", e))
                            }
                        }
                    }
                    Ok(_) => {}
                    Err(e) => crate::log_debug(&format!("SAPI: get_wav_data_size failed: {}", e)),
                }

                let mut ff_progress = |p: u32| {
                    let base = (part_idx as u32).saturating_mul(10000);
                    let overall_10000 = (base.saturating_add(p)).saturating_div(total_parts.max(1));
                    let pct = (overall_10000 / 100).min(100);
                    if pct != last_convert_pct {
                        last_convert_pct = pct;
                    }
                };

                if let Err(e) = crate::ffmpeg_export::convert_audio_file(
                    &part_wav,
                    &part_mp3,
                    &ff_settings,
                    None,
                    Some(&mut ff_progress),
                ) {
                    let dest_wav = options.output_path.with_extension("wav");
                    if let Err(rename_err) = std::fs::rename(&part_wav, &dest_wav) {
                        crate::log_debug(&format!(
                            "Failed to preserve SAPI5 WAV part: {}",
                            rename_err
                        ));
                    }
                    for p in &converted_parts {
                        if let Err(rem_err) = std::fs::remove_file(p) {
                            crate::log_debug(&format!(
                                "Failed to remove temp MP3 part: {}",
                                rem_err
                            ));
                        }
                    }
                    if let Err(rem_err) = std::fs::remove_dir_all(&temp_root) {
                        crate::log_debug(&format!(
                            "Failed to remove temp SAPI5 folder: {}",
                            rem_err
                        ));
                    }
                    return Err(mf_error_message(options.language, &e));
                }

                if let Err(e) = std::fs::remove_file(&part_wav) {
                    crate::log_debug(&format!("Failed to remove SAPI5 temp WAV: {}", e));
                }

                converted_parts.push(part_mp3);
                progress_callback(*chunk_idx + 1);
            }

            let mut out_file =
                std::fs::File::create(options.output_path).map_err(|e| e.to_string())?;
            for path in &converted_parts {
                let mut in_file = std::fs::File::open(path).map_err(|e| e.to_string())?;
                std::io::copy(&mut in_file, &mut out_file).map_err(|e| e.to_string())?;
            }

            for p in &converted_parts {
                if let Err(e) = std::fs::remove_file(p) {
                    crate::log_debug(&format!("Failed to remove temp MP3 part: {}", e));
                }
            }
            if let Err(e) = std::fs::remove_dir_all(&temp_root) {
                crate::log_debug(&format!("Failed to remove temp SAPI5 folder: {}", e));
            }
        } else {
            let wav_path = options.output_path.to_path_buf();
            crate::log_debug(&format!("SAPI: Target wav_path={:?}", wav_path));
            let stream: ISpStream = CoCreateInstance(&SpFileStream, None, CLSCTX_ALL)
                .map_err(|e| format!("Failed to create SpFileStream: {}", e))?;
            let path_wide = to_wide(wav_path.to_str().ok_or("Invalid path")?);
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

            crate::log_debug(&format!(
                "SAPI: entering chunk loop. chunks={}",
                options.chunks.len()
            ));
            for (i, chunk) in options.chunks.iter().enumerate() {
                if options.cancel.load(Ordering::Relaxed) {
                    if let Err(e) = stream.Close() {
                        crate::log_debug(&format!("Failed to close SAPI5 stream: {}", e));
                    }
                    if let Err(e) = std::fs::remove_file(&wav_path) {
                        crate::log_debug(&format!("Failed to remove SAPI5 temp WAV: {}", e));
                    }
                    return Err("Cancelled".to_string());
                }
                if chunk.trim().is_empty() {
                    continue;
                }
                crate::log_debug(&format!(
                    "SAPI: chunk start. idx={} len={}",
                    i + 1,
                    chunk.len()
                ));
                let ssml = mk_sapi_ssml(chunk, options.rate, options.pitch, options.volume);
                let chunk_wide = to_wide(&ssml);
                voice
                    .Speak(PCWSTR(chunk_wide.as_ptr()), SPF_IS_XML.0 as u32, None)
                    .map_err(|e| format!("Speak failed: {}", e))?;
                crate::log_debug(&format!("SAPI: chunk done. idx={}", i + 1));
                progress_callback(i + 1);
            }
            if let Err(e) = voice.WaitUntilDone(u32::MAX) {
                crate::log_debug(&format!("Failed to wait for SAPI5: {}", e));
            }
            if let Err(e) = stream.Close() {
                crate::log_debug(&format!("Failed to close SAPI5 stream: {}", e));
            }
        }

        Ok(())
    }
}

pub fn speak_sapi_to_file(
    options: SapiExportOptions,
    mut progress_callback: impl FnMut(usize),
) -> Result<(), String> {
    let _com = ComGuard::new_sta().map_err(|e| format!("CoInitializeEx failed: {}", e))?;
    let thread_id = std::thread::current().id();
    crate::log_debug(&format!(
        "SAPI5 export start. thread={:?} chunks={} voice={} rate={} pitch={} volume={} out={:?}",
        thread_id,
        options.chunks.len(),
        options.voice_name,
        options.rate,
        options.pitch,
        options.volume,
        options.output_path
    ));
    let voice = SapiVoice::new()?;
    speak_sapi_to_file_with_voice(&voice.voice, options, &mut progress_callback)
}

pub fn speak_sapi_to_file_with_sapi_voice(
    voice: &SapiVoice,
    options: SapiExportOptions,
    progress_callback: impl FnMut(usize),
) -> Result<(), String> {
    speak_sapi_to_file_with_voice(&voice.voice, options, progress_callback)
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

fn map_ssml_pos_to_text_pos(ssml: &str, ssml_pos_utf16: usize) -> usize {
    let mut ssml_units = 0usize;
    let mut text_units = 0usize;
    let mut in_tag = false;
    let mut in_entity = false;

    for ch in ssml.chars() {
        let ch_units = ch.len_utf16();
        if ssml_units >= ssml_pos_utf16 {
            break;
        }
        ssml_units = ssml_units.saturating_add(ch_units);

        if in_tag {
            if ch == '>' {
                in_tag = false;
            }
            continue;
        }

        if in_entity {
            if ch == ';' {
                // XML entity contributes exactly one text character.
                text_units = text_units.saturating_add(1);
                in_entity = false;
            }
            continue;
        }

        if ch == '<' {
            in_tag = true;
            continue;
        }
        if ch == '&' {
            in_entity = true;
            continue;
        }

        text_units = text_units.saturating_add(ch_units);
    }

    text_units
}

fn resolve_stream_pos_to_text_pos(text: &str, ssml: &str, stream_pos_utf16: usize) -> usize {
    let text_len = text.encode_utf16().count();
    if stream_pos_utf16 == 0 {
        return 0;
    }
    if stream_pos_utf16 <= text_len {
        return stream_pos_utf16;
    }
    map_ssml_pos_to_text_pos(ssml, stream_pos_utf16).min(text_len)
}

fn word_boundary_interest_mask() -> u64 {
    const SPFEI_FLAGCHECK: u64 = (1u64 << 30) | (1u64 << 33);
    let shift = SPEI_WORD_BOUNDARY.0;
    if (0..64).contains(&shift) {
        (1u64 << (shift as u32)) | SPFEI_FLAGCHECK
    } else {
        0
    }
}

fn spevent_event_id(event: &SPEVENT) -> i32 {
    (event._bitfield as u32 & 0xFFFF) as i32
}

fn drain_word_boundary_events(source: &ISpEventSource, last_text_pos_utf16: &mut usize) -> bool {
    let mut saw_word_boundary = false;
    unsafe {
        loop {
            let mut event = SPEVENT::default();
            let mut fetched = 0u32;
            if let Err(e) =
                source.GetEvents(1, &mut event as *mut SPEVENT, &mut fetched as *mut u32)
            {
                crate::log_debug(&format!("Failed to read SAPI5 events: {}", e));
                break;
            }
            if fetched == 0 {
                break;
            }
            if spevent_event_id(&event) == SPEI_WORD_BOUNDARY.0 {
                // SAPI word-boundary semantics:
                // - lParam = input word position
                // - wParam = input word length
                *last_text_pos_utf16 = event.lParam.0.max(0) as usize;
                saw_word_boundary = true;
            }
        }
    }
    saw_word_boundary
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
