use crate::settings::VoiceInfo;
use crate::tts_engine::TtsCommand;
use once_cell::sync::Lazy;
use std::io::Write;
use std::os::windows::process::CommandExt;
use std::path::Path;
use std::path::PathBuf;
use std::process::{ChildStdin, Command, Stdio};
use std::sync::Arc;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use tokio::sync::mpsc;

static VOICE_CACHE: Lazy<Mutex<Vec<VoiceInfo>>> = Lazy::new(|| Mutex::new(Vec::new()));

pub struct Sapi4Options {
    pub rate: i32,
    pub pitch: i32,
    pub volume: i32,
    pub cancel: Arc<AtomicBool>,
}

pub fn speak_sapi4_to_file(
    chunks: &[String],
    voice_index: i32,
    output: &Path,
    options: Sapi4Options,
    mut on_progress: impl FnMut(usize),
) -> Result<(), String> {
    if chunks.is_empty() {
        return Ok(());
    }
    let text = chunks.join("\n");
    let is_mp3 = output
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("mp3"))
        .unwrap_or(false);
    let wav_path = if is_mp3 {
        output.with_extension("wav")
    } else {
        output.to_path_buf()
    };
    let exe_path = select_sapi4_bridge_for_file()?;
    let mut child = Command::new(exe_path)
        .arg("--voice")
        .arg(voice_index.to_string())
        .arg("--rate")
        .arg(options.rate.to_string())
        .arg("--pitch")
        .arg(options.pitch.to_string())
        .arg("--volume")
        .arg(options.volume.to_string())
        .arg("--output")
        .arg(output)
        .stdin(Stdio::piped())
        .creation_flags(0x08000000)
        .spawn()
        .map_err(|e| e.to_string())?;
    if let Some(mut stdin) = child.stdin.take() {
        stdin
            .write_all(text.as_bytes())
            .map_err(|e| e.to_string())?;
    }
    let mut reported = 0usize;
    let max_report = chunks.len().saturating_sub(1);
    let mut last_size = 0u64;
    loop {
        if options.cancel.load(Ordering::SeqCst) {
            let _ = child.kill();
            return Err("Cancelled".to_string());
        }
        let size = std::fs::metadata(&wav_path).map(|m| m.len()).unwrap_or(0);
        if size > last_size {
            last_size = size;
            if reported < max_report {
                reported += 1;
                on_progress(0);
            }
        }
        match child.try_wait() {
            Ok(Some(_)) => break,
            Ok(None) => std::thread::sleep(std::time::Duration::from_millis(50)),
            Err(e) => return Err(e.to_string()),
        }
    }

    for _ in reported..chunks.len() {
        on_progress(0);
    }

    Ok(())
}

fn select_sapi4_bridge_for_file() -> Result<PathBuf, String> {
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
    let mut fallback = exe_path;
    fallback.set_file_name("sapi4_bridge.exe");
    Ok(fallback)
}

fn send_line(stdin: &mut ChildStdin, line: &str) -> std::io::Result<()> {
    stdin.write_all(line.as_bytes())?;
    stdin.write_all(b"\n")?;
    stdin.flush()
}

fn send_speak(stdin: &mut ChildStdin, text: &str) -> std::io::Result<()> {
    let bytes = text.as_bytes();
    stdin.write_all(format!("SPEAK {}\n", bytes.len()).as_bytes())?;
    stdin.write_all(bytes)?;
    stdin.flush()
}

pub fn play_sapi4(
    voice_index: i32,
    text: String,
    tts_rate: i32,
    tts_pitch: i32,
    tts_volume: i32,
    cancel: Arc<AtomicBool>,
    mut command_rx: mpsc::UnboundedReceiver<TtsCommand>,
) {
    std::thread::spawn(move || {
        let mut exe_path = match std::env::current_exe() {
            Ok(path) => path,
            Err(_) => return,
        };
        exe_path.set_file_name("sapi4_bridge.exe");
        let mut child = match Command::new(exe_path)
            .arg("--voice")
            .arg(voice_index.to_string())
            .arg("--rate")
            .arg(tts_rate.to_string())
            .arg("--pitch")
            .arg(tts_pitch.to_string())
            .arg("--volume")
            .arg(tts_volume.to_string())
            .arg("--server")
            .stdin(Stdio::piped())
            .creation_flags(0x08000000)
            .spawn()
        {
            Ok(child) => child,
            Err(_) => return,
        };
        let mut stdin = match child.stdin.take() {
            Some(stdin) => stdin,
            None => return,
        };
        if send_speak(&mut stdin, &text).is_err() {
            let _ = child.wait();
            return;
        }

        loop {
            if cancel.load(Ordering::Relaxed) {
                let _ = send_line(&mut stdin, "STOP");
                break;
            }

            match command_rx.blocking_recv() {
                Some(TtsCommand::Pause) => {
                    let _ = send_line(&mut stdin, "PAUSE");
                }
                Some(TtsCommand::Resume) => {
                    let _ = send_line(&mut stdin, "RESUME");
                }
                Some(TtsCommand::Stop) => {
                    let _ = send_line(&mut stdin, "STOP");
                    break;
                }
                None => {
                    let _ = send_line(&mut stdin, "STOP");
                    break;
                }
            }
        }

        let _ = child.wait();
    });
}

fn cache_path() -> Option<PathBuf> {
    let mut path = std::env::current_exe().ok()?;
    path.set_file_name("sapi4_voices.cache");
    Some(path)
}

fn parse_voices(output: &str) -> Vec<VoiceInfo> {
    let mut voices = Vec::new();
    for line in output.lines() {
        if let Some(stripped) = line.strip_prefix("VOICE:") {
            let parts: Vec<&str> = stripped.split('|').collect();
            if parts.len() == 2 {
                voices.push(VoiceInfo {
                    short_name: format!("SAPI4#{}|{}", parts[0], parts[1]),
                    locale: if parts[1].contains("Italian") {
                        "it-IT"
                    } else {
                        "en-US"
                    }
                    .to_string(),
                    is_multilingual: false,
                });
            }
        }
    }
    voices
}

pub fn get_voices() -> Vec<VoiceInfo> {
    let mut cache = VOICE_CACHE.lock().unwrap();
    if !cache.is_empty() {
        return cache.clone();
    }

    if let Some(path) = cache_path() {
        if let Ok(cached) = std::fs::read_to_string(&path) {
            let voices = parse_voices(&cached);
            if !voices.is_empty() {
                *cache = voices.clone();
                return voices;
            }
        }
    }

    if let Ok(mut exe_path) = std::env::current_exe() {
        exe_path.set_file_name("sapi4_bridge.exe");
        if let Ok(out) = Command::new(exe_path)
            .arg("--list")
            .creation_flags(0x08000000)
            .output()
        {
            let stdout = String::from_utf8_lossy(&out.stdout);
            let voices = parse_voices(&stdout);
            if !voices.is_empty() {
                if let Some(path) = cache_path() {
                    let _ = std::fs::write(path, stdout.as_bytes());
                }
                *cache = voices.clone();
                return voices;
            }
        }
    }
    Vec::new()
}
