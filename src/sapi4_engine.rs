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

pub fn speak(voice_index: i32, text: &str) {
    if let Ok(mut exe_path) = std::env::current_exe() {
        exe_path.set_file_name("sapi4_bridge.exe");
        if let Ok(mut child) = Command::new(exe_path)
            .arg("--voice")
            .arg(voice_index.to_string())
            .stdin(Stdio::piped())
            .creation_flags(0x08000000)
            .spawn()
        {
            if let Some(mut stdin) = child.stdin.take() {
                let _ = stdin.write_all(text.as_bytes());
            }
        }
    }
}

pub fn stop() {
    let _ = Command::new("taskkill")
        .args(&["/F", "/IM", "sapi4_bridge.exe", "/T"])
        .creation_flags(0x08000000)
        .status();
}

fn spawn_sapi4_process(voice_index: i32, text: &str) -> Result<(), String> {
    let mut exe_path = std::env::current_exe().map_err(|e| e.to_string())?;
    exe_path.set_file_name("sapi4_bridge.exe");
    let mut child = Command::new(exe_path)
        .arg("--voice")
        .arg(voice_index.to_string())
        .stdin(Stdio::piped())
        .creation_flags(0x08000000)
        .spawn()
        .map_err(|e| e.to_string())?;
    if let Some(mut stdin) = child.stdin.take() {
        stdin
            .write_all(text.as_bytes())
            .map_err(|e| e.to_string())?;
    }
    let _ = child.wait();
    Ok(())
}

pub fn speak_sapi4_to_file(
    chunks: &[String],
    voice_index: i32,
    output: &Path,
    cancel: Arc<AtomicBool>,
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

    let mut exe_path = std::env::current_exe().map_err(|e| e.to_string())?;
    exe_path.set_file_name("sapi4_bridge.exe");
    let mut child = Command::new(exe_path)
        .arg("--voice")
        .arg(voice_index.to_string())
        .arg("--output")
        .arg(&wav_path)
        .stdin(Stdio::piped())
        .creation_flags(0x08000000)
        .spawn()
        .map_err(|e| e.to_string())?;
    if let Some(mut stdin) = child.stdin.take() {
        stdin
            .write_all(text.as_bytes())
            .map_err(|e| e.to_string())?;
    }
    loop {
        if cancel.load(Ordering::SeqCst) {
            let _ = child.kill();
            return Err("Cancelled".to_string());
        }
        match child.try_wait() {
            Ok(Some(_)) => break,
            Ok(None) => std::thread::sleep(std::time::Duration::from_millis(50)),
            Err(e) => return Err(e.to_string()),
        }
    }

    for idx in 0..chunks.len() {
        on_progress(idx);
    }

    if is_mp3 {
        if let Ok(data_size) = crate::audio_utils::get_wav_data_size(&wav_path) {
            if data_size == 0 {
                let _ = crate::audio_utils::write_silence_file(&wav_path, 44100, 2, 16, 500);
            }
        }
        if let Err(e) = crate::mf_encoder::encode_wav_to_mp3(&wav_path, output) {
            let dest_wav = output.with_extension("wav");
            let _ = std::fs::rename(&wav_path, &dest_wav);
            return Err(e);
        }
    }

    Ok(())
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
        if line.starts_with("VOICE:") {
            let parts: Vec<&str> = line[6..].split('|').collect();
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
