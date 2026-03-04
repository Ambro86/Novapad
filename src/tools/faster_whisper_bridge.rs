use crate::settings::Language;
use reqwest::blocking::Client;
use reqwest::header::USER_AGENT;
use serde::Deserialize;
use std::io::{BufRead, BufReader, Read};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc;
use std::time::Duration;
use std::{fs, io::Write};

const CREATE_NO_WINDOW: u32 = 0x0800_0000;
const BRIDGE_FILE_NAME: &str = "faster_whisper_bridge.exe";
const BRIDGE_DOWNLOAD_URL: &str =
    "https://raw.githubusercontent.com/Ambro86/Sonarpad/master/dll/faster_whisper_bridge.exe";

#[derive(Clone, Copy)]
pub enum BridgeModel {
    Small,
    Medium,
    LargeV3,
}

impl BridgeModel {
    fn as_name(self) -> &'static str {
        match self {
            BridgeModel::Small => "small",
            BridgeModel::Medium => "medium",
            BridgeModel::LargeV3 => "large-v3",
        }
    }
}

#[derive(Deserialize)]
struct BridgeOutput {
    ok: bool,
    #[serde(default)]
    text: String,
    #[serde(default)]
    error: String,
}

fn handle_bridge_line(
    line: &str,
    progress: &mut Option<Box<dyn FnMut(i32) + Send>>,
    bridge_result: &mut Option<BridgeOutput>,
) {
    if let Some(raw_pct) = line.strip_prefix("PROGRESS:")
        && let Ok(pct) = raw_pct.trim().parse::<i32>()
    {
        if let Some(cb) = progress.as_mut() {
            cb(pct.clamp(0, 100));
        }
        return;
    }
    if line.starts_with('{')
        && let Ok(parsed) = serde_json::from_str::<BridgeOutput>(line)
    {
        *bridge_result = Some(parsed);
    }
}

fn language_code(language: Language) -> &'static str {
    match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::Swedish => "sv",
        Language::Vietnamese => "vi",
        Language::Czech => "cs",
        Language::Polish => "pl",
        Language::French => "fr",
        Language::Serbian => "sr",
        Language::Ukrainian => "uk",
        Language::Lithuanian => "lt",
        Language::Chinese => "zh",
    }
}

fn bridge_install_path() -> PathBuf {
    crate::settings::settings_dir()
        .join("tools")
        .join(BRIDGE_FILE_NAME)
}

fn download_bridge_binary(target_path: &Path, cancel: &Arc<AtomicBool>) -> Result<(), String> {
    let parent = target_path
        .parent()
        .ok_or_else(|| "invalid bridge target path".to_string())?;
    fs::create_dir_all(parent).map_err(|e| format!("create tools dir failed: {e}"))?;

    let client = Client::builder()
        .connect_timeout(Duration::from_secs(30))
        .timeout(Duration::from_secs(60 * 15))
        .build()
        .map_err(|e| format!("http client build failed: {e}"))?;

    let mut response = client
        .get(BRIDGE_DOWNLOAD_URL)
        .header(USER_AGENT, "Sonarpad/faster-whisper-bridge")
        .send()
        .map_err(|e| format!("bridge download request failed: {e}"))?;
    if !response.status().is_success() {
        return Err(format!(
            "bridge download failed: HTTP {}",
            response.status()
        ));
    }

    let part_path = target_path.with_extension("exe.part");
    let mut out =
        fs::File::create(&part_path).map_err(|e| format!("create bridge .part failed: {e}"))?;
    let mut buf = vec![0u8; 128 * 1024];
    loop {
        if cancel.load(Ordering::Relaxed) {
            if let Err(err) = fs::remove_file(&part_path) {
                crate::log_debug(&format!("bridge .part cleanup failed: {err}"));
            }
            return Err("cancelled".to_string());
        }
        let read = response
            .read(&mut buf)
            .map_err(|e| format!("bridge download read failed: {e}"))?;
        if read == 0 {
            break;
        }
        out.write_all(&buf[..read])
            .map_err(|e| format!("bridge write failed: {e}"))?;
    }
    out.flush()
        .map_err(|e| format!("bridge flush failed: {e}"))?;
    fs::rename(&part_path, target_path).map_err(|e| format!("bridge finalize failed: {e}"))?;
    Ok(())
}

fn ensure_bridge_binary(cancel: &Arc<AtomicBool>) -> Result<PathBuf, String> {
    let bridge_path = bridge_install_path();
    if bridge_path.exists()
        && let Ok(meta) = fs::metadata(&bridge_path)
        && meta.len() > 0
    {
        return Ok(bridge_path);
    }
    crate::log_debug(&format!(
        "Transcription: bridge missing, downloading from {}",
        BRIDGE_DOWNLOAD_URL
    ));
    download_bridge_binary(&bridge_path, cancel)?;
    crate::log_debug(&format!(
        "Transcription: bridge ready at {}",
        bridge_path.display()
    ));
    Ok(bridge_path)
}

fn spawn_bridge_process(
    bridge_path: &Path,
    wav_path: &Path,
    model: BridgeModel,
    language: Language,
    work_dir: &Path,
) -> Result<std::process::Child, String> {
    let model_name = model.as_name();
    let lang = language_code(language);
    let download_root = crate::settings::settings_dir()
        .join("models")
        .join("faster-whisper");
    let args = vec![
        "--input".to_string(),
        wav_path.display().to_string(),
        "--model".to_string(),
        model_name.to_string(),
        "--language".to_string(),
        lang.to_string(),
        "--download-root".to_string(),
        download_root.display().to_string(),
    ];

    let mut command = Command::new(bridge_path);
    command.args(&args);

    command
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .creation_flags(CREATE_NO_WINDOW)
        .current_dir(work_dir)
        .spawn()
        .map_err(|e| format!("start bridge failed: {e}"))
}

fn select_and_spawn_bridge(
    wav_path: &Path,
    model: BridgeModel,
    language: Language,
    cancel: &Arc<AtomicBool>,
) -> Result<std::process::Child, String> {
    let work_dir = crate::settings::settings_dir();
    let bridge_path = ensure_bridge_binary(cancel)?;
    spawn_bridge_process(&bridge_path, wav_path, model, language, &work_dir)
}

pub fn transcribe_wav(
    wav_path: &Path,
    model: BridgeModel,
    language: Language,
    cancel: &Arc<AtomicBool>,
    mut progress: Option<Box<dyn FnMut(i32) + Send>>,
) -> Result<String, String> {
    let mut child = select_and_spawn_bridge(wav_path, model, language, cancel)?;

    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "bridge stdout unavailable".to_string())?;
    let stderr = child
        .stderr
        .take()
        .ok_or_else(|| "bridge stderr unavailable".to_string())?;

    let (line_tx, line_rx) = mpsc::channel::<String>();
    let stdout_thread = std::thread::spawn(move || {
        let mut reader = BufReader::new(stdout);
        let mut line = Vec::new();
        loop {
            line.clear();
            match reader.read_until(b'\n', &mut line) {
                Ok(0) => break,
                Ok(_) => {
                    let trimmed = String::from_utf8_lossy(&line).trim().to_string();
                    if !trimmed.is_empty() && line_tx.send(trimmed).is_err() {
                        break;
                    }
                }
                Err(err) => {
                    let _ignore = line_tx.send(format!("ERROR:stdout_read:{err}"));
                    break;
                }
            }
        }
    });

    let (err_tx, err_rx) = mpsc::channel::<String>();
    let stderr_thread = std::thread::spawn(move || {
        let mut reader = BufReader::new(stderr);
        let mut text = String::new();
        if let Err(err) = reader.read_to_string(&mut text) {
            text = format!("stderr read failed: {err}");
        }
        let _ignore = err_tx.send(text);
    });

    let mut bridge_result: Option<BridgeOutput> = None;
    let mut cancelled = false;

    loop {
        if cancel.load(Ordering::Relaxed) {
            cancelled = true;
            if let Err(err) = child.kill() {
                crate::log_debug(&format!("bridge kill failed: {err}"));
            }
        }

        match line_rx.recv_timeout(Duration::from_millis(150)) {
            Ok(line) => {
                handle_bridge_line(&line, &mut progress, &mut bridge_result);
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {}
            Err(mpsc::RecvTimeoutError::Disconnected) => {}
        }

        match child.try_wait() {
            Ok(Some(_)) => break,
            Ok(None) => {}
            Err(err) => return Err(format!("bridge wait failed: {err}")),
        }
    }

    if let Err(_panic) = stdout_thread.join() {
        return Err("bridge stdout thread panicked".to_string());
    }
    while let Ok(line) = line_rx.try_recv() {
        handle_bridge_line(&line, &mut progress, &mut bridge_result);
    }
    if let Err(_panic) = stderr_thread.join() {
        return Err("bridge stderr thread panicked".to_string());
    }
    let stderr_text = err_rx.recv().unwrap_or_default();

    if cancelled || cancel.load(Ordering::Relaxed) {
        return Err("cancelled".to_string());
    }

    if let Some(result) = bridge_result {
        if result.ok {
            return Ok(result.text);
        }
        if !result.error.trim().is_empty() {
            return Err(result.error);
        }
    }

    if stderr_text.trim().is_empty() {
        Err("bridge returned no transcript".to_string())
    } else {
        Err(stderr_text)
    }
}
