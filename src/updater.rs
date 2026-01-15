#![allow(clippy::io_other_error)]
use std::fs::{File, OpenOptions};
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;

use serde::Deserialize;
use windows::Win32::Foundation::{CloseHandle, HWND, LPARAM, WPARAM};
use windows::Win32::Storage::FileSystem::{
    CreateFileW, FILE_ATTRIBUTE_NORMAL, FILE_GENERIC_READ, FILE_GENERIC_WRITE, FILE_SHARE_DELETE,
    FILE_SHARE_READ, FILE_SHARE_WRITE, FlushFileBuffers, GetDiskFreeSpaceExW,
    MOVEFILE_DELAY_UNTIL_REBOOT, MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH, MoveFileExW,
    OPEN_EXISTING, REPLACEFILE_WRITE_THROUGH, ReplaceFileW,
};
use windows::Win32::System::Threading::{OpenProcess, PROCESS_SYNCHRONIZE, WaitForSingleObject};
use windows::Win32::UI::Input::KeyboardAndMouse::{SetActiveWindow, SetFocus};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    FindWindowW, IDYES, MB_ICONERROR, MB_ICONINFORMATION, MB_ICONQUESTION, MB_OK, MB_SETFOREGROUND,
    MB_YESNO, MESSAGEBOX_STYLE, MessageBoxW, PostMessageW, SW_SHOW, SetForegroundWindow,
    ShowWindow, WM_CLOSE,
};
use windows::core::PCWSTR;

use crate::accessibility::to_wide;
use crate::i18n;
use crate::log_debug;
use crate::settings::{Language, load_settings};
use crate::with_state;

const REPO_OWNER: &str = "Ambro86";
const REPO_NAME: &str = "Novapad";
const USER_AGENT: &str = "NovapadUpdater";
const DIRECT_DOWNLOAD_URL: &str =
    "https://github.com/Ambro86/Novapad/releases/latest/download/novapad.exe";
const MIN_FREE_SPACE_BYTES: u64 = 5 * 1024 * 1024;
const UPDATE_LOCK_NAME: &str = "novapad.update.lock";
const UPDATER_DIR_NAME: &str = "Novapad\\updater";

const EXIT_OK: i32 = 0;
const EXIT_INTEGRITY_FAILED: i32 = 3;
const EXIT_DOWNLOAD_FAILED: i32 = 4;
const EXIT_UAC_CANCELLED: i32 = 5;
const EXIT_REPLACE_FAILED: i32 = 6;
const EXIT_REBOOT_REQUIRED: i32 = 10;

fn app_language(hwnd: HWND) -> Language {
    unsafe { with_state(hwnd, |state| state.settings.language).unwrap_or_default() }
}

#[derive(Deserialize)]
struct ReleaseAsset {
    name: String,
    browser_download_url: String,
    size: u64,
}

#[derive(Deserialize)]
struct ReleaseInfo {
    tag_name: String,
    assets: Vec<ReleaseAsset>,
}

pub(crate) fn check_for_update(hwnd: HWND, interactive: bool) {
    let language = app_language(hwnd);
    thread::spawn(move || {
        let current_version = env!("CARGO_PKG_VERSION");
        let latest = match fetch_latest_release() {
            Ok(info) => info,
            Err(err) => {
                log_debug(&format!("Update check failed: {err}"));
                if interactive {
                    show_update_error(language, UpdateError::Network);
                }
                return;
            }
        };

        let latest_version = normalize_version(&latest.tag_name);
        if !is_newer_version(&latest_version, current_version) {
            if interactive {
                show_update_info(language, UpdateInfo::NoUpdate);
            }
            return;
        }

        let Some(asset) = select_portable_asset(&latest.assets) else {
            log_debug("Update check: no portable asset found.");
            if interactive {
                show_update_error(language, UpdateError::NoPortableAsset);
            }
            return;
        };
        let sha_asset = select_sha256_asset(&latest.assets, &asset.name);

        let current_exe = match std::env::current_exe() {
            Ok(path) => path,
            Err(err) => {
                log_debug(&format!("Update check: current exe not available: {err}"));
                if interactive {
                    show_update_error_with_url(
                        language,
                        "updater.error.access",
                        DIRECT_DOWNLOAD_URL,
                    );
                }
                return;
            }
        };
        if let Err(err) = probe_dir_writable(&current_exe) {
            log_debug(&format!(
                "Update check: exe dir not writable: {err} class={}",
                classify_io_error(&err)
            ));
        }
        if let Err(err) = can_open_exe_for_update(&current_exe) {
            log_debug(&format!(
                "Update check: exe open diagnostic failed: {err} class={}",
                classify_io_error(&err)
            ));
        }
        let mut update_lock = match acquire_update_lock(&current_exe) {
            Ok(lock) => lock,
            Err(UpdateLockError::InProgress) => {
                if interactive {
                    show_update_error(language, UpdateError::Concurrent);
                }
                return;
            }
            Err(UpdateLockError::Other(err)) => {
                log_debug(&format!("Update check: cannot create update lock: {err}"));
                if interactive {
                    show_update_error_with_url(
                        language,
                        "updater.error.access",
                        DIRECT_DOWNLOAD_URL,
                    );
                }
                return;
            }
        };
        if asset.size > 0 {
            let required = asset.size.saturating_add(MIN_FREE_SPACE_BYTES);
            match available_disk_bytes(&current_exe) {
                Ok(available) => {
                    if available < required {
                        if interactive {
                            let needed = format_mb(required);
                            let available = format_mb(available);
                            show_update_error_args(
                                language,
                                "updater.error.space",
                                &[("needed", &needed), ("available", &available)],
                            );
                        }
                        return;
                    }
                }
                Err(err) => {
                    log_debug(&format!("Update check: disk space check failed: {err}"));
                }
            }
        }

        if !prompt_update(hwnd, language, current_version, &latest_version) {
            return;
        }

        match download_and_update(
            hwnd,
            language,
            &asset.browser_download_url,
            asset.size,
            sha_asset.map(|asset| asset.browser_download_url.as_str()),
            &asset.name,
        ) {
            Ok(UpdateAction::Started) => {
                update_lock.keep();
            }
            Ok(UpdateAction::Deferred) => {}
            Err(err) => {
                log_debug(&format!("Update failed: {err}"));
                if interactive {
                    show_update_error(language, UpdateError::Download);
                }
            }
        }
    });
}

fn fetch_latest_release() -> Result<ReleaseInfo, String> {
    let url = format!("https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest");
    let client = reqwest::blocking::Client::builder()
        .user_agent(USER_AGENT)
        .build()
        .map_err(|err| err.to_string())?;
    let resp = client
        .get(url)
        .send()
        .map_err(|err| err.to_string())?
        .error_for_status()
        .map_err(|err| err.to_string())?;
    resp.json().map_err(|err| err.to_string())
}

fn select_portable_asset(assets: &[ReleaseAsset]) -> Option<&ReleaseAsset> {
    assets.iter().find(|asset| {
        let name = asset.name.to_lowercase();
        name.ends_with(".exe") && !name.contains("setup") && !name.contains(".msi")
    })
}

fn select_sha256_asset<'a>(assets: &'a [ReleaseAsset], exe_name: &str) -> Option<&'a ReleaseAsset> {
    assets.iter().find(|asset| {
        let name = asset.name.to_lowercase();
        (name.ends_with(".sha256") || name.ends_with(".sha256.txt"))
            && name.contains(&exe_name.to_lowercase())
    })
}

fn normalize_version(tag: &str) -> String {
    tag.trim().trim_start_matches('v').to_string()
}

fn parse_version(value: &str) -> Option<(u32, u32, u32)> {
    let mut parts = value.split('.').take(3);
    let major = parts.next()?.parse().ok()?;
    let minor = parts.next().unwrap_or("0").parse().ok()?;
    let patch = parts.next().unwrap_or("0").parse().ok()?;
    Some((major, minor, patch))
}

fn is_newer_version(latest: &str, current: &str) -> bool {
    let Some(latest) = parse_version(latest) else {
        return false;
    };
    let Some(current) = parse_version(current) else {
        return false;
    };
    latest > current
}

fn prompt_update(hwnd: HWND, language: Language, current: &str, latest: &str) -> bool {
    let text = i18n::tr_f(
        language,
        "updater.prompt",
        &[("current", current), ("latest", latest)],
    );
    let title = i18n::tr(language, "updater.title");
    let result = show_update_message(hwnd, &text, &title, MB_YESNO | MB_ICONQUESTION);
    result == IDYES
}

fn prompt_restart_after_download(hwnd: HWND, language: Language) -> bool {
    let text = i18n::tr(language, "updater.prompt.restart");
    let title = i18n::tr(language, "updater.title");
    let result = show_update_message(hwnd, &text, &title, MB_YESNO | MB_ICONQUESTION);
    result == IDYES
}

fn prompt_pending_update(hwnd: HWND, language: Language) -> bool {
    let text = i18n::tr(language, "updater.prompt.pending");
    let title = i18n::tr(language, "updater.title");
    let result = show_update_message(hwnd, &text, &title, MB_YESNO | MB_ICONQUESTION);
    result == IDYES
}

enum UpdateAction {
    Started,
    Deferred,
}

fn download_and_update(
    hwnd: HWND,
    language: Language,
    url: &str,
    expected_size: u64,
    sha_url: Option<&str>,
    asset_name: &str,
) -> Result<UpdateAction, String> {
    let current_exe = std::env::current_exe().map_err(|err| err.to_string())?;
    let temp_path = temp_update_path(&current_exe)?;

    log_debug(&format!(
        "Download start: url={} temp={} expected_size={expected_size}",
        url,
        temp_path.display()
    ));
    download_file(url, &temp_path).map_err(|err| {
        log_debug(&format!("Download failed: {err}"));
        err
    })?;

    let mut expected_hash = None;
    if let Some(sha_url) = sha_url {
        log_debug(&format!("Download sha256: url={sha_url}"));
        expected_hash = download_sha256_optional(sha_url, asset_name);
        if expected_hash.is_some() {
            log_debug("Sha256 found for asset.");
        } else {
            log_debug("Sha256 file missing or no match.");
        }
    }

    if let Err(err) = stabilize_download(&temp_path) {
        log_debug(&format!("Download stabilization failed: {err}"));
        return Err(err);
    }

    if let Err(err) =
        verify_download_integrity(&temp_path, Some(expected_size), expected_hash.as_deref())
    {
        log_debug(&format!("Download integrity failed: {err}"));
        let _ = std::fs::remove_file(&temp_path);
        return Err(err);
    }
    log_debug("Download integrity ok.");

    write_update_metadata(&temp_path, expected_size, expected_hash.as_deref());

    if !prompt_restart_after_download(hwnd, language) {
        return Ok(UpdateAction::Deferred);
    }

    launch_self_updater(&current_exe, &temp_path).map_err(|err| err.to_string())?;

    unsafe {
        let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
    }
    Ok(UpdateAction::Started)
}

fn temp_update_path(current_exe: &Path) -> Result<PathBuf, String> {
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "Invalid executable name".to_string())?;
    let mut path = std::env::temp_dir();
    path.push(format!("{file_name}.update"));
    Ok(path)
}

fn pending_update_path() -> Option<PathBuf> {
    let current_exe = std::env::current_exe().ok()?;
    temp_update_path(&current_exe).ok()
}

fn probe_dir_writable(current_exe: &Path) -> Result<(), io::Error> {
    let dir = current_exe
        .parent()
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Missing executable directory"))?;
    let probe_name = format!("novapad_write_probe_{}.tmp", std::process::id());
    let probe_path = dir.join(probe_name);
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&probe_path)?;
    let _ = file.write_all(b"ok");
    let _ = std::fs::remove_file(&probe_path);
    Ok(())
}

fn can_open_exe_for_update(current_exe: &Path) -> Result<(), io::Error> {
    let path_w = to_wide(&current_exe.to_string_lossy());
    unsafe {
        let handle = match CreateFileW(
            PCWSTR(path_w.as_ptr()),
            FILE_GENERIC_READ.0,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            None,
        ) {
            Ok(handle) => handle,
            Err(_) => return Err(io::Error::last_os_error()),
        };
        let _ = CloseHandle(handle);
    }
    Ok(())
}

fn available_disk_bytes(current_exe: &Path) -> Result<u64, String> {
    let dir = current_exe
        .parent()
        .ok_or_else(|| "Missing executable directory".to_string())?;
    let path_w = to_wide(&dir.to_string_lossy());
    let mut free_bytes: u64 = 0;
    unsafe {
        if GetDiskFreeSpaceExW(PCWSTR(path_w.as_ptr()), Some(&mut free_bytes), None, None).is_err()
        {
            return Err(io::Error::last_os_error().to_string());
        }
    }
    Ok(free_bytes)
}

fn format_mb(bytes: u64) -> String {
    const MB: u64 = 1024 * 1024;
    let mb = bytes.div_ceil(MB);
    mb.to_string()
}

fn is_sharing_violation(err: &io::Error) -> bool {
    matches!(err.raw_os_error(), Some(32 | 33))
}

fn is_access_denied(err: &io::Error) -> bool {
    matches!(err.raw_os_error(), Some(5))
}

fn is_transient_win32_code(code: i32) -> bool {
    matches!(code, 5 | 32 | 33)
}

fn classify_win32_code(code: i32) -> &'static str {
    match code {
        5 => "PERMISSION",
        32 | 33 => "LOCK",
        _ => "OTHER",
    }
}

fn classify_io_error(err: &io::Error) -> &'static str {
    if let Some(code) = err.raw_os_error() {
        classify_win32_code(code)
    } else {
        "OTHER"
    }
}

fn retry_delay(attempt: u32) -> std::time::Duration {
    let base_ms: u64 = 200;
    let max_ms: u64 = 5_000;
    let shift = attempt.min(4);
    let delay = (base_ms << shift).min(max_ms);
    std::time::Duration::from_millis(delay)
}

fn path_volume_label(path: &Path) -> String {
    let prefix = path.components().next();
    match prefix {
        Some(std::path::Component::Prefix(prefix)) => format!("{:?}", prefix.kind()),
        _ => "unknown".to_string(),
    }
}

fn is_protected_install_dir(dir: &Path) -> bool {
    let dir_str = dir.to_string_lossy().to_lowercase();
    let mut protected = Vec::new();
    if let Ok(val) = std::env::var("programfiles") {
        protected.push(val.to_lowercase());
    }
    if let Ok(val) = std::env::var("programfiles(x86)") {
        protected.push(val.to_lowercase());
    }
    if let Ok(val) = std::env::var("windir") {
        protected.push(val.to_lowercase());
    }
    protected.iter().any(|root| dir_str.starts_with(root))
}

fn log_win32_error(context: &str, err: &io::Error) {
    if let Some(code) = err.raw_os_error() {
        let class = classify_win32_code(code);
        log_debug(&format!("{context}: win32={code} class={class} msg={err}"));
    } else {
        log_debug(&format!("{context}: class=OTHER {err}"));
    }
}

fn flush_file_on_disk(path: &Path) -> Result<(), io::Error> {
    let path_w = to_wide(&path.to_string_lossy());
    unsafe {
        let handle = CreateFileW(
            PCWSTR(path_w.as_ptr()),
            FILE_GENERIC_WRITE.0,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            None,
        )?;
        let _ = FlushFileBuffers(handle);
        let _ = CloseHandle(handle);
    }
    Ok(())
}

fn copy_to_target_dir(new_exe: &Path, current_exe: &Path) -> Result<PathBuf, io::Error> {
    let dir = current_exe
        .parent()
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Missing executable directory"))?;
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Invalid executable name"))?;
    let temp_name = format!("{file_name}.tmp.{}", std::process::id());
    let temp_path = dir.join(temp_name);
    std::fs::copy(new_exe, &temp_path)?;
    let _ = flush_file_on_disk(&temp_path);
    Ok(temp_path)
}

fn temp_update_meta_path(update_path: &Path) -> PathBuf {
    PathBuf::from(format!("{}.meta", update_path.to_string_lossy()))
}

fn write_update_metadata(update_path: &Path, expected_size: u64, sha256: Option<&str>) {
    let meta_path = temp_update_meta_path(update_path);
    let mut lines = Vec::new();
    lines.push(format!("size={expected_size}"));
    if let Some(hash) = sha256 {
        lines.push(format!("sha256={hash}"));
    }
    let _ = std::fs::write(meta_path, lines.join("\r\n"));
}

fn read_update_metadata(update_path: &Path) -> (Option<u64>, Option<String>) {
    let meta_path = temp_update_meta_path(update_path);
    let content = std::fs::read_to_string(meta_path).ok();
    let Some(content) = content else {
        return (None, None);
    };
    let mut size = None;
    let mut hash = None;
    for line in content.lines() {
        if let Some(rest) = line.strip_prefix("size=") {
            size = rest.trim().parse().ok();
        } else if let Some(rest) = line.strip_prefix("sha256=") {
            let value = rest.trim();
            if !value.is_empty() {
                hash = Some(value.to_string());
            }
        }
    }
    (size, hash)
}

fn download_file(url: &str, target: &Path) -> Result<(), String> {
    let client = reqwest::blocking::Client::builder()
        .user_agent(USER_AGENT)
        .build()
        .map_err(|err| err.to_string())?;
    let mut resp = client.get(url).send().map_err(|err| err.to_string())?;
    resp = resp.error_for_status().map_err(|err| err.to_string())?;
    let expected_len = resp.content_length();

    let mut file = File::create(target).map_err(|err| err.to_string())?;
    let written = io::copy(&mut resp, &mut file).map_err(|err| err.to_string())?;
    file.flush().map_err(|err| err.to_string())?;
    if let Some(expected) = expected_len {
        if written != expected {
            let _ = std::fs::remove_file(target);
            return Err("Download incomplete".to_string());
        }
    }
    Ok(())
}

fn compute_sha256(path: &Path) -> Result<String, String> {
    use sha2::{Digest, Sha256};
    let mut file = File::open(path).map_err(|err| err.to_string())?;
    let mut hasher = Sha256::new();
    let mut buf = [0u8; 8192];
    loop {
        let read = file.read(&mut buf).map_err(|err| err.to_string())?;
        if read == 0 {
            break;
        }
        hasher.update(&buf[..read]);
    }
    let digest = hasher.finalize();
    Ok(format!("{:x}", digest))
}

fn parse_sha256_file(content: &str, target_name: &str) -> Option<String> {
    for line in content.lines() {
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        let parts: Vec<&str> = line.split_whitespace().collect();
        if parts.len() >= 2 {
            let hash = parts[0];
            let name = parts[1].trim_start_matches('*');
            if name.eq_ignore_ascii_case(target_name) {
                return Some(hash.to_string());
            }
        } else if parts.len() == 1 && line.len() >= 64 {
            return Some(parts[0].to_string());
        }
    }
    None
}

fn download_sha256_optional(url: &str, target_name: &str) -> Option<String> {
    let mut temp = std::env::temp_dir();
    temp.push(format!("novapad_sha256_{}.tmp", std::process::id()));
    match download_file(url, &temp) {
        Ok(()) => {
            let content = std::fs::read_to_string(&temp).ok();
            let _ = std::fs::remove_file(&temp);
            content.and_then(|text| parse_sha256_file(&text, target_name))
        }
        Err(err) => {
            log_debug(&format!("Sha256 download failed: {err}"));
            let _ = std::fs::remove_file(&temp);
            None
        }
    }
}

fn stabilize_download(path: &Path) -> Result<(), String> {
    let start = std::time::Instant::now();
    let mut last_len: Option<u64> = None;
    let mut last_write: Option<std::time::SystemTime> = None;
    let mut stable_hits = 0u32;
    let mut attempt = 0u32;
    loop {
        attempt += 1;
        match std::fs::File::open(path) {
            Ok(mut file) => {
                let mut buf = [0u8; 512];
                let _ = file.read(&mut buf);
                if let Ok(meta) = file.metadata() {
                    let len = meta.len();
                    let write_time = meta.modified().ok();
                    if last_len == Some(len) && last_write == write_time {
                        stable_hits += 1;
                    } else {
                        stable_hits = 0;
                    }
                    last_len = Some(len);
                    last_write = write_time;
                    if stable_hits >= 1 {
                        log_debug(&format!(
                            "Download stabilization ok. attempts={attempt} size={len}"
                        ));
                        return Ok(());
                    }
                }
            }
            Err(err) => {
                if let Some(code) = err.raw_os_error() {
                    if is_transient_win32_code(code) {
                        log_win32_error("Download stabilization transient", &err);
                    } else {
                        log_win32_error("Download stabilization failed", &err);
                        return Err(err.to_string());
                    }
                } else {
                    return Err(err.to_string());
                }
            }
        }
        if start.elapsed() > std::time::Duration::from_secs(5) {
            return Err("Download stabilization timed out".to_string());
        }
        std::thread::sleep(retry_delay(attempt));
    }
}

fn verify_download_integrity(
    path: &Path,
    expected_size: Option<u64>,
    expected_sha256: Option<&str>,
) -> Result<(), String> {
    let meta = std::fs::metadata(path).map_err(|err| err.to_string())?;
    let size = meta.len();
    if let Some(expected) = expected_size {
        if size != expected {
            return Err(format!(
                "Integrity failed: size mismatch expected={expected} actual={size}"
            ));
        }
    }
    if let Some(expected_hash) = expected_sha256 {
        let actual = compute_sha256(path)?;
        if !actual.eq_ignore_ascii_case(expected_hash) {
            return Err(format!(
                "Integrity failed: sha256 mismatch expected={expected_hash} actual={actual}"
            ));
        }
    }
    Ok(())
}

fn is_same_executable(a: &Path, b: &Path) -> bool {
    let Ok(a_meta) = a.metadata() else {
        return false;
    };
    let Ok(b_meta) = b.metadata() else {
        return false;
    };
    if a_meta.len() != b_meta.len() {
        return false;
    }
    let Ok(a_hash) = compute_sha256(a) else {
        return false;
    };
    let Ok(b_hash) = compute_sha256(b) else {
        return false;
    };
    a_hash.eq_ignore_ascii_case(&b_hash)
}

fn paths_equal(a: &Path, b: &Path) -> bool {
    let ca = std::fs::canonicalize(a).unwrap_or_else(|_| a.to_path_buf());
    let cb = std::fs::canonicalize(b).unwrap_or_else(|_| b.to_path_buf());
    ca == cb
}

fn relaunch_elevated(self_exe: &Path, args: &[String]) -> io::Result<i32> {
    let exe_w = to_wide(&self_exe.to_string_lossy());
    let mut params = Vec::new();
    for arg in args.iter().skip(1) {
        params.push(quote_arg(arg));
    }
    if !params.iter().any(|arg| arg == "--elevated") {
        params.push("--elevated".to_string());
    }
    let params_str = params.join(" ");
    let params_w = to_wide(&params_str);
    let verb = to_wide("runas");
    unsafe {
        let result = ShellExecuteW(
            HWND(0),
            PCWSTR(verb.as_ptr()),
            PCWSTR(exe_w.as_ptr()),
            PCWSTR(params_w.as_ptr()),
            PCWSTR::null(),
            SW_SHOW,
        );
        if result.0 as isize <= 32 {
            let err = io::Error::last_os_error();
            log_win32_error("Elevation failed", &err);
            return Ok(EXIT_UAC_CANCELLED);
        }
    }
    Ok(EXIT_OK)
}

fn quote_arg(arg: &str) -> String {
    if arg.contains(' ') || arg.contains('"') {
        format!("\"{}\"", arg.replace('"', "\\\""))
    } else {
        arg.to_string()
    }
}

fn launch_self_updater(current_exe: &Path, new_exe: &Path) -> io::Result<()> {
    let pid = std::process::id();
    let mut last_err: Option<io::Error> = None;
    for updater_exe in runner_updater_candidates(current_exe) {
        if !ensure_dir_writable_for_runner(&updater_exe) {
            continue;
        }
        match std::fs::copy(current_exe, &updater_exe) {
            Ok(_) => {
                log_debug(&format!(
                    "Update launch: runner={} current={} new={}",
                    updater_exe.display(),
                    current_exe.display(),
                    new_exe.display()
                ));
                std::process::Command::new(updater_exe)
                    .arg("--self-update")
                    .arg("--pid")
                    .arg(pid.to_string())
                    .arg("--current")
                    .arg(current_exe)
                    .arg("--new")
                    .arg(new_exe)
                    .arg("--restart")
                    .spawn()?;
                return Ok(());
            }
            Err(err) => {
                log_win32_error("Update launch: runner copy failed", &err);
                last_err = Some(err);
            }
        }
    }
    Err(last_err.unwrap_or_else(|| {
        io::Error::new(io::ErrorKind::PermissionDenied, "Runner path not writable")
    }))
}

fn runner_updater_candidates(current_exe: &Path) -> Vec<PathBuf> {
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("novapad.exe");
    let mut candidates = Vec::new();
    if let Some(base) = std::env::var_os("LOCALAPPDATA") {
        let mut path = PathBuf::from(base);
        path.push(UPDATER_DIR_NAME);
        path.push(format!("{file_name}.Updater.exe"));
        candidates.push(path);
    }
    let mut temp_path = std::env::temp_dir();
    temp_path.push(format!("{file_name}.Updater.exe"));
    candidates.push(temp_path);
    candidates
}

fn ensure_dir_writable_for_runner(path: &Path) -> bool {
    let Some(parent) = path.parent() else {
        return false;
    };
    if std::fs::create_dir_all(parent).is_err() {
        return false;
    }
    let probe_name = format!("novapad_updater_probe_{}.tmp", std::process::id());
    let probe_path = parent.join(probe_name);
    match OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&probe_path)
    {
        Ok(_) => {
            let _ = std::fs::remove_file(&probe_path);
            true
        }
        Err(_) => false,
    }
}

pub(crate) fn run_self_update(args: &[String]) -> Result<i32, String> {
    let mut pid: Option<u32> = None;
    let mut current: Option<PathBuf> = None;
    let mut new: Option<PathBuf> = None;
    let mut restart = false;
    let mut elevated = false;
    let mut schedule_retry = false;

    let mut it = args.iter().peekable();
    while let Some(arg) = it.next() {
        match arg.as_str() {
            "--pid" => {
                let Some(value) = it.next() else {
                    return Err("Missing --pid value".to_string());
                };
                pid = value.parse().ok();
            }
            "--current" => {
                let Some(value) = it.next() else {
                    return Err("Missing --current value".to_string());
                };
                current = Some(PathBuf::from(value));
            }
            "--new" => {
                let Some(value) = it.next() else {
                    return Err("Missing --new value".to_string());
                };
                new = Some(PathBuf::from(value));
            }
            "--restart" => restart = true,
            "--elevated" => elevated = true,
            "--schedule-retry" => schedule_retry = true,
            _ => {}
        }
    }

    let pid = pid.ok_or_else(|| "Missing pid".to_string())?;
    let current = current.ok_or_else(|| "Missing current path".to_string())?;
    let new = new.ok_or_else(|| "Missing new path".to_string())?;

    let self_exe = std::env::current_exe().map_err(|err| err.to_string())?;
    if paths_equal(&self_exe, &current) {
        for runner in runner_updater_candidates(&current) {
            if !ensure_dir_writable_for_runner(&runner) {
                continue;
            }
            if std::fs::copy(&self_exe, &runner).is_err() {
                continue;
            }
            log_debug(&format!(
                "Self-update launched from target. Relaunching runner: {}",
                runner.display()
            ));
            let mut cmd = std::process::Command::new(runner);
            for arg in args {
                cmd.arg(arg);
            }
            let _ = cmd.spawn();
            return Ok(EXIT_OK);
        }
        log_debug("Self-update launched from target. Runner copy failed.");
        return Ok(EXIT_OK);
    }

    log_debug(&format!(
        "Self-update: pid={} current={} new={} self={}",
        pid,
        current.display(),
        new.display(),
        self_exe.display()
    ));
    log_debug(&format!(
        "Self-update volumes: current={} new={}",
        path_volume_label(&current),
        path_volume_label(&new)
    ));
    let cross_volume = path_volume_label(&current) != path_volume_label(&new);
    log_debug(&format!("Self-update cross-volume: {cross_volume}"));
    let (meta_size, meta_hash) = read_update_metadata(&new);
    if let Err(err) = stabilize_download(&new) {
        log_debug(&format!("Self-update: staged file unstable: {err}"));
        return Ok(EXIT_DOWNLOAD_FAILED);
    }
    if let Err(err) = verify_download_integrity(&new, meta_size, meta_hash.as_deref()) {
        log_debug(&format!("Self-update: integrity failed: {err}"));
        let _ = std::fs::remove_file(&new);
        let _ = std::fs::remove_file(temp_update_meta_path(&new));
        return Ok(EXIT_INTEGRITY_FAILED);
    }

    wait_for_process_exit(pid);
    let _ = recover_backup_if_needed(&current);

    let dir = current
        .parent()
        .ok_or_else(|| "Missing executable directory".to_string())?;
    let probe_err = probe_dir_writable(&current).err();
    if let Some(err) = &probe_err {
        log_debug(&format!(
            "Self-update: dir probe failed: {err} class={}",
            classify_io_error(err)
        ));
    }
    let needs_elevation =
        is_protected_install_dir(dir) || probe_err.as_ref().is_some_and(is_access_denied);
    if needs_elevation && !elevated {
        log_debug("Self-update: requesting elevation.");
        let relaunch = relaunch_elevated(&self_exe, args).map_err(|err| err.to_string())?;
        return Ok(relaunch);
    }

    let candidate = match copy_to_target_dir(&new, &current) {
        Ok(path) => path,
        Err(err) => {
            log_win32_error("Self-update: copy to target dir failed", &err);
            return Ok(EXIT_REPLACE_FAILED);
        }
    };
    log_debug(&format!(
        "Self-update: staged candidate={} target={} backup={}",
        candidate.display(),
        current.display(),
        backup_executable_path(&current)
            .map(|p| p.display().to_string())
            .unwrap_or_else(|_| "(unknown)".to_string())
    ));

    match replace_executable(&current, &candidate) {
        ReplaceResult::Replaced => {}
        ReplaceResult::ScheduledReboot => {
            log_debug("Self-update: scheduled replacement on reboot.");
            if schedule_retry && elevated {
                log_debug("Schedule retry (elevated) succeeded.");
            }
            let _ = remove_update_lock(&current);
            let language = load_settings().language;
            show_update_info(language, UpdateInfo::RebootRequired);
            return Ok(EXIT_REBOOT_REQUIRED);
        }
        ReplaceResult::ScheduleFailed { err } => {
            if is_access_denied(&err) && !elevated {
                log_win32_error(
                    "Self-update: reboot schedule denied, retrying with elevation",
                    &err,
                );
                let mut retry_args = args.to_vec();
                if !retry_args.iter().any(|arg| arg == "--schedule-retry") {
                    retry_args.push("--schedule-retry".to_string());
                }
                let relaunch =
                    relaunch_elevated(&self_exe, &retry_args).map_err(|err| err.to_string())?;
                if relaunch == EXIT_UAC_CANCELLED {
                    log_debug("Self-update: elevation cancelled during reboot schedule retry.");
                }
                return Ok(relaunch);
            }
            if schedule_retry && elevated {
                log_debug("Schedule retry (elevated) failed.");
            }
            log_win32_error("Self-update: reboot schedule failed", &err);
            let language = load_settings().language;
            if err.kind() == io::ErrorKind::PermissionDenied || is_access_denied(&err) {
                show_permission_error(language);
            } else {
                show_update_error(language, UpdateError::Replace);
            }
            return Ok(EXIT_REPLACE_FAILED);
        }
        ReplaceResult::Failed { restored, err } => {
            let language = load_settings().language;
            if err.kind() == io::ErrorKind::PermissionDenied {
                show_permission_error(language);
            } else if is_sharing_violation(&err) {
                if restored {
                    show_update_error(language, UpdateError::ReplaceLockedRestored);
                } else {
                    show_update_error(language, UpdateError::ReplaceLocked);
                }
            } else if restored {
                show_update_error(language, UpdateError::ReplaceRestored);
            } else {
                show_update_error(language, UpdateError::Replace);
            }
            return Ok(EXIT_REPLACE_FAILED);
        }
    }
    let _ = std::fs::remove_file(&new);
    let _ = std::fs::remove_file(temp_update_meta_path(&new));

    let language = load_settings().language;
    if restart {
        match std::process::Command::new(&current)
            .current_dir(dir)
            .spawn()
        {
            Ok(_) => show_update_info(language, UpdateInfo::Completed),
            Err(err) => {
                let _ = restore_backup(&current);
                let _ = remove_update_lock(&current);
                log_debug(&format!("Self-update: restart failed: {err}"));
                show_update_error(language, UpdateError::RestartFailed);
                return Ok(EXIT_REPLACE_FAILED);
            }
        }
    } else {
        show_update_info(language, UpdateInfo::Completed);
    }
    let _ = remove_update_lock(&current);
    Ok(EXIT_OK)
}

fn wait_for_process_exit(pid: u32) {
    unsafe {
        if let Ok(handle) = OpenProcess(PROCESS_SYNCHRONIZE, false, pid) {
            let result = WaitForSingleObject(handle, 20_000);
            log_debug(&format!(
                "Self-update: waited for pid={pid} result={result:?}"
            ));
            let _ = CloseHandle(handle);
        } else {
            log_debug(&format!("Self-update: failed to open pid={pid}"));
        }
    }
}

enum ReplaceResult {
    Replaced,
    ScheduledReboot,
    ScheduleFailed { err: io::Error },
    Failed { restored: bool, err: io::Error },
}

fn replace_executable(current: &Path, new: &Path) -> ReplaceResult {
    let backup = match backup_executable_path(current) {
        Ok(path) => path,
        Err(err) => {
            return ReplaceResult::Failed {
                restored: current.exists(),
                err,
            };
        }
    };
    let mut last_err: Option<io::Error> = None;
    let start = std::time::Instant::now();
    let mut attempt: u32 = 0;
    loop {
        attempt += 1;
        let _ = recover_backup_if_needed(current);
        match replace_with_win32(current, new, &backup) {
            Ok(()) => return ReplaceResult::Replaced,
            Err(err) => {
                if last_err.is_none() {
                    log_debug("Replace: first failure recorded.");
                }
                log_win32_error(&format!("Replace attempt {attempt} failed"), &err);
                last_err = Some(err);
            }
        }
        let elapsed = start.elapsed();
        if elapsed > std::time::Duration::from_secs(20) {
            break;
        }
        if let Some(code) = last_err.as_ref().and_then(|e| e.raw_os_error()) {
            if is_transient_win32_code(code) {
                std::thread::sleep(retry_delay(attempt));
                continue;
            }
        }
        break;
    }
    if let Some(err) = last_err.as_ref() {
        if err.raw_os_error().is_some_and(is_transient_win32_code) {
            return match schedule_replace_on_reboot(new, current) {
                Ok(()) => ReplaceResult::ScheduledReboot,
                Err(err) => ReplaceResult::ScheduleFailed { err },
            };
        }
    }
    ReplaceResult::Failed {
        restored: current.exists(),
        err: last_err.unwrap_or_else(|| {
            io::Error::new(io::ErrorKind::Other, "Failed to replace executable")
        }),
    }
}

fn replace_with_win32(target: &Path, source: &Path, backup: &Path) -> io::Result<()> {
    if target.parent() != source.parent() {
        return Err(io::Error::new(
            io::ErrorKind::Other,
            "Source must be in target directory",
        ));
    }
    let target_w = to_wide(&target.to_string_lossy());
    let source_w = to_wide(&source.to_string_lossy());
    let backup_w = to_wide(&backup.to_string_lossy());
    unsafe {
        if ReplaceFileW(
            PCWSTR(target_w.as_ptr()),
            PCWSTR(source_w.as_ptr()),
            PCWSTR(backup_w.as_ptr()),
            REPLACEFILE_WRITE_THROUGH,
            None,
            None,
        )
        .is_ok()
        {
            return Ok(());
        }
    }
    let err = io::Error::last_os_error();
    log_win32_error("ReplaceFileW failed", &err);

    unsafe {
        if MoveFileExW(
            PCWSTR(target_w.as_ptr()),
            PCWSTR(backup_w.as_ptr()),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
        .is_err()
        {
            return Err(io::Error::last_os_error());
        }
        if MoveFileExW(
            PCWSTR(source_w.as_ptr()),
            PCWSTR(target_w.as_ptr()),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
        .is_err()
        {
            let move_err = io::Error::last_os_error();
            let _ = MoveFileExW(
                PCWSTR(backup_w.as_ptr()),
                PCWSTR(target_w.as_ptr()),
                MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
            );
            return Err(move_err);
        }
    }
    Ok(())
}

fn schedule_replace_on_reboot(source: &Path, target: &Path) -> io::Result<()> {
    if source.parent() != target.parent() {
        log_debug(&format!(
            "Reboot schedule denied: source not in target dir. source={} target={}",
            source.display(),
            target.display()
        ));
        return Err(io::Error::new(
            io::ErrorKind::Other,
            "Source must be in target directory",
        ));
    }
    let source_w = to_wide(&source.to_string_lossy());
    let target_w = to_wide(&target.to_string_lossy());
    unsafe {
        if MoveFileExW(
            PCWSTR(source_w.as_ptr()),
            PCWSTR(target_w.as_ptr()),
            MOVEFILE_DELAY_UNTIL_REBOOT | MOVEFILE_REPLACE_EXISTING,
        )
        .is_ok()
        {
            log_debug(&format!(
                "Scheduled replace on reboot: source={} target={}",
                source.display(),
                target.display()
            ));
            return Ok(());
        }
    }
    let err = io::Error::last_os_error();
    log_win32_error("Schedule replace on reboot failed", &err);
    Err(err)
}

fn recover_backup_if_needed(current: &Path) -> io::Result<()> {
    let backup = backup_executable_path(current)?;
    if !current.exists() && backup.exists() {
        log_debug(&format!(
            "Recovery: restoring backup {} -> {}",
            backup.display(),
            current.display()
        ));
        std::fs::rename(&backup, current)?;
    }
    Ok(())
}

fn backup_executable_path(current: &Path) -> Result<PathBuf, io::Error> {
    let file_name = current
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Invalid executable name"))?;
    Ok(current.with_file_name(format!("{file_name}.old")))
}

enum UpdateLockError {
    InProgress,
    Other(String),
}

struct UpdateLock {
    path: PathBuf,
    keep: bool,
}

impl UpdateLock {
    fn keep(&mut self) {
        self.keep = true;
    }
}

impl Drop for UpdateLock {
    fn drop(&mut self) {
        if !self.keep {
            let _ = std::fs::remove_file(&self.path);
        }
    }
}

fn acquire_update_lock(current_exe: &Path) -> Result<UpdateLock, UpdateLockError> {
    let paths =
        update_lock_paths(current_exe).map_err(|err| UpdateLockError::Other(err.to_string()))?;
    let mut last_err: Option<io::Error> = None;
    for path in paths {
        if path.exists() {
            if let Some(pid) = read_lock_pid(&path) {
                if is_process_running(pid) {
                    return Err(UpdateLockError::InProgress);
                }
            }
            let _ = std::fs::remove_file(&path);
        }
        let mut file = match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => file,
            Err(err) if err.kind() == io::ErrorKind::AlreadyExists => {
                return Err(UpdateLockError::InProgress);
            }
            Err(err) => {
                if is_access_denied(&err) {
                    last_err = Some(err);
                    continue;
                }
                return Err(UpdateLockError::Other(err.to_string()));
            }
        };
        let _ = writeln!(file, "{}", std::process::id());
        return Ok(UpdateLock { path, keep: false });
    }
    Err(UpdateLockError::Other(
        last_err
            .unwrap_or_else(|| io::Error::new(io::ErrorKind::Other, "Lock unavailable"))
            .to_string(),
    ))
}

fn update_lock_path_primary(current_exe: &Path) -> Result<PathBuf, io::Error> {
    let dir = current_exe
        .parent()
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Missing executable directory"))?;
    Ok(dir.join(UPDATE_LOCK_NAME))
}

fn update_lock_path_temp(current_exe: &Path) -> PathBuf {
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("novapad.exe");
    let mut path = std::env::temp_dir();
    path.push(format!("{UPDATE_LOCK_NAME}.{file_name}"));
    path
}

fn update_lock_paths(current_exe: &Path) -> Result<Vec<PathBuf>, io::Error> {
    let primary = update_lock_path_primary(current_exe)?;
    let temp = update_lock_path_temp(current_exe);
    Ok(vec![primary, temp])
}

fn read_lock_pid(path: &Path) -> Option<u32> {
    let content = std::fs::read_to_string(path).ok()?;
    content.trim().parse::<u32>().ok()
}

fn is_process_running(pid: u32) -> bool {
    unsafe {
        if let Ok(handle) = OpenProcess(PROCESS_SYNCHRONIZE, false, pid) {
            let _ = CloseHandle(handle);
            return true;
        }
    }
    false
}

fn remove_update_lock(current_exe: &Path) -> io::Result<()> {
    let paths = update_lock_paths(current_exe)?;
    for path in paths {
        if path.exists() {
            let _ = std::fs::remove_file(path);
        }
    }
    Ok(())
}

fn restore_backup(current: &Path) -> io::Result<()> {
    let backup = backup_executable_path(current)?;
    if !backup.exists() {
        return Ok(());
    }
    let failed = current.with_extension("failed");
    let _ = std::fs::rename(current, &failed);
    std::fs::rename(&backup, current)?;
    let _ = std::fs::remove_file(failed);
    Ok(())
}

fn show_permission_error(language: Language) {
    let text = i18n::tr(language, "updater.permission_error");
    let title = i18n::tr(language, "updater.title");
    let owner = find_main_window();
    let _ = show_update_message(
        owner,
        &text,
        &title,
        MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
    );
}

enum UpdateError {
    Network,
    NoPortableAsset,
    Download,
    Replace,
    ReplaceRestored,
    ReplaceLocked,
    ReplaceLockedRestored,
    Concurrent,
    RestartFailed,
}

#[derive(PartialEq, Eq)]
enum UpdateInfo {
    NoUpdate,
    NoPending,
    Completed,
    RebootRequired,
}

fn show_update_error(language: Language, error: UpdateError) {
    let text_key = match error {
        UpdateError::Network => "updater.error.network",
        UpdateError::NoPortableAsset => "updater.error.no_portable",
        UpdateError::Download => "updater.error.download",
        UpdateError::Replace => "updater.error.replace",
        UpdateError::ReplaceRestored => "updater.error.replace_restored",
        UpdateError::ReplaceLocked => "updater.error.replace_locked",
        UpdateError::ReplaceLockedRestored => "updater.error.replace_locked_restored",
        UpdateError::Concurrent => "updater.error.concurrent",
        UpdateError::RestartFailed => "updater.error.restart_failed",
    };
    let text = i18n::tr(language, text_key);
    let title = i18n::tr(language, "updater.title");
    let owner = find_main_window();
    let _ = show_update_message(
        owner,
        &text,
        &title,
        MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
    );
}

fn show_update_error_with_url(language: Language, key: &str, url: &str) {
    let text = i18n::tr_f(language, key, &[("url", url)]);
    let title = i18n::tr(language, "updater.title");
    let owner = find_main_window();
    let _ = show_update_message(
        owner,
        &text,
        &title,
        MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
    );
}

fn show_update_error_args(language: Language, key: &str, args: &[(&str, &str)]) {
    let text = i18n::tr_f(language, key, args);
    let title = i18n::tr(language, "updater.title");
    let owner = find_main_window();
    let _ = show_update_message(
        owner,
        &text,
        &title,
        MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
    );
}

pub(crate) fn cleanup_backup_on_start() {
    let Ok(current_exe) = std::env::current_exe() else {
        return;
    };
    let _ = recover_backup_if_needed(&current_exe);
    let Ok(backup) = backup_executable_path(&current_exe) else {
        return;
    };
    if current_exe.exists() && backup.exists() {
        let _ = std::fs::remove_file(backup);
    }
}

pub(crate) fn cleanup_update_lock_on_start() {
    let Ok(current_exe) = std::env::current_exe() else {
        return;
    };
    let Ok(lock_paths) = update_lock_paths(&current_exe) else {
        return;
    };
    for lock_path in lock_paths {
        if !lock_path.exists() {
            continue;
        }
        let pid = read_lock_pid(&lock_path);
        if pid.is_none_or(|pid| !is_process_running(pid)) {
            let _ = std::fs::remove_file(lock_path);
        }
    }
}

pub(crate) fn check_pending_update(hwnd: HWND, force: bool) {
    static PROMPTED: AtomicBool = AtomicBool::new(false);
    if !force && PROMPTED.swap(true, Ordering::SeqCst) {
        return;
    }
    let Some(pending) = pending_update_path() else {
        if force {
            let language = app_language(hwnd);
            show_update_info(language, UpdateInfo::NoPending);
        }
        return;
    };
    if !pending.exists() {
        if force {
            let language = app_language(hwnd);
            show_update_info(language, UpdateInfo::NoPending);
        }
        return;
    }
    let Ok(meta) = pending.metadata() else {
        if force {
            let language = app_language(hwnd);
            show_update_info(language, UpdateInfo::NoPending);
        }
        return;
    };
    if meta.len() == 0 {
        let _ = std::fs::remove_file(&pending);
        let _ = std::fs::remove_file(temp_update_meta_path(&pending));
        if force {
            let language = app_language(hwnd);
            show_update_info(language, UpdateInfo::NoPending);
        }
        return;
    }

    let language = app_language(hwnd);
    let (meta_size, meta_hash) = read_update_metadata(&pending);
    if let Err(err) = stabilize_download(&pending) {
        log_debug(&format!("Pending update: staged file unstable: {err}"));
        let _ = std::fs::remove_file(&pending);
        let _ = std::fs::remove_file(temp_update_meta_path(&pending));
        show_update_error(language, UpdateError::Download);
        return;
    }
    if let Err(err) = verify_download_integrity(&pending, meta_size, meta_hash.as_deref()) {
        log_debug(&format!("Pending update: integrity failed: {err}"));
        let _ = std::fs::remove_file(&pending);
        let _ = std::fs::remove_file(temp_update_meta_path(&pending));
        show_update_error(language, UpdateError::Download);
        return;
    }
    let current_exe = match std::env::current_exe() {
        Ok(path) => path,
        Err(err) => {
            log_debug(&format!("Pending update: current exe not available: {err}"));
            return;
        }
    };
    if is_same_executable(&pending, &current_exe) {
        log_debug("Pending update matches current exe. Clearing pending update.");
        let _ = std::fs::remove_file(&pending);
        let _ = std::fs::remove_file(temp_update_meta_path(&pending));
        if force {
            show_update_info(language, UpdateInfo::NoPending);
        }
        return;
    }
    let mut update_lock = match acquire_update_lock(&current_exe) {
        Ok(lock) => lock,
        Err(UpdateLockError::InProgress) => {
            show_update_error(language, UpdateError::Concurrent);
            return;
        }
        Err(UpdateLockError::Other(err)) => {
            log_debug(&format!("Pending update: cannot create update lock: {err}"));
            show_update_error_with_url(language, "updater.error.access", DIRECT_DOWNLOAD_URL);
            return;
        }
    };

    if !prompt_pending_update(hwnd, language) {
        return;
    }

    if let Err(err) = launch_self_updater(&current_exe, &pending) {
        log_debug(&format!("Pending update: failed to launch updater: {err}"));
        show_update_error(language, UpdateError::Download);
        return;
    }
    update_lock.keep();
    unsafe {
        let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
    }
}

pub(crate) fn cleanup_update_temp_on_start() {
    let Ok(current_exe) = std::env::current_exe() else {
        return;
    };
    let Some(file_name) = current_exe.file_name().and_then(|name| name.to_str()) else {
        return;
    };
    let temp_dir = std::env::temp_dir();
    let Ok(entries) = std::fs::read_dir(&temp_dir) else {
        return;
    };
    let update_name = format!("{file_name}.update");
    let update_meta_name = format!("{update_name}.meta");
    let updater_prefix = format!("{file_name}.updater.");
    for entry in entries.flatten() {
        let path = entry.path();
        let Some(name) = path.file_name().and_then(|name| name.to_str()) else {
            continue;
        };
        if (name.starts_with(&updater_prefix) && name.ends_with(".exe")) || name == update_meta_name
        {
            let _ = std::fs::remove_file(path);
        } else if name == update_name {
            if let Ok(meta) = path.metadata() {
                if meta.len() == 0 {
                    let _ = std::fs::remove_file(&path);
                    let _ = std::fs::remove_file(temp_update_meta_path(&path));
                }
            }
        }
    }
}

fn show_update_info(language: Language, info: UpdateInfo) {
    let text_key = match info {
        UpdateInfo::NoUpdate => "updater.info.no_update",
        UpdateInfo::NoPending => "updater.info.no_pending",
        UpdateInfo::Completed => "updater.info.completed",
        UpdateInfo::RebootRequired => "updater.info.reboot_required",
    };
    let text = i18n::tr(language, text_key);
    let title = i18n::tr(language, "updater.title");
    let owner = find_main_window();
    let _ = show_update_message(
        owner,
        &text,
        &title,
        MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND,
    );
}

fn focus_main_window() {
    let hwnd = find_main_window();
    if hwnd.0 == 0 {
        return;
    }
    unsafe {
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = SetForegroundWindow(hwnd);
        let _ = SetActiveWindow(hwnd);
        let _ = SetFocus(hwnd);
    }
}

fn find_main_window() -> HWND {
    let class_name = to_wide("NovapadWin32");
    unsafe { FindWindowW(PCWSTR(class_name.as_ptr()), PCWSTR::null()) }
}

fn show_update_message(
    owner: HWND,
    text: &str,
    title: &str,
    flags: MESSAGEBOX_STYLE,
) -> windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_RESULT {
    let text_w = to_wide(text);
    let title_w = to_wide(title);
    let result = unsafe {
        MessageBoxW(
            owner,
            PCWSTR(text_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            flags,
        )
    };
    focus_main_window();
    if owner.0 != 0 {
        unsafe {
            let _ = PostMessageW(owner, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
        }
    }
    result
}
