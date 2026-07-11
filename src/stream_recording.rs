use std::fs;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use windows::Win32::Foundation::{CloseHandle, HWND};
use windows::Win32::UI::WindowsAndMessaging::{MB_ICONWARNING, MB_OK};
use windows::core::PCWSTR;

use crate::accessibility::to_wide;
use windows::Win32::System::Threading::{
    GetExitCodeProcess, OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION,
};

use crate::app_windows::interpreter_select_window::InterpreterContextAction;
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::Language;

const CREATE_NO_WINDOW: u32 = 0x0800_0000;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum StreamRecordingKind {
    Radio,
    Tv,
}

impl StreamRecordingKind {
    fn folder_name(self) -> &'static str {
        match self {
            Self::Radio => "Registrazioni Radio",
            Self::Tv => "Registrazioni TV",
        }
    }

    fn extension(self) -> &'static str {
        match self {
            Self::Radio => "mp3",
            Self::Tv => "mp4",
        }
    }
}

pub(crate) fn start_radio_recording_and_playback(
    parent: HWND,
    url: &str,
    title: &str,
    language: Language,
) -> Result<PathBuf, String> {
    start_recording_and_playback(
        parent,
        url,
        title,
        None,
        false,
        StreamRecordingKind::Radio,
        Some(language),
    )
}

pub(crate) fn start_tv_recording_and_playback(
    parent: HWND,
    url: &str,
    title: &str,
    user_agent: &str,
    prefer_audio_description: bool,
) -> Result<PathBuf, String> {
    start_recording_and_playback(
        parent,
        url,
        title,
        Some(user_agent),
        prefer_audio_description,
        StreamRecordingKind::Tv,
        None,
    )
}

fn start_recording_and_playback(
    parent: HWND,
    url: &str,
    title: &str,
    user_agent: Option<&str>,
    prefer_audio_description: bool,
    kind: StreamRecordingKind,
    ui_language: Option<Language>,
) -> Result<PathBuf, String> {
    let url = url.trim();
    if url.is_empty() {
        return Err(recording_text(
            ui_language,
            "radio.recording_error_empty_url",
            "L'indirizzo dello stream è vuoto.",
        ));
    }

    let ffmpeg = find_ffmpeg_executable(ui_language)?;
    let output_path = next_recording_path(kind, title, ui_language)?;

    // Avviamo prima mpv. Alcuni redirect HLS dinamici, soprattutto quelli RAI,
    // non tollerano bene che FFmpeg apra il flusso per primo: il player resta
    // senza tracce e quindi senza audio. Dopo che mpv ha caricato le tracce,
    // apriamo una seconda connessione indipendente per la registrazione.
    let playback_result = match kind {
        StreamRecordingKind::Radio => {
            crate::launch_stream_url_in_mpv(parent, url, Some(title), None, None, None)
        }
        StreamRecordingKind::Tv => crate::launch_tv_stream_in_mpv(
            parent,
            url,
            title,
            user_agent.unwrap_or_default(),
            prefer_audio_description,
        ),
    };

    if let Err(err) = playback_result {
        remove_recording_file(&output_path, "player launch failure");
        return Err(err);
    }

    let Some(mpv_process_id) = crate::active_mpv_process_id(parent) else {
        remove_recording_file(&output_path, "missing player process");
        return Err(recording_text(
            ui_language,
            "radio.recording_error_link_player",
            "Il player è stato avviato, ma non è stato possibile collegare la registrazione al suo processo.",
        ));
    };

    if kind == StreamRecordingKind::Tv
        && !crate::wait_for_active_mpv_tracks(parent, Duration::from_secs(12))
    {
        crate::stop_managed_mpv_playback(parent);
        remove_recording_file(&output_path, "TV loading timeout");
        return Err(recording_text(
            ui_language,
            "radio.recording_error_link_player",
            "Il canale TV non ha completato il caricamento e la registrazione non è stata avviata.",
        ));
    }

    let mut recorder =
        match spawn_ffmpeg_recorder(&ffmpeg, url, &output_path, user_agent, kind, ui_language) {
            Ok(recorder) => recorder,
            Err(err) => {
                crate::stop_managed_mpv_playback(parent);
                remove_recording_file(&output_path, "FFmpeg launch failure");
                return Err(err);
            }
        };

    // Verifica rapida: se FFmpeg termina immediatamente, evitiamo di annunciare
    // una registrazione che in realtà non è partita.
    thread::sleep(Duration::from_millis(350));
    match recorder.try_wait() {
        Ok(Some(status)) => {
            crate::stop_managed_mpv_playback(parent);
            remove_recording_file(&output_path, "FFmpeg immediate exit");
            return Err(format!(
                "FFmpeg ha terminato subito la registrazione con stato {status}."
            ));
        }
        Ok(None) => {}
        Err(err) => {
            crate::stop_managed_mpv_playback(parent);
            remove_recording_file(&output_path, "FFmpeg startup check failure");
            return Err(format!(
                "Impossibile verificare l'avvio della registrazione: {err}"
            ));
        }
    }

    crate::log_debug(&format!(
        "Stream recording started: kind={kind:?} mpv_pid={mpv_process_id} output={}",
        output_path.display()
    ));

    let output_for_thread = output_path.clone();
    thread::spawn(move || {
        monitor_recording_until_player_closes(recorder, mpv_process_id, output_for_thread);
    });

    Ok(output_path)
}

fn spawn_ffmpeg_recorder(
    ffmpeg: &Path,
    url: &str,
    output_path: &Path,
    user_agent: Option<&str>,
    kind: StreamRecordingKind,
    ui_language: Option<Language>,
) -> Result<Child, String> {
    let mut command = Command::new(ffmpeg);
    command
        .arg("-y")
        .arg("-hide_banner")
        .arg("-loglevel")
        .arg("warning")
        .arg("-reconnect")
        .arg("1")
        .arg("-reconnect_streamed")
        .arg("1")
        .arg("-reconnect_delay_max")
        .arg("5");

    if let Some(user_agent) = user_agent.map(str::trim).filter(|value| !value.is_empty()) {
        command.arg("-user_agent").arg(user_agent);
    }

    command
        .arg("-fflags")
        .arg("+genpts+discardcorrupt")
        .arg("-i")
        .arg(url);
    match kind {
        StreamRecordingKind::Radio => {
            // Le radio possono trasmettere in AAC, Opus, Vorbis o MP3. Per ottenere
            // sempre un vero file MP3 riproducibile, l'audio viene ricodificato.
            command
                .arg("-map")
                .arg("0:a:0?")
                .arg("-vn")
                .arg("-sn")
                .arg("-dn")
                .arg("-c:a")
                .arg("libmp3lame")
                .arg("-q:a")
                .arg("2")
                .arg("-id3v2_version")
                .arg("3")
                .arg("-write_id3v1")
                .arg("1")
                .arg("-f")
                .arg("mp3");
        }
        StreamRecordingKind::Tv => {
            // Il video viene copiato senza perdita; l'audio viene convertito in AAC,
            // formato sempre compatibile con MP4. L'MP4 frammentato rimane leggibile
            // anche se la registrazione viene interrotta mentre è in corso.
            command
                .arg("-map")
                .arg("0:v:0?")
                .arg("-map")
                .arg("0:a:0?")
                .arg("-sn")
                .arg("-dn")
                .arg("-c:v")
                .arg("copy")
                .arg("-c:a")
                .arg(if is_rai_like_stream(url) {
                    "copy"
                } else {
                    "aac"
                });
            if !is_rai_like_stream(url) {
                command.arg("-b:a").arg("192k").arg("-ar").arg("48000");
            }
            command
                .arg("-movflags")
                .arg("+frag_keyframe+empty_moov+default_base_moof")
                .arg("-f")
                .arg("mp4");
        }
    }
    command
        .arg("-map_metadata")
        .arg("-1")
        .arg(output_path)
        .stdin(Stdio::piped())
        .stdout(Stdio::null())
        .stderr(Stdio::piped());

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(CREATE_NO_WINDOW);
    }

    let mut child = command.spawn().map_err(|err| {
        recording_format(
            ui_language,
            "radio.recording_error_start_ffmpeg",
            &[("error", err.to_string())],
            "Impossibile avviare FFmpeg per la registrazione: {error}",
        )
    })?;

    if let Some(stderr) = child.stderr.take() {
        let output_for_log = output_path.to_path_buf();
        thread::spawn(move || {
            let reader = BufReader::new(stderr);
            for line in reader.lines().map_while(Result::ok) {
                let line = redact_http_urls(&line);
                if !line.trim().is_empty() {
                    crate::log_debug(&format!(
                        "FFmpeg recording {}: {}",
                        output_for_log.display(),
                        line
                    ));
                }
            }
        });
    }

    Ok(child)
}

fn is_rai_like_stream(url: &str) -> bool {
    let lower = url.to_ascii_lowercase();
    lower.contains("rai") || lower.contains("mediapolis") || lower.contains("relinker")
}

fn monitor_recording_until_player_closes(
    mut recorder: Child,
    mpv_process_id: u32,
    output_path: PathBuf,
) {
    let mut recorder_failed = false;
    loop {
        match recorder.try_wait() {
            Ok(Some(status)) => {
                recorder_failed = !status.success();
                crate::log_debug(&format!(
                    "Stream recording ended before player: status={status} output={}",
                    output_path.display()
                ));
                break;
            }
            Ok(None) => {}
            Err(err) => {
                recorder_failed = true;
                crate::log_debug(&format!("Stream recording status check failed: {err}"));
                break;
            }
        }

        if !is_process_alive(mpv_process_id) {
            stop_recorder(&mut recorder);
            break;
        }
        thread::sleep(Duration::from_millis(500));
    }

    let file_size = fs::metadata(&output_path)
        .map(|metadata| metadata.len())
        .unwrap_or(0);
    let file_too_small = file_size < 1024;
    crate::log_debug(&format!(
        "Stream recording finalized: failed={recorder_failed} bytes={file_size} output={}",
        output_path.display()
    ));
    if recorder_failed || file_too_small {
        crate::log_debug(&format!(
            "Removing unusable stream recording: failed={recorder_failed} output={}",
            output_path.display()
        ));
        remove_recording_file(&output_path, "unusable recording");
    }
}

fn remove_recording_file(path: &Path, context: &str) {
    let Err(err) = fs::remove_file(path) else {
        return;
    };
    if err.kind() != std::io::ErrorKind::NotFound {
        crate::log_debug(&format!(
            "Unable to remove recording file during {context}: path={} error={err}",
            path.display()
        ));
    }
}

fn stop_recorder(recorder: &mut Child) {
    if let Some(mut stdin) = recorder.stdin.take() {
        if let Err(err) = stdin.write_all(b"q\n") {
            crate::log_debug(&format!(
                "Unable to request a graceful FFmpeg shutdown: {err}"
            ));
        }
        if let Err(err) = stdin.flush() {
            crate::log_debug(&format!("Unable to flush FFmpeg stdin: {err}"));
        }
    }
    // MP4 e MP3 hanno bisogno di qualche secondo per scrivere trailer e indici.
    for _ in 0..100 {
        match recorder.try_wait() {
            Ok(Some(_)) => return,
            Ok(None) => thread::sleep(Duration::from_millis(100)),
            Err(_) => break,
        }
    }
    if let Err(err) = recorder.kill() {
        crate::log_debug(&format!("Unable to terminate FFmpeg: {err}"));
    }
    if let Err(err) = recorder.wait() {
        crate::log_debug(&format!("Unable to wait for FFmpeg termination: {err}"));
    }
}

fn redact_http_urls(line: &str) -> String {
    let mut result = line.to_string();
    for scheme in ["https://", "http://"] {
        while let Some(start) = result.find(scheme) {
            let tail = &result[start..];
            let relative_end = tail
                .char_indices()
                .skip(1)
                .find_map(|(index, ch)| {
                    (ch.is_whitespace() || matches!(ch, '\'' | '"' | ']' | ')' | '>'))
                        .then_some(index)
                })
                .unwrap_or(tail.len());
            result.replace_range(start..start + relative_end, "[URL redatto]");
        }
    }
    result
}

fn is_process_alive(process_id: u32) -> bool {
    unsafe {
        let Ok(handle) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, process_id) else {
            return false;
        };
        let mut exit_code = 0u32;
        let result = GetExitCodeProcess(handle, &mut exit_code).is_ok() && exit_code == 259;
        if let Err(err) = CloseHandle(handle) {
            crate::log_debug(&format!("Unable to close process handle: {err}"));
        }
        result
    }
}

fn find_ffmpeg_executable(ui_language: Option<Language>) -> Result<PathBuf, String> {
    let mut candidates = Vec::<PathBuf>::new();
    if let Ok(path) = std::env::var("SONARPAD_FFMPEG_PATH") {
        let path = PathBuf::from(path.trim());
        if !path.as_os_str().is_empty() {
            candidates.push(path);
        }
    }
    if let Ok(exe) = std::env::current_exe()
        && let Some(dir) = exe.parent()
    {
        candidates.push(dir.join("ffmpeg.exe"));
        candidates.push(dir.join("bin").join("ffmpeg.exe"));
        candidates.push(dir.join("deps").join("ffmpeg.exe"));
        candidates.push(dir.join("runtime").join("ffmpeg.exe"));
    }
    candidates.push(PathBuf::from("ffmpeg.exe"));
    candidates.push(PathBuf::from("ffmpeg"));

    for candidate in candidates {
        if candidate.components().count() > 1 && !candidate.is_file() {
            continue;
        }
        let mut command = Command::new(&candidate);
        command
            .arg("-version")
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null());
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(CREATE_NO_WINDOW);
        }
        if command
            .status()
            .map(|status| status.success())
            .unwrap_or(false)
        {
            return Ok(candidate);
        }
    }

    Err(recording_text(
        ui_language,
        "radio.recording_error_ffmpeg",
        "FFmpeg non è stato trovato. Copia ffmpeg.exe accanto a Sonarpad oppure imposta SONARPAD_FFMPEG_PATH.",
    ))
}

fn next_recording_path(
    kind: StreamRecordingKind,
    title: &str,
    ui_language: Option<Language>,
) -> Result<PathBuf, String> {
    let folder = recordings_folder(kind);
    fs::create_dir_all(&folder).map_err(|err| {
        recording_format(
            ui_language,
            "radio.recording_error_create_folder",
            &[("error", err.to_string())],
            "Impossibile creare la cartella delle registrazioni: {error}",
        )
    })?;
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let safe_title = sanitize_file_name(title);
    Ok(folder.join(format!(
        "{}_{}.{}",
        if safe_title.is_empty() {
            match kind {
                StreamRecordingKind::Radio => "Radio",
                StreamRecordingKind::Tv => "TV",
            }
        } else {
            &safe_title
        },
        timestamp,
        kind.extension()
    )))
}

pub(crate) fn recordings_folder(kind: StreamRecordingKind) -> PathBuf {
    let settings = crate::settings::load_settings();
    match kind {
        StreamRecordingKind::Radio => {
            if settings.radio_save_folder.trim().is_empty() {
                PathBuf::from(crate::settings::default_radio_save_folder())
            } else {
                PathBuf::from(settings.radio_save_folder)
            }
        }
        StreamRecordingKind::Tv => {
            if settings.tv_save_folder.trim().is_empty() {
                PathBuf::from(crate::settings::default_tv_save_folder())
            } else {
                PathBuf::from(settings.tv_save_folder)
            }
        }
    }
}

fn mixed_legacy_recordings_folder() -> PathBuf {
    let settings = crate::settings::load_settings();
    if settings.podcast_save_folder.trim().is_empty() {
        PathBuf::from(crate::settings::default_podcast_save_folder())
    } else {
        PathBuf::from(settings.podcast_save_folder)
    }
}

fn legacy_recordings_folder(kind: StreamRecordingKind) -> PathBuf {
    crate::settings::settings_dir().join(kind.folder_name())
}

fn sanitize_file_name(value: &str) -> String {
    let mut result = value
        .trim()
        .chars()
        .map(|ch| match ch {
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
            ch if ch.is_control() => '_',
            ch => ch,
        })
        .collect::<String>();
    while result.ends_with(' ') || result.ends_with('.') {
        result.pop();
    }
    result.chars().take(100).collect()
}

pub(crate) fn open_recordings(parent: HWND, language: Language, kind: StreamRecordingKind) {
    let mut selected_id: Option<String> = None;
    let mut close_silently_if_empty = false;
    loop {
        let files = list_recordings(kind);
        if files.is_empty() {
            if close_silently_if_empty {
                crate::log_debug(
                    "Recordings list is empty after deletion; closing without an error.",
                );
                return;
            }
            let text = match kind {
                StreamRecordingKind::Radio => translated(
                    language,
                    "radio.recordings_empty",
                    "Non ci sono registrazioni radio.",
                ),
                StreamRecordingKind::Tv => "Non ci sono registrazioni TV.".to_string(),
            };
            let title = translated(language, "app.warning_title", "Attenzione");
            let title_wide = to_wide(&title);
            let text_wide = to_wide(&text);
            crate::message_box_modal(
                parent,
                PCWSTR(text_wide.as_ptr()),
                PCWSTR(title_wide.as_ptr()),
                MB_OK | MB_ICONWARNING,
            );
            return;
        }

        let items = files
            .iter()
            .map(|path| MultilineSelectionItem {
                id: path.to_string_lossy().to_string(),
                title: path
                    .file_stem()
                    .and_then(|value| value.to_str())
                    .unwrap_or("Registrazione")
                    .to_string(),
                description: fs::metadata(path)
                    .ok()
                    .map(|metadata| format_file_size(metadata.len())),
            })
            .collect::<Vec<_>>();

        let delete_label = match kind {
            StreamRecordingKind::Radio => {
                translated(language, "radio.delete_recording", "Elimina registrazione")
            }
            StreamRecordingKind::Tv => "Elimina registrazione".to_string(),
        };
        let recording_deleted = Arc::new(AtomicBool::new(false));
        let recording_deleted_handler = Arc::clone(&recording_deleted);
        let delete_action = InterpreterContextAction {
            label: delete_label,
            ctrl_c_shortcut: false,
            enabled: Arc::new(|id| Path::new(id).is_file()),
            handler: Arc::new(move |id| match fs::remove_file(&id) {
                Ok(()) => {
                    recording_deleted_handler.store(true, Ordering::SeqCst);
                    crate::log_debug(&format!(
                        "Recording deleted; refreshing list immediately: {id}"
                    ));
                    let dialog = crate::get_foreground_window_safe();
                    if dialog.0 != 0 && crate::is_window_handle_valid(dialog) {
                        crate::log_if_err!(crate::post_message_w_safe(
                            dialog,
                            windows::Win32::UI::WindowsAndMessaging::WM_CLOSE,
                            windows::Win32::Foundation::WPARAM(0),
                            windows::Win32::Foundation::LPARAM(0),
                        ));
                    }
                }
                Err(err) => {
                    crate::log_debug(&format!("Recording delete failed for {id}: {err}"));
                }
            }),
        };

        let title = match kind {
            StreamRecordingKind::Radio => {
                translated(language, "radio.recordings", "Registrazioni radio")
            }
            StreamRecordingKind::Tv => "Registrazioni TV".to_string(),
        };
        let result = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            title,
            items,
            selected_id.clone(),
            MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: translated(language, "radio.search", "Ricerca"),
                show_search_edit: false,
                secondary_action_label: None,
                context_actions: vec![delete_action],
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: kind == StreamRecordingKind::Tv,
            },
        );

        if recording_deleted.swap(false, Ordering::SeqCst) {
            selected_id = None;
            close_silently_if_empty = true;
            continue;
        }
        close_silently_if_empty = false;

        match result {
            MultilineSelectionResult::Selected(path) => {
                selected_id = Some(path.clone());
                let playback_result = match kind {
                    StreamRecordingKind::Radio => {
                        crate::launch_local_video_in_mpv(parent, Path::new(&path))
                    }
                    StreamRecordingKind::Tv => {
                        crate::launch_local_tv_recording_in_mpv(parent, Path::new(&path))
                    }
                };
                if let Err(err) = playback_result {
                    crate::show_error(parent, language, &err);
                }
                if kind == StreamRecordingKind::Radio {
                    return;
                }
                // For TV, reopen the recordings list on the same item. Esc while
                // mpv is active stops playback and returns focus here; a second
                // Esc closes the recordings list.
                continue;
            }
            MultilineSelectionResult::Cancelled => return,
            MultilineSelectionResult::Search(_) | MultilineSelectionResult::SecondaryAction => {}
        }
    }
}

fn list_recordings(kind: StreamRecordingKind) -> Vec<PathBuf> {
    let mut folders = vec![recordings_folder(kind)];
    for legacy in [
        mixed_legacy_recordings_folder(),
        PathBuf::from(crate::settings::default_podcast_save_folder()),
        legacy_recordings_folder(kind),
    ] {
        if !folders.iter().any(|folder| folder == &legacy) {
            folders.push(legacy);
        }
    }

    let mut files = folders
        .into_iter()
        .flat_map(|folder| {
            fs::read_dir(folder)
                .ok()
                .into_iter()
                .flatten()
                .filter_map(Result::ok)
                .map(|entry| entry.path())
                .collect::<Vec<_>>()
        })
        .filter(|path| path.is_file())
        .filter(|path| !is_podcast_recording_file(path))
        .filter(|path| {
            let extension = path
                .extension()
                .and_then(|value| value.to_str())
                .map(str::to_ascii_lowercase);
            match (kind, extension.as_deref()) {
                (StreamRecordingKind::Radio, Some(extension)) => {
                    matches!(extension, "mp3" | "mka" | "aac" | "wav")
                }
                (StreamRecordingKind::Tv, Some(extension)) => {
                    matches!(extension, "mp4" | "mkv" | "ts")
                }
                _ => false,
            }
        })
        .collect::<Vec<_>>();
    files.sort_by(|a, b| {
        fs::metadata(b)
            .and_then(|metadata| metadata.modified())
            .ok()
            .cmp(
                &fs::metadata(a)
                    .and_then(|metadata| metadata.modified())
                    .ok(),
            )
    });
    files
}

fn is_podcast_recording_file(path: &Path) -> bool {
    path.file_name()
        .and_then(|value| value.to_str())
        .map(|name| name.to_ascii_lowercase().starts_with("podcast_"))
        .unwrap_or(false)
}

fn format_file_size(bytes: u64) -> String {
    if bytes >= 1_073_741_824 {
        format!("{:.2} GB", bytes as f64 / 1_073_741_824.0)
    } else if bytes >= 1_048_576 {
        format!("{:.1} MB", bytes as f64 / 1_048_576.0)
    } else if bytes >= 1024 {
        format!("{:.1} KB", bytes as f64 / 1024.0)
    } else {
        format!("{bytes} byte")
    }
}

fn recording_text(language: Option<Language>, key: &str, fallback: &str) -> String {
    language
        .map(|language| translated(language, key, fallback))
        .unwrap_or_else(|| fallback.to_string())
}

fn recording_format(
    language: Option<Language>,
    key: &str,
    args: &[(&str, String)],
    fallback: &str,
) -> String {
    if let Some(language) = language {
        let borrowed = args
            .iter()
            .map(|(name, value)| (*name, value.as_str()))
            .collect::<Vec<_>>();
        let value = crate::i18n::tr_f(language, key, &borrowed);
        if value != key {
            return value;
        }
    }
    let mut value = fallback.to_string();
    for (name, replacement) in args {
        value = value.replace(&format!("{{{name}}}"), replacement);
    }
    value
}

fn translated(language: Language, key: &str, fallback: &str) -> String {
    let value = crate::i18n::tr(language, key);
    if value == key {
        fallback.to_string()
    } else {
        value
    }
}
