use crate::settings::FileFormat;
use crate::with_state;
use rodio::{Decoder, OutputStream, Sink, Source};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Duration;
use windows::Win32::Foundation::HWND;

pub struct AudiobookPlayer {
    pub path: PathBuf,
    pub sink: Arc<Sink>,
    pub _stream: OutputStream, // Deve essere mantenuto in vita
    pub is_paused: bool,
    pub start_instant: std::time::Instant,
    pub accumulated_seconds: u64,
    pub volume: f32,
    pub muted: bool,
    pub prev_volume: f32,
}

pub fn parse_time_input(input: &str) -> Result<u64, String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err("empty".to_string());
    }
    if trimmed.chars().all(|c| c.is_ascii_digit()) {
        return trimmed.parse::<u64>().map_err(|_| "invalid".to_string());
    }
    if trimmed.contains(':') {
        let parts: Vec<&str> = trimmed.split(':').collect();
        if parts.len() == 2 || parts.len() == 3 {
            let mut nums = Vec::with_capacity(parts.len());
            for part in parts {
                let part = part.trim();
                if part.is_empty() || !part.chars().all(|c| c.is_ascii_digit()) {
                    return Err("invalid".to_string());
                }
                nums.push(part.parse::<u64>().map_err(|_| "invalid".to_string())?);
            }
            if nums.len() == 2 {
                let minutes = nums[0];
                let seconds = nums[1];
                if seconds >= 60 {
                    return Err("invalid".to_string());
                }
                return Ok(minutes * 60 + seconds);
            }
            let hours = nums[0];
            let minutes = nums[1];
            let seconds = nums[2];
            if minutes >= 60 || seconds >= 60 {
                return Err("invalid".to_string());
            }
            return Ok(hours * 3600 + minutes * 60 + seconds);
        }
    }
    Err("invalid".to_string())
}

pub fn audiobook_duration_secs(path: &Path) -> Option<u64> {
    let file = std::fs::File::open(path).ok()?;
    let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).ok()?;
    if let Some(dur) = source.total_duration() {
        return Some(dur.as_secs());
    }
    mp3_duration::from_path(path).ok().map(|d| d.as_secs())
}

pub unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    let path_buf = path.to_path_buf();

    let bookmark_pos = with_state(hwnd, |state| {
        state
            .bookmarks
            .files
            .get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last()) // Usa l'ultimo segnalibro per l'audio
            .map(|bm| bm.position)
            .unwrap_or(0)
    })
    .unwrap_or(0);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };

        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        // Salta alla posizione del segnalibro se presente
        if bookmark_pos > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(bookmark_pos as u64));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: bookmark_pos as u64,
            volume: 1.0,
            muted: false,
            prev_volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    let start_action = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                player.sink.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
                player.sink.pause();
                player.is_paused = true;
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
            }
            return None;
        }

        let doc = state.docs.get(state.current)?;
        if !matches!(doc.format, FileFormat::Audiobook) {
            return None;
        }
        let path = doc.path.clone()?;
        let from_start = state
            .last_stopped_audiobook
            .as_ref()
            .map(|p| p == &path)
            .unwrap_or(false);
        if from_start {
            state.last_stopped_audiobook = None;
        }
        Some((path, from_start))
    })
    .flatten();

    if let Some((path, from_start)) = start_action {
        if from_start {
            start_audiobook_at(hwnd, &path, 0);
        } else {
            start_audiobook_playback(hwnd, &path);
        }
    }
}

pub unsafe fn seek_audiobook(hwnd: HWND, seconds: i64) {
    let result = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if !player.is_paused {
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (player.accumulated_seconds as i64 + seconds).max(0);
            player.accumulated_seconds = new_pos as u64;
            Some((player.path.clone(), new_pos))
        } else {
            None
        }
    })
    .flatten();

    let (path, current_pos) = match result {
        Some(v) => v,
        None => return,
    };

    stop_audiobook_playback(hwnd);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = OutputStream::try_default().unwrap();
        let sink: Arc<Sink> = Arc::new(Sink::try_new(&handle).unwrap());
        let file = std::fs::File::open(&path).unwrap();
        let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).unwrap();

        let skipped = source.skip_duration(Duration::from_secs(current_pos as u64));
        sink.append(skipped);

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: current_pos as u64,
            volume: 1.0,
            muted: false,
            prev_volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn seek_audiobook_to(hwnd: HWND, seconds: u64) -> Result<(), String> {
    let path = with_state(hwnd, |state| {
        state
            .active_audiobook
            .as_ref()
            .map(|player| player.path.clone())
    })
    .flatten()
    .ok_or_else(|| "No active audiobook".to_string())?;

    start_audiobook_at(hwnd, &path, seconds);
    Ok(())
}

pub unsafe fn stop_audiobook_playback(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            state.last_stopped_audiobook = Some(player.path.clone());
            player.sink.stop();
        }
    });
}

pub unsafe fn start_audiobook_at(hwnd: HWND, path: &Path, seconds: u64) {
    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    let hwnd_main = hwnd;

    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };

        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        if seconds > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(seconds));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: 1.0,
            muted: false,
            prev_volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn change_audiobook_volume(hwnd: HWND, delta: f32) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                player.prev_volume = (player.prev_volume + delta).clamp(0.0, 3.0);
                return;
            }
            player.volume = (player.volume + delta).clamp(0.0, 3.0);
            player.sink.set_volume(player.volume);
        }
    });
}

pub unsafe fn audiobook_volume_level(hwnd: HWND) -> Option<f32> {
    with_state(hwnd, |state| {
        state
            .active_audiobook
            .as_ref()
            .map(|player| if player.muted { 0.0 } else { player.volume })
    })
    .flatten()
}

pub unsafe fn toggle_audiobook_mute(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.muted {
                let restored = if player.prev_volume > 0.0 {
                    player.prev_volume
                } else {
                    1.0
                };
                player.volume = restored;
                player.muted = false;
                player.sink.set_volume(player.volume);
            } else {
                if player.volume > 0.0 {
                    player.prev_volume = player.volume;
                }
                player.volume = 0.0;
                player.muted = true;
                player.sink.set_volume(0.0);
            }
        }
    });
}

#[cfg(test)]
mod tests {
    use super::parse_time_input;

    #[test]
    fn parse_seconds() {
        assert_eq!(parse_time_input("90").unwrap(), 90);
    }

    #[test]
    fn parse_mm_ss() {
        assert_eq!(parse_time_input("01:30").unwrap(), 90);
        assert_eq!(parse_time_input("10:00").unwrap(), 600);
    }

    #[test]
    fn parse_hh_mm_ss() {
        assert_eq!(parse_time_input("00:01:30").unwrap(), 90);
    }

    #[test]
    fn parse_invalid() {
        assert!(parse_time_input("").is_err());
        assert!(parse_time_input("abc").is_err());
        assert!(parse_time_input("1:99").is_err());
        assert!(parse_time_input("1:2:99").is_err());
        assert!(parse_time_input("1:2:3:4").is_err());
    }
}
