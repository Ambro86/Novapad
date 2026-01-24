use crate::log_debug;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;

pub const SUBTITLE_EXTENSIONS: &[&str] = &[
    "srt", "vtt", "ass", "ssa", "sub", "sbv", "lrc", "smi", "sami",
];

#[derive(Clone)]
pub struct SubtitleCue {
    pub start: Duration,
    pub end: Duration,
    pub text: String,
    pub audio_path: Option<PathBuf>,
}

pub fn find_subtitle_for_media(media_path: &Path) -> Option<PathBuf> {
    let stem = media_path.file_stem()?.to_string_lossy();
    let dir = media_path.parent()?;
    for ext in SUBTITLE_EXTENSIONS {
        let candidate = dir.join(format!("{stem}.{ext}"));
        if candidate.exists() {
            return Some(candidate);
        }
    }
    None
}

pub fn load_subtitles(path: &Path) -> Result<Vec<SubtitleCue>, String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    let content = fs::read_to_string(path).map_err(|e| e.to_string())?;
    let cues = match ext.as_str() {
        "srt" => parse_srt(&content),
        "vtt" => parse_vtt(&content),
        "ass" | "ssa" => parse_ass(&content),
        "sub" => parse_sub(&content),
        "sbv" => parse_sbv(&content),
        "lrc" => parse_lrc(&content),
        "smi" | "sami" => parse_sami(&content),
        _ => parse_unknown(&content),
    };
    if cues.is_empty() {
        return Err(format!("Unsupported subtitle format: {}", ext));
    }
    Ok(cues)
}

fn parse_srt(input: &str) -> Vec<SubtitleCue> {
    let mut cues = Vec::new();
    let mut lines = input.lines().peekable();
    while let Some(line) = lines.next() {
        if line.trim().is_empty() {
            continue;
        }
        let time_line = if line.contains("-->") {
            line
        } else {
            match lines.next() {
                Some(next) => next,
                None => break,
            }
        };
        let (start, end) = match parse_time_range(time_line, ',', true) {
            Some(range) => range,
            None => continue,
        };
        let mut text_lines = Vec::new();
        while let Some(text_line) = lines.peek().copied() {
            if text_line.trim().is_empty() {
                lines.next();
                break;
            }
            text_lines.push(text_line);
            lines.next();
        }
        let text = strip_tags_simple(&text_lines.join("\n"));
        let text = text.trim().to_string();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    cues
}

fn parse_vtt(input: &str) -> Vec<SubtitleCue> {
    let mut cues = Vec::new();
    let mut lines = input.lines().peekable();
    while let Some(line) = lines.next() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with("WEBVTT") {
            continue;
        }
        let time_line = if trimmed.contains("-->") {
            trimmed
        } else {
            match lines.next() {
                Some(next) => next.trim(),
                None => break,
            }
        };
        let (start, end) = match parse_time_range(time_line, '.', false) {
            Some(range) => range,
            None => continue,
        };
        let mut text_lines = Vec::new();
        while let Some(text_line) = lines.peek().copied() {
            if text_line.trim().is_empty() {
                lines.next();
                break;
            }
            text_lines.push(text_line.trim_end());
            lines.next();
        }
        let text = strip_tags_simple(&text_lines.join("\n"));
        let text = text.trim().to_string();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    cues
}

fn parse_ass(input: &str) -> Vec<SubtitleCue> {
    let mut cues = Vec::new();
    for line in input.lines() {
        let trimmed = line.trim();
        if !trimmed.to_lowercase().starts_with("dialogue:") {
            continue;
        }
        let mut parts = trimmed.splitn(10, ',');
        let _prefix = parts.next();
        let start = match parts.next().and_then(parse_time_ass) {
            Some(v) => v,
            None => continue,
        };
        let end = match parts.next().and_then(parse_time_ass) {
            Some(v) => v,
            None => continue,
        };
        let text = parts.nth(6).unwrap_or_default();
        let text = strip_ass_tags(text);
        let text = text.replace("\\N", "\n").trim().to_string();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    cues
}

fn parse_sub(input: &str) -> Vec<SubtitleCue> {
    let mut cues = Vec::new();
    let mut fps = None;
    for line in input.lines() {
        let trimmed = line.trim();
        if !trimmed.starts_with('{') {
            continue;
        }
        let mut rest = trimmed;
        let start = match parse_microdvd_brace(&mut rest) {
            Some(v) => v,
            None => continue,
        };
        let end = match parse_microdvd_brace(&mut rest) {
            Some(v) => v,
            None => continue,
        };
        let text = rest.trim();
        if start == 1
            && end == 1
            && let Ok(parsed) = text.parse::<f32>()
        {
            fps = Some(parsed.max(1.0));
            continue;
        }
        let fps = fps.unwrap_or(25.0);
        let start = Duration::from_secs_f64(start as f64 / fps as f64);
        let end = Duration::from_secs_f64(end as f64 / fps as f64);
        let text = text.replace('|', "\n").trim().to_string();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    if fps.is_none() {
        log_debug("Subtitle: MicroDVD FPS not found, assuming 25 fps.");
    }
    cues
}

fn parse_sbv(input: &str) -> Vec<SubtitleCue> {
    let mut cues = Vec::new();
    let mut lines = input.lines().peekable();
    while let Some(line) = lines.next() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        let mut parts = trimmed.split(',');
        let start = match parts
            .next()
            .map(str::trim)
            .and_then(|t| parse_timestamp(t, '.', true))
        {
            Some(v) => v,
            None => continue,
        };
        let end = match parts
            .next()
            .map(str::trim)
            .and_then(|t| parse_timestamp(t, '.', true))
        {
            Some(v) => v,
            None => continue,
        };
        let mut text_lines = Vec::new();
        while let Some(text_line) = lines.peek().copied() {
            if text_line.trim().is_empty() {
                lines.next();
                break;
            }
            text_lines.push(text_line.trim_end());
            lines.next();
        }
        let text = strip_tags_simple(&text_lines.join("\n"));
        let text = text.trim().to_string();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    cues
}

fn parse_lrc(input: &str) -> Vec<SubtitleCue> {
    let mut items: Vec<(Duration, String)> = Vec::new();
    for line in input.lines() {
        let mut times = Vec::new();
        let mut rest = line;
        while let Some(start) = rest.find('[') {
            let after = &rest[start + 1..];
            let end = match after.find(']') {
                Some(v) => v,
                None => break,
            };
            let time_str = &after[..end];
            if let Some(t) = parse_lrc_timestamp(time_str) {
                times.push(t);
            }
            rest = &after[end + 1..];
        }
        let text = strip_tags_simple(rest).trim().to_string();
        if text.is_empty() {
            continue;
        }
        for t in times {
            items.push((t, text.clone()));
        }
    }
    items.sort_by_key(|(t, _)| *t);
    let mut cues = Vec::new();
    for idx in 0..items.len() {
        let start = items[idx].0;
        let end = if idx + 1 < items.len() {
            items[idx + 1].0
        } else {
            start + Duration::from_secs(3)
        };
        let text = items[idx].1.clone();
        if !text.is_empty() {
            cues.push(SubtitleCue {
                start,
                end,
                text,
                audio_path: None,
            });
        }
    }
    cues
}

fn parse_sami(input: &str) -> Vec<SubtitleCue> {
    let lower = input.to_lowercase();
    let mut entries: Vec<(Duration, String)> = Vec::new();
    let mut cursor = 0usize;
    while let Some(sync_pos) = lower[cursor..].find("<sync") {
        let sync_pos = cursor + sync_pos;
        let tag_end = match lower[sync_pos..].find('>') {
            Some(v) => sync_pos + v,
            None => break,
        };
        let tag = &lower[sync_pos..=tag_end];
        let start_ms = parse_sami_start_ms(tag);
        let Some(start_ms) = start_ms else {
            cursor = tag_end + 1;
            continue;
        };
        let text_start = tag_end + 1;
        let next_sync = lower[text_start..]
            .find("<sync")
            .map(|v| text_start + v)
            .unwrap_or(lower.len());
        let raw_text = &input[text_start..next_sync];
        let text = strip_tags_simple(raw_text)
            .replace("&nbsp;", " ")
            .trim()
            .to_string();
        if !text.is_empty() {
            entries.push((Duration::from_millis(start_ms), text));
        }
        cursor = next_sync;
    }
    entries.sort_by_key(|(t, _)| *t);
    let mut cues = Vec::new();
    for idx in 0..entries.len() {
        let start = entries[idx].0;
        let end = if idx + 1 < entries.len() {
            entries[idx + 1].0
        } else {
            start + Duration::from_secs(3)
        };
        let text = entries[idx].1.clone();
        cues.push(SubtitleCue {
            start,
            end,
            text,
            audio_path: None,
        });
    }
    cues
}

fn parse_time_range(line: &str, sep: char, allow_hours: bool) -> Option<(Duration, Duration)> {
    let mut parts = line.split("-->");
    let start = parts.next()?.trim();
    let end = parts.next()?.split_whitespace().next().unwrap_or("");
    let start = parse_timestamp(start, sep, allow_hours)?;
    let end = parse_timestamp(end, sep, allow_hours)?;
    Some((start, end))
}

fn parse_timestamp(text: &str, sep: char, allow_hours: bool) -> Option<Duration> {
    let mut time_parts: Vec<&str> = text.split(':').collect();
    if !allow_hours && time_parts.len() == 2 {
        time_parts.insert(0, "0");
    }
    if time_parts.len() != 3 {
        return None;
    }
    let hours = time_parts[0].parse::<u64>().ok()?;
    let minutes = time_parts[1].parse::<u64>().ok()?;
    let sec_parts: Vec<&str> = time_parts[2].split(sep).collect();
    let seconds = sec_parts.first()?.parse::<u64>().ok()?;
    let millis = sec_parts
        .get(1)
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(0);
    let total = hours * 3600 + minutes * 60 + seconds;
    Some(Duration::from_millis(total * 1000 + millis))
}

fn parse_time_ass(text: &str) -> Option<Duration> {
    let parts: Vec<&str> = text.trim().split(':').collect();
    if parts.len() != 3 {
        return None;
    }
    let hours = parts[0].parse::<u64>().ok()?;
    let minutes = parts[1].parse::<u64>().ok()?;
    let sec_parts: Vec<&str> = parts[2].split('.').collect();
    let seconds = sec_parts.first()?.parse::<u64>().ok()?;
    let centis = sec_parts
        .get(1)
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(0);
    let total = hours * 3600 + minutes * 60 + seconds;
    Some(Duration::from_millis(total * 1000 + centis * 10))
}

fn strip_ass_tags(text: &str) -> String {
    let mut out = String::new();
    let mut in_tag = false;
    for ch in text.chars() {
        match ch {
            '{' => in_tag = true,
            '}' => in_tag = false,
            _ if !in_tag => out.push(ch),
            _ => {}
        }
    }
    out
}

fn strip_tags_simple(text: &str) -> String {
    let mut out = String::new();
    let mut in_tag = false;
    let chars: Vec<char> = text.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        let ch = chars[i];
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            '&' => {
                if let Some((entity, consumed)) = parse_entity(&chars, i + 1) {
                    out.push_str(&entity);
                    i += consumed;
                } else if !in_tag {
                    out.push(ch);
                }
            }
            _ if !in_tag => out.push(ch),
            _ => {}
        }
        i += 1;
    }
    out
}

fn parse_entity(chars: &[char], start: usize) -> Option<(String, usize)> {
    let mut entity = String::new();
    let mut i = start;
    while i < chars.len() {
        let ch = chars[i];
        if ch == ';' {
            break;
        }
        if entity.len() > 8 {
            return None;
        }
        entity.push(ch);
        i += 1;
    }
    if i >= chars.len() || chars[i] != ';' {
        return None;
    }
    let decoded = match entity.as_str() {
        "amp" => "&",
        "lt" => "<",
        "gt" => ">",
        "quot" => "\"",
        "apos" => "'",
        "nbsp" => " ",
        _ => return None,
    };
    let consumed = i - start + 1;
    Some((decoded.to_string(), consumed))
}

fn parse_microdvd_brace(input: &mut &str) -> Option<u64> {
    if !input.starts_with('{') {
        return None;
    }
    let end = input.find('}')?;
    let num = &input[1..end];
    *input = &input[end + 1..];
    num.parse::<u64>().ok()
}

fn parse_unknown(input: &str) -> Vec<SubtitleCue> {
    if input.contains("WEBVTT") || input.contains("-->") {
        let cues = parse_vtt(input);
        if !cues.is_empty() {
            return cues;
        }
        let cues = parse_srt(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    if input.to_lowercase().contains("dialogue:") {
        let cues = parse_ass(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    if input.contains('{') && input.contains('}') {
        let cues = parse_sub(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    if input.to_lowercase().contains("<sync") {
        let cues = parse_sami(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    if input.contains('[') && input.contains(']') {
        let cues = parse_lrc(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    if input.contains(',') && input.contains(':') {
        let cues = parse_sbv(input);
        if !cues.is_empty() {
            return cues;
        }
    }
    Vec::new()
}

fn parse_lrc_timestamp(text: &str) -> Option<Duration> {
    let mut parts = text.split(':');
    let minutes = parts.next()?.parse::<u64>().ok()?;
    let seconds_part = parts.next().unwrap_or("0");
    let mut sec_parts = seconds_part.split('.');
    let seconds = sec_parts.next().unwrap_or("0").parse::<u64>().ok()?;
    let hundredths = sec_parts
        .next()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(0);
    let millis = if hundredths >= 100 {
        hundredths
    } else {
        hundredths * 10
    };
    let total = minutes * 60 + seconds;
    Some(Duration::from_millis(total * 1000 + millis))
}

fn parse_sami_start_ms(tag: &str) -> Option<u64> {
    let start_idx = tag.find("start")?;
    let after = &tag[start_idx + 5..];
    let eq_idx = after.find('=')?;
    let mut digits = String::new();
    for ch in after[eq_idx + 1..].chars() {
        if ch.is_ascii_digit() {
            digits.push(ch);
        } else if !digits.is_empty() {
            break;
        }
    }
    digits.parse::<u64>().ok()
}
