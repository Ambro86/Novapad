use serde::Deserialize;
use serde_json::Value;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Chapter {
    pub start_ms: u64,
    pub title: String,
    pub url: Option<String>,
    pub image: Option<String>,
}

#[derive(Deserialize)]
struct ChaptersFile {
    chapters: Option<Vec<ChapterEntry>>,
}

#[derive(Deserialize)]
struct ChapterEntry {
    #[serde(rename = "startTime")]
    start_time: Option<Value>,
    title: Option<String>,
    url: Option<String>,
    #[serde(rename = "img")]
    image: Option<String>,
}

pub fn parse_chapters_json(bytes: &[u8]) -> Vec<Chapter> {
    let parsed = serde_json::from_slice::<ChaptersFile>(bytes).ok();
    let Some(parsed) = parsed else {
        return Vec::new();
    };
    let Some(entries) = parsed.chapters else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for entry in entries {
        let Some(start_value) = entry.start_time.as_ref() else {
            continue;
        };
        let Some(start_ms) = parse_start_time(start_value) else {
            continue;
        };
        let Some(title) = entry
            .title
            .as_ref()
            .map(|t| t.trim())
            .filter(|t| !t.is_empty())
        else {
            continue;
        };
        out.push(Chapter {
            start_ms,
            title: title.to_string(),
            url: entry
                .url
                .as_ref()
                .map(|u| u.trim())
                .filter(|u| !u.is_empty())
                .map(|u| u.to_string()),
            image: entry
                .image
                .as_ref()
                .map(|i| i.trim())
                .filter(|i| !i.is_empty())
                .map(|i| i.to_string()),
        });
    }
    out.sort_by_key(|chapter| chapter.start_ms);
    out
}

pub fn current_chapter_index(current_pos_ms: u64, chapters: &[Chapter]) -> Option<usize> {
    if chapters.is_empty() {
        return None;
    }
    let mut current = None;
    for (idx, chapter) in chapters.iter().enumerate() {
        if chapter.start_ms <= current_pos_ms {
            current = Some(idx);
        } else {
            break;
        }
    }
    current
}

pub fn chapter_label(chapter: &Chapter) -> String {
    let seconds = chapter.start_ms / 1000;
    format!("{}  {}", format_time_hms(seconds), chapter.title)
}

fn parse_start_time(value: &Value) -> Option<u64> {
    match value {
        Value::Number(num) => {
            let seconds = num.as_f64()?;
            if seconds < 0.0 {
                return None;
            }
            Some((seconds * 1000.0).floor() as u64)
        }
        Value::String(text) => parse_time_string(text),
        _ => None,
    }
}

fn parse_time_string(text: &str) -> Option<u64> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }
    let parts: Vec<&str> = trimmed.split(':').collect();
    let total_seconds = match parts.len() {
        1 => parts[0].parse::<f64>().ok()?,
        2 => {
            let minutes = parts[0].parse::<f64>().ok()?;
            let seconds = parts[1].parse::<f64>().ok()?;
            minutes * 60.0 + seconds
        }
        3 => {
            let hours = parts[0].parse::<f64>().ok()?;
            let minutes = parts[1].parse::<f64>().ok()?;
            let seconds = parts[2].parse::<f64>().ok()?;
            hours * 3600.0 + minutes * 60.0 + seconds
        }
        _ => return None,
    };
    if total_seconds < 0.0 {
        return None;
    }
    Some((total_seconds * 1000.0).floor() as u64)
}

fn format_time_hms(seconds: u64) -> String {
    let hours = seconds / 3600;
    let minutes = (seconds % 3600) / 60;
    let secs = seconds % 60;
    if hours > 0 {
        format!("{:02}:{:02}:{:02}", hours, minutes, secs)
    } else {
        format!("{:02}:{:02}", minutes, secs)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_chapters_json_formats() {
        let json = r#"
        {
            "chapters": [
                { "startTime": 0, "title": "Intro" },
                { "startTime": "00:01:30", "title": "Segment A", "url": "http://example.com" },
                { "startTime": "1:02:03.500", "title": "Segment B", "img": "http://img" },
                { "startTime": -5, "title": "Bad" },
                { "startTime": "bad", "title": "Bad2" },
                { "title": "Missing time" }
            ]
        }
        "#;
        let chapters = parse_chapters_json(json.as_bytes());
        assert_eq!(chapters.len(), 3);
        assert_eq!(chapters[0].start_ms, 0);
        assert_eq!(chapters[1].start_ms, 90_000);
        assert_eq!(chapters[2].start_ms, 3_723_500);
        assert_eq!(chapters[1].url.as_deref(), Some("http://example.com"));
        assert_eq!(chapters[2].image.as_deref(), Some("http://img"));
    }

    #[test]
    fn current_chapter_index_edges() {
        let chapters = vec![
            Chapter {
                start_ms: 0,
                title: "One".to_string(),
                url: None,
                image: None,
            },
            Chapter {
                start_ms: 10_000,
                title: "Two".to_string(),
                url: None,
                image: None,
            },
            Chapter {
                start_ms: 20_000,
                title: "Three".to_string(),
                url: None,
                image: None,
            },
        ];
        assert_eq!(current_chapter_index(0, &chapters), Some(0));
        assert_eq!(current_chapter_index(9_999, &chapters), Some(0));
        assert_eq!(current_chapter_index(10_000, &chapters), Some(1));
        assert_eq!(current_chapter_index(25_000, &chapters), Some(2));
        assert_eq!(current_chapter_index(500, &[]), None);
    }
}
