use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Clone, Serialize, Deserialize)]
pub struct Bookmark {
    pub position: i32,
    pub snippet: String,
    pub timestamp: String,
    #[serde(default)]
    pub automatic: bool,
}

#[derive(Default, Serialize, Deserialize)]
pub struct BookmarkStore {
    pub files: HashMap<String, Vec<Bookmark>>,
}

impl Bookmark {
    pub fn manual(position: i32, snippet: String, timestamp: String) -> Self {
        Self {
            position,
            snippet,
            timestamp,
            automatic: false,
        }
    }

    pub fn automatic(position: i32, snippet: String, timestamp: String) -> Self {
        Self {
            position,
            snippet,
            timestamp,
            automatic: true,
        }
    }

    pub fn is_visible(&self, automatic_bookmark_enabled: bool) -> bool {
        automatic_bookmark_enabled || !self.automatic
    }
}

pub fn sort_bookmarks(bookmarks: &mut [Bookmark]) {
    bookmarks.sort_by_key(|bookmark| bookmark.position);
}

pub fn normalize_store(store: &mut BookmarkStore) {
    for bookmarks in store.files.values_mut() {
        sort_bookmarks(bookmarks);
    }
}

fn bookmark_store_path() -> Option<PathBuf> {
    let mut path = crate::settings::settings_dir();
    path.push("bookmarks.json");
    Some(path)
}

pub fn load_bookmarks() -> BookmarkStore {
    let Some(path) = bookmark_store_path() else {
        return BookmarkStore::default();
    };
    let data = std::fs::read_to_string(path).ok();
    let Some(data) = data else {
        return BookmarkStore::default();
    };
    let mut store = serde_json::from_str(&data).unwrap_or_default();
    normalize_store(&mut store);
    store
}

pub fn save_bookmarks(store: &BookmarkStore) {
    let Some(path) = bookmark_store_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        crate::log_if_err!(std::fs::create_dir_all(parent));
    }
    let mut normalized = BookmarkStore {
        files: store.files.clone(),
    };
    normalize_store(&mut normalized);
    if let Ok(json) = serde_json::to_string_pretty(&normalized) {
        crate::log_if_err!(std::fs::write(path, json));
    }
}

#[cfg(test)]
mod tests {
    use super::Bookmark;

    #[test]
    fn legacy_bookmarks_default_to_manual() {
        let bookmark: Bookmark = serde_json::from_str(
            r#"{"position":42,"snippet":"sample","timestamp":"2026-04-29 10:00:00"}"#,
        )
        .expect("legacy bookmark should deserialize");

        assert!(!bookmark.automatic);
        assert!(bookmark.is_visible(false));
    }
}
