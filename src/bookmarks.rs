use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Clone, Serialize, Deserialize)]
pub struct Bookmark {
    pub position: i32,
    pub snippet: String,
    pub timestamp: String,
}

#[derive(Default, Serialize, Deserialize)]
pub struct BookmarkStore {
    pub files: HashMap<String, Vec<Bookmark>>,
}

fn bookmark_store_path() -> Option<PathBuf> {
    let base = std::env::var_os("APPDATA")?;
    let mut path = PathBuf::from(base);
    path.push("Novapad");
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
    serde_json::from_str(&data).unwrap_or_default()
}

pub fn save_bookmarks(store: &BookmarkStore) {
    let Some(path) = bookmark_store_path() else {
        return;
    };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(store) {
        let _ = std::fs::write(path, json);
    }
}
