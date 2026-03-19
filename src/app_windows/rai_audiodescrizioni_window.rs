use std::collections::HashSet;

use windows::Win32::Foundation::HWND;

use crate::app_windows::interpreter_select_window;
use crate::settings::Language;
use crate::tools::rai_audiodescrizioni::{self, CatalogItem};
use crate::{show_error, with_state};

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }

    crate::screen_reader_speak("Caricamento audiodescrizioni Rai");
    let catalog = match rai_audiodescrizioni::load_catalog() {
        Ok(catalog) => catalog,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };

    if catalog.items.is_empty() {
        show_error(
            parent,
            language,
            "Nessuna audiodescrizione Rai disponibile nel catalogo.",
        );
        return;
    }

    crate::screen_reader_speak("Caricamento audiodescrizioni Rai");
    let (display_items, labels) = build_display_items(&catalog.items);
    let Some(selected_label) = interpreter_select_window::select_interpreter(
        parent,
        labels,
        language,
        "Rai audiodescrizioni".to_string(),
    ) else {
        return;
    };

    let Some((_, selected_item)) = display_items
        .into_iter()
        .find(|(label, _)| label == &selected_label)
    else {
        show_error(
            parent,
            language,
            "Impossibile aprire l'audiodescrizione selezionata.",
        );
        return;
    };

    let resolved_url = match rai_audiodescrizioni::resolve_audio_url(&selected_item.audio_url) {
        Ok(url) => url,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };

    let title = selected_item.title.trim().to_string();
    let title = if title.is_empty() { None } else { Some(title) };
    crate::play_named_remote_audio_from_url(parent, resolved_url, title, Some("audio/mpeg"));
}

fn build_display_items(items: &[CatalogItem]) -> (Vec<(String, CatalogItem)>, Vec<String>) {
    let mut used = HashSet::new();
    let mut display_items = Vec::with_capacity(items.len());
    let mut labels = Vec::with_capacity(items.len());

    for item in items {
        let base_label = format_item_label(item);
        let unique_label = ensure_unique_label(base_label, item, &mut used);
        labels.push(unique_label.clone());
        display_items.push((unique_label, item.clone()));
    }

    (display_items, labels)
}

fn format_item_label(item: &CatalogItem) -> String {
    let mut parts = Vec::new();
    let title = item.title.trim();
    if !title.is_empty() {
        parts.push(title.to_string());
    }
    let display_date = item.date.trim().to_string();
    if !display_date.is_empty() {
        parts.push(display_date);
    } else {
        let gen_date = item.gen_date.as_deref().unwrap_or("").trim();
        if !gen_date.is_empty() {
            parts.push(gen_date.to_string());
        }
    }
    let description = item.description.trim();
    if !description.is_empty() {
        parts.push(description.to_string());
    } else {
        let set_name = item.set_name.trim();
        if !set_name.is_empty() {
            parts.push(set_name.to_string());
        }
    }

    if parts.is_empty() {
        "Audiodescrizione Rai".to_string()
    } else {
        parts.join(" - ")
    }
}

fn ensure_unique_label(
    base_label: String,
    item: &CatalogItem,
    used: &mut HashSet<String>,
) -> String {
    if used.insert(base_label.clone()) {
        return base_label;
    }

    let date = item.date.trim();
    if !date.is_empty() {
        let dated = format!("{base_label} - {date}");
        if used.insert(dated.clone()) {
            return dated;
        }
    }

    let mut index = 2usize;
    loop {
        let candidate = format!("{base_label} ({index})");
        if used.insert(candidate.clone()) {
            return candidate;
        }
        index += 1;
    }
}
