use std::collections::HashSet;

use windows::Win32::Foundation::HWND;

use crate::app_windows::interpreter_select_window;
use crate::app_windows::interpreter_select_window::{
    GroupedSelectGroup, GroupedSelectItem, InterpreterSelectionResult,
};
use crate::settings::Language;
use crate::tools::rai_audiodescrizioni::{self, CatalogGroup, CatalogItem};
use crate::{RaiAudioOrigin, show_error, with_state};

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
    let selection = interpreter_select_window::select_interpreter_with_secondary_action(
        parent,
        labels,
        language,
        "Rai audiodescrizioni".to_string(),
        "Mostra tutte le audiodescrizioni".to_string(),
    );

    match selection {
        Some(InterpreterSelectionResult::Item(selected_label)) => {
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
            open_item(parent, language, &selected_item, RaiAudioOrigin::Recenti);
        }
        Some(InterpreterSelectionResult::SecondaryAction) => {
            open_grouped(parent);
        }
        None => {}
    }
}

pub fn open_grouped(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    let initial_item_id =
        with_state(parent, |state| state.last_rai_grouped_item_id.clone()).unwrap_or(None);
    open_grouped_catalog(parent, language, initial_item_id);
}

fn open_grouped_catalog(parent: HWND, language: Language, initial_item_id: Option<String>) {
    crate::screen_reader_speak("Caricamento catalogo completo audiodescrizioni Rai");
    let groups = match rai_audiodescrizioni::load_grouped_catalog() {
        Ok(groups) => groups,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };

    if groups.is_empty() {
        show_error(
            parent,
            language,
            "Nessuna audiodescrizione Rai disponibile nel catalogo completo.",
        );
        return;
    }

    let grouped_items = build_grouped_items(&groups);
    let Some(selected_value) = interpreter_select_window::select_grouped_interpreter(
        parent,
        grouped_items,
        language,
        "Tutte le audiodescrizioni Rai".to_string(),
        initial_item_id,
    ) else {
        return;
    };

    for group in groups {
        for item in group.items {
            if item.item_id == selected_value {
                with_state(parent, |state| {
                    state.last_rai_grouped_item_id = Some(item.item_id.clone());
                });
                open_item(parent, language, &item, RaiAudioOrigin::Tutte);
                return;
            }
        }
    }

    show_error(
        parent,
        language,
        "Impossibile aprire l'audiodescrizione selezionata.",
    );
}

fn open_item(
    parent: HWND,
    language: Language,
    selected_item: &CatalogItem,
    rai_origin: RaiAudioOrigin,
) {
    let resolved_url = match rai_audiodescrizioni::resolve_audio_url(&selected_item.audio_url) {
        Ok(url) => url,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };

    let title = selected_item.title.trim().to_string();
    let title = if title.is_empty() { None } else { Some(title) };
    crate::play_named_remote_audio_from_url_with_rai_origin(
        parent,
        resolved_url,
        title,
        Some("audio/mpeg"),
        rai_origin,
    );
}

fn build_grouped_items(groups: &[CatalogGroup]) -> Vec<GroupedSelectGroup> {
    groups
        .iter()
        .map(|group| GroupedSelectGroup {
            label: group.title.clone(),
            items: group
                .items
                .iter()
                .map(|item| GroupedSelectItem {
                    label: item.title.clone(),
                    value: item.item_id.clone(),
                })
                .collect(),
        })
        .collect()
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
