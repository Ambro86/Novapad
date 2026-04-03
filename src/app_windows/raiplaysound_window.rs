use windows::Win32::Foundation::HWND;

use crate::app_windows::interpreter_select_window;
use crate::settings::Language;
use crate::tools::raiplaysound::{self, BrowseItem, BrowseItemKind, BrowsePage};
use crate::{RaiAudioOrigin, show_error, with_state};

enum BrowseOutcome {
    Cancelled,
    AudioStarted,
}

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }

    crate::screen_reader_speak("Caricamento generi RaiPlay Sound");
    let page = match raiplaysound::load_root_page() {
        Ok(page) => page,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };
    with_state(parent, |state| {
        state.raiplaysound_navigation_stack.clear();
        state.last_raiplaysound_page_path = None;
        state.last_raiplaysound_item_id = None;
    });
    let _outcome = browse_page(parent, language, page, None, Vec::new());
}

pub fn reopen_last(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }

    let (saved_stack, saved_page, saved_item_id) = with_state(parent, |state| {
        (
            state.raiplaysound_navigation_stack.clone(),
            state.last_raiplaysound_page_path.clone(),
            state.last_raiplaysound_item_id.clone(),
        )
    })
    .unwrap_or((Vec::new(), None, None));

    let page = match saved_page {
        Some(path) => raiplaysound::load_page(&path),
        None => raiplaysound::load_root_page(),
    };
    match page {
        Ok(page) => {
            crate::set_foreground_window_safe(parent);
            let _outcome = browse_page(parent, language, page, saved_item_id, saved_stack);
        }
        Err(err) => show_error(parent, language, &err),
    }
}

fn browse_page(
    parent: HWND,
    language: Language,
    mut page: BrowsePage,
    mut selected_id: Option<String>,
    mut history: Vec<(String, Option<String>)>,
) -> BrowseOutcome {
    loop {
        if page.items.is_empty() {
            show_error(
                parent,
                language,
                "Nessun contenuto disponibile in RaiPlay Sound.",
            );
            return BrowseOutcome::Cancelled;
        }

        let labels = page
            .items
            .iter()
            .map(|item| item.label.clone())
            .collect::<Vec<_>>();
        let initial_label = selected_id.as_deref().and_then(|id| {
            page.items
                .iter()
                .find(|item| item.id == id)
                .map(|item| item.label.clone())
        });

        let Some(selected_label) = interpreter_select_window::select_interpreter_with_context_actions_without_parent_restore_on_accept_but_restore_on_cancel(
            parent,
            labels,
            language,
            page.title.clone(),
            initial_label,
            Vec::new(),
        ) else {
            if let Some((previous_page_path, previous_selected_id)) = history.pop() {
                match raiplaysound::load_page(&previous_page_path) {
                    Ok(previous_page) => {
                        page = previous_page;
                        selected_id = previous_selected_id;
                        continue;
                    }
                    Err(err) => {
                        show_error(parent, language, &err);
                    }
                }
            }
            return BrowseOutcome::Cancelled;
        };

        let Some(selected_item) = page
            .items
            .iter()
            .find(|item| item.label == selected_label)
            .cloned()
        else {
            show_error(
                parent,
                language,
                "Impossibile aprire l'elemento selezionato.",
            );
            return BrowseOutcome::Cancelled;
        };

        selected_id = Some(selected_item.id.clone());
        match selected_item.kind {
            BrowseItemKind::Page => {
                let Some(path_id) = selected_item.path_id.as_deref() else {
                    show_error(
                        parent,
                        language,
                        "La pagina RaiPlay Sound selezionata non ha un percorso valido.",
                    );
                    continue;
                };
                crate::screen_reader_speak("Caricamento contenuto RaiPlay Sound");
                match raiplaysound::load_page(path_id) {
                    Ok(next_page) => {
                        history.push((page.source.clone(), selected_id.clone()));
                        page = next_page;
                        selected_id = None;
                        continue;
                    }
                    Err(err) => show_error(parent, language, &err),
                }
            }
            BrowseItemKind::Audio => {
                open_audio_item(parent, language, &page, &selected_item, &history);
                return BrowseOutcome::AudioStarted;
            }
        }
    }
}

fn open_audio_item(
    parent: HWND,
    language: Language,
    page: &BrowsePage,
    item: &BrowseItem,
    history: &[(String, Option<String>)],
) {
    let Some(audio_url) = item
        .audio_url
        .as_ref()
        .filter(|value| !value.trim().is_empty())
    else {
        show_error(
            parent,
            language,
            "Il contenuto selezionato non ha un URL audio disponibile.",
        );
        return;
    };

    let title = item.title.trim();
    let title = if title.is_empty() {
        None
    } else {
        Some(title.to_string())
    };
    with_state(parent, |state| {
        state.raiplaysound_navigation_stack = history.to_vec();
        state.last_raiplaysound_page_path = Some(page.source.clone());
        state.last_raiplaysound_item_id = Some(item.id.clone());
    });
    crate::play_named_remote_audio_from_url_with_rai_origin(
        parent,
        audio_url.clone(),
        title,
        Some("audio/mpeg"),
        RaiAudioOrigin::RaiPlaySound,
    );
}
