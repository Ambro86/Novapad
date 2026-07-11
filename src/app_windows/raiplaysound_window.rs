use std::sync::Arc;
use windows::Win32::Foundation::HWND;

use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
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
    let mut current_search_query = String::new();
    loop {
        if page.items.is_empty() {
            show_error(
                parent,
                language,
                "Nessun contenuto disponibile in RaiPlay Sound.",
            );
            return BrowseOutcome::Cancelled;
        }

        let selection_items = page
            .items
            .iter()
            .map(|item| MultilineSelectionItem {
                id: item.id.clone(),
                title: item.title.clone(),
                description: item.description.clone(),
            })
            .collect::<Vec<_>>();
        let context_items_for_copy_enabled = page.items.clone();
        let context_items_for_copy_handler = page.items.clone();
        let copy_context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: format!(
                    "{} (Ctrl+C)",
                    crate::i18n::tr(language, "rai_audiodescrizioni.copy_audio_url")
                ),
                ctrl_c_shortcut: true,
                enabled: Arc::new(move |selected_id: &str| {
                    context_items_for_copy_enabled
                        .iter()
                        .find(|item| item.id == selected_id)
                        .map(|item| {
                            item.kind == BrowseItemKind::Audio
                                && item
                                    .audio_url
                                    .as_ref()
                                    .map(|url| !url.trim().is_empty())
                                    .unwrap_or(false)
                        })
                        .unwrap_or(false)
                }),
                handler: Arc::new(move |selected_id: String| {
                    if let Some(item) = context_items_for_copy_handler
                        .iter()
                        .find(|item| item.id == selected_id)
                        && let Some(audio_url) = item.audio_url.as_ref()
                    {
                        crate::app_windows::rai_audiodescrizioni_window::copy_text_to_clipboard(
                            parent,
                            &crate::app_windows::rai_audiodescrizioni_window::format_resolved_audio_url_clipboard_text(
                                language, &item.title, audio_url,
                            ),
                        );
                    }
                }),
            };
        let context_items_for_podcast_enabled = page.items.clone();
        let container_url = page.source.clone();
        let container_url_for_enabled = container_url.clone();
        let container_title = page.title.clone();
        let add_to_podcasts_context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: "Aggiungi ai podcast".to_string(),
                ctrl_c_shortcut: false,
                enabled: Arc::new(move |selected_id: &str| {
                    if !raiplaysound::is_raiplaysound_url(&container_url_for_enabled) {
                        return false;
                    }
                    context_items_for_podcast_enabled
                        .iter()
                        .find(|item| item.id == selected_id)
                        .map(|item| {
                            item.kind == BrowseItemKind::Audio
                                && item
                                    .audio_url
                                    .as_ref()
                                    .map(|url| !url.trim().is_empty())
                                    .unwrap_or(false)
                        })
                        .unwrap_or(false)
                }),
                handler: Arc::new(move |selected_id: String| {
                    if selected_id.trim().is_empty() {
                        return;
                    }
                    if crate::app_windows::podcasts_window::add_podcast_source(
                        parent,
                        &container_url,
                        &container_title,
                    )
                    .is_some()
                    {
                        crate::screen_reader_speak(&crate::i18n::tr(language, "podcasts.added"));
                        crate::app_windows::podcasts_window::show_subscription_confirmation(
                            parent, language,
                        );
                    } else {
                        crate::screen_reader_speak("Podcast gia presente");
                    }
                }),
            };

        let selection = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            page.title.clone(),
            selection_items,
            selected_id.clone(),
            MultilineSearchOptions {
                initial_query: current_search_query.clone(),
                search_button_label: "Cerca".to_string(),
                show_search_edit: true,
                secondary_action_label: None,
                context_actions: vec![copy_context_action, add_to_podcasts_context_action],
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: false,
            },
        );
        let selected_item_id = match selection {
            MultilineSelectionResult::Selected(id) => id,
            MultilineSelectionResult::Search(query) => {
                let trimmed_query = query.trim();
                if trimmed_query.is_empty() {
                    continue;
                }
                crate::screen_reader_speak("Ricerca RaiPlay Sound in corso");
                match raiplaysound::search(trimmed_query) {
                    Ok(search_page) => {
                        history.push((page.source.clone(), selected_id.clone()));
                        page = search_page;
                        selected_id = None;
                        current_search_query = trimmed_query.to_string();
                        continue;
                    }
                    Err(err) => {
                        show_error(parent, language, &err);
                        continue;
                    }
                }
            }
            MultilineSelectionResult::SecondaryAction => {
                continue;
            }
            MultilineSelectionResult::Cancelled => {
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
            }
        };

        let Some(selected_item) = page
            .items
            .iter()
            .find(|item| item.id == selected_item_id)
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
        current_search_query.clear();
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
