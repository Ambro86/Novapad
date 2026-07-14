use std::collections::HashSet;
use std::sync::Arc;
use windows::Win32::Foundation::HWND;

use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::Language;
use crate::tools::raiplay::{self, BrowseItem, BrowseItemKind, BrowsePage, PlaybackTarget};
use crate::{RaiAudioOrigin, show_error, with_state};

enum BrowseOutcome {
    Cancelled,
    MediaStarted,
}

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }

    crate::screen_reader_speak("Caricamento RaiPlay");
    let page = match raiplay::load_root_page() {
        Ok(page) => page,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };
    with_state(parent, |state| {
        state.raiplay_navigation_stack.clear();
        state.last_raiplay_page_path = None;
        state.last_raiplay_item_id = None;
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
            state.raiplay_navigation_stack.clone(),
            state.last_raiplay_page_path.clone(),
            state.last_raiplay_item_id.clone(),
        )
    })
    .unwrap_or((Vec::new(), None, None));

    let page = match saved_page {
        Some(path) => raiplay::load_page(&path),
        None => raiplay::load_root_page(),
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
    let mut auto_opened_single_pages = HashSet::new();
    loop {
        if page.items.is_empty() {
            if let Some((previous_page_path, previous_selected_id)) = history.pop() {
                match raiplay::load_page(&previous_page_path) {
                    Ok(previous_page) => {
                        page = previous_page;
                        selected_id = previous_selected_id;
                        current_search_query.clear();
                        continue;
                    }
                    Err(err) => show_error(parent, language, &err),
                }
            }
            return BrowseOutcome::Cancelled;
        }

        if let Some(single_item) = page
            .items
            .first()
            .filter(|_| page.items.len() == 1)
            .cloned()
            && single_item.kind == BrowseItemKind::Page
            && let Some(path_id) = single_item.path_id.as_deref()
        {
            if !auto_opened_single_pages.insert(page.source.clone()) {
                crate::log_debug(&format!(
                    "RaiPlay auto-open single page skipped due to loop: source={} target={}",
                    page.source, path_id
                ));
            } else {
                crate::screen_reader_speak("Caricamento contenuto RaiPlay");
                match raiplay::load_page(path_id) {
                    Ok(next_page) => {
                        page = next_page;
                        selected_id = None;
                        current_search_query.clear();
                        continue;
                    }
                    Err(err) => show_error(parent, language, &err),
                }
            }
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
        let context_items_for_enabled = page.items.clone();
        let context_items_for_handler = page.items.clone();
        let context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: format!(
                    "{} (Ctrl+C)",
                    crate::i18n::tr(language, "rai_audiodescrizioni.copy_audio_url")
                ),
                ctrl_c_shortcut: true,
                enabled: Arc::new(move |selected_id: &str| {
                    context_items_for_enabled
                        .iter()
                        .find(|item| item.id == selected_id)
                        .map(|item| {
                            item.kind == BrowseItemKind::Media
                                && item
                                    .media_url
                                    .as_ref()
                                    .map(|url| !url.trim().is_empty())
                                    .unwrap_or(false)
                        })
                        .unwrap_or(false)
                }),
                handler: Arc::new(move |selected_id: String| {
                    if let Some(item) = context_items_for_handler
                        .iter()
                        .find(|item| item.id == selected_id)
                        && let Some(media_url) = item.media_url.as_ref()
                    {
                        let clipboard_url = match raiplay::resolve_playback_target(media_url) {
                            Ok(PlaybackTarget::DirectStream { url, .. })
                            | Ok(PlaybackTarget::Download(url)) => url,
                            Err(err) => {
                                crate::log_debug(&format!(
                                    "RaiPlay copy URL fallback to original media URL: {}",
                                    err
                                ));
                                media_url.trim().to_string()
                            }
                        };
                        crate::app_windows::rai_audiodescrizioni_window::copy_text_to_clipboard(
                            parent,
                            &crate::app_windows::rai_audiodescrizioni_window::format_audio_url_clipboard_text(
                                language,
                                &item.title,
                                &clipboard_url,
                            ),
                        );
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
                context_actions: vec![context_action],
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: false,
                refresh: None,
            },
        );
        let selected_item_id = match selection {
            MultilineSelectionResult::Selected(id) => id,
            MultilineSelectionResult::Search(query) => {
                let trimmed_query = query.trim();
                if trimmed_query.is_empty() {
                    continue;
                }
                crate::screen_reader_speak("Ricerca RaiPlay in corso");
                match raiplay::search(trimmed_query) {
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
                    match raiplay::load_page(&previous_page_path) {
                        Ok(previous_page) => {
                            page = previous_page;
                            selected_id = previous_selected_id;
                            current_search_query.clear();
                            continue;
                        }
                        Err(err) => show_error(parent, language, &err),
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
                        "La pagina RaiPlay selezionata non ha un percorso valido.",
                    );
                    continue;
                };
                crate::screen_reader_speak("Caricamento contenuto RaiPlay");
                match raiplay::load_page(path_id) {
                    Ok(next_page) => {
                        history.push((page.source.clone(), selected_id.clone()));
                        page = next_page;
                        selected_id = None;
                        continue;
                    }
                    Err(err) => show_error(parent, language, &err),
                }
            }
            BrowseItemKind::Media => {
                if open_media_item(parent, language, &page, &selected_item, &history) {
                    return BrowseOutcome::MediaStarted;
                }
                continue;
            }
        }
    }
}

fn open_media_item(
    parent: HWND,
    language: Language,
    page: &BrowsePage,
    item: &BrowseItem,
    history: &[(String, Option<String>)],
) -> bool {
    let Some(media_url) = item
        .media_url
        .as_ref()
        .filter(|value| !value.trim().is_empty())
    else {
        show_error(
            parent,
            language,
            "Il contenuto selezionato non ha un URL media disponibile.",
        );
        return false;
    };

    let title = item.title.trim();
    let title = if title.is_empty() {
        None
    } else {
        Some(title.to_string())
    };
    let container_title =
        dedupe_raiplay_container_title(preferred_container_title(page, item), title.as_deref());
    let playback_target = match raiplay::resolve_playback_target(media_url) {
        Ok(target) => target,
        Err(err) => {
            if err == raiplay::DRM_NOT_SUPPORTED_ERROR {
                show_error(
                    parent,
                    language,
                    &crate::i18n::tr(language, "stream_audio.drm_not_supported"),
                );
            } else {
                show_error(parent, language, &err);
            }
            return false;
        }
    };
    with_state(parent, |state| {
        state.raiplay_navigation_stack = history.to_vec();
        state.last_raiplay_page_path = Some(page.source.clone());
        state.last_raiplay_item_id = Some(item.id.clone());
    });
    match playback_target {
        PlaybackTarget::Download(url) => {
            crate::log_debug(&format!(
                "RaiPlay playback: download target title={:?} container={:?} url={}",
                title,
                container_title.as_deref(),
                url
            ));
            crate::play_named_remote_audio_from_url_with_rai_origin(
                parent,
                url,
                title,
                container_title.as_deref(),
                RaiAudioOrigin::RaiPlay,
            );
        }
        PlaybackTarget::DirectStream {
            url,
            media_url,
            is_live,
            live_audio_tracks,
        } => {
            crate::log_debug(&format!(
                "RaiPlay playback: direct stream title={:?} container={:?} audio_url={} media_url={} is_live={} live_tracks={}",
                title,
                container_title.as_deref(),
                url,
                media_url,
                is_live,
                live_audio_tracks.len()
            ));
            if is_live {
                crate::play_live_stream_audio_from_url_with_rai_origin(
                    parent,
                    url,
                    container_title.clone(),
                    title,
                    live_audio_tracks,
                    RaiAudioOrigin::RaiPlay,
                );
            } else if let Err(err) = crate::launch_raiplay_in_mpv(
                parent,
                media_url.as_str(),
                container_title.as_deref(),
                title.as_deref(),
                RaiAudioOrigin::RaiPlay,
            ) {
                show_error(parent, language, &err);
                return false;
            }
            if with_state(parent, |state| {
                state.active_podcast_episode_media_url = Some(media_url);
            })
            .is_none()
            {
                crate::log_debug("Failed to persist RaiPlay media URL for save");
            }
        }
    };
    true
}

fn dedupe_raiplay_container_title(
    container_title: Option<String>,
    item_title: Option<&str>,
) -> Option<String> {
    let container_title = container_title?;
    let Some(item_title) = item_title.map(str::trim).filter(|title| !title.is_empty()) else {
        return Some(container_title);
    };
    if container_title.trim().eq_ignore_ascii_case(item_title) {
        None
    } else {
        Some(container_title)
    }
}

fn preferred_container_title(page: &BrowsePage, item: &BrowseItem) -> Option<String> {
    let page_title = page.title.trim();
    if page_title.is_empty() {
        return None;
    }
    if !page_title.eq_ignore_ascii_case("Episodi") {
        return Some(page_title.to_string());
    }
    if let Some(program_title) = item
        .program_title
        .as_deref()
        .map(str::trim)
        .filter(|title| !title.is_empty() && !title.eq_ignore_ascii_case("Episodi"))
    {
        return Some(program_title.to_string());
    }
    let description = item.description.as_deref()?.trim();
    let inferred = description.rsplit_once(" - ")?.1.trim();
    if inferred.is_empty() || inferred.eq_ignore_ascii_case("Episodi") {
        None
    } else {
        Some(inferred.to_string())
    }
}
