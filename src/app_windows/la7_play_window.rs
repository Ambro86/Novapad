use std::collections::HashMap;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicIsize, Ordering},
};

use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::Language;
use crate::tools::la7_play::{self, BrowseItem, BrowsePage, ItemKind};
use crate::{RaiAudioOrigin, show_error, with_state};
use windows::Win32::Foundation::HWND;

type La7ContextTargetCache = Arc<Mutex<HashMap<String, Result<String, String>>>>;

static LA7_CONTEXT_TRANSCRIPTION_PENDING_PARENT: AtomicIsize = AtomicIsize::new(0);
static LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT: AtomicIsize = AtomicIsize::new(0);
static LA7_CONTEXT_TRANSCRIPTION_FOCUS_PARENT: AtomicIsize = AtomicIsize::new(0);
static LA7_CONTEXT_AUDIO_DESCRIPTION_PENDING_PARENT: AtomicIsize = AtomicIsize::new(0);
static LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT: AtomicIsize = AtomicIsize::new(0);
static LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT: AtomicIsize = AtomicIsize::new(0);

pub(crate) fn mark_context_transcription_started(parent: HWND) {
    LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT.store(0, Ordering::SeqCst);
    LA7_CONTEXT_TRANSCRIPTION_FOCUS_PARENT.store(parent.0, Ordering::SeqCst);
    LA7_CONTEXT_TRANSCRIPTION_PENDING_PARENT.store(parent.0, Ordering::SeqCst);
    crate::log_debug(&format!(
        "La7 Play context transcription started parent={:?}",
        parent
    ));
}

pub(crate) fn finish_context_transcription(parent: HWND, succeeded: bool) {
    if LA7_CONTEXT_TRANSCRIPTION_PENDING_PARENT
        .compare_exchange(parent.0, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_err()
    {
        return;
    }
    LA7_CONTEXT_TRANSCRIPTION_FOCUS_PARENT.store(0, Ordering::SeqCst);
    if succeeded {
        LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT.store(parent.0, Ordering::SeqCst);
        crate::log_debug(&format!(
            "La7 Play context transcription completed: browser exit armed parent={:?}",
            parent
        ));
    } else {
        LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT.store(0, Ordering::SeqCst);
        crate::log_debug(&format!(
            "La7 Play context transcription did not complete parent={:?}",
            parent
        ));
    }
}

pub(crate) fn context_transcription_keeps_progress_foreground(parent: HWND) -> bool {
    LA7_CONTEXT_TRANSCRIPTION_FOCUS_PARENT.load(Ordering::SeqCst) == parent.0
}

fn take_context_transcription_browser_exit(parent: HWND) -> bool {
    if LA7_CONTEXT_TRANSCRIPTION_EXIT_PARENT
        .compare_exchange(parent.0, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_ok()
    {
        crate::log_debug(&format!(
            "La7 Play context transcription: closing navigation instead of restoring history parent={:?}",
            parent
        ));
        true
    } else {
        false
    }
}

pub(crate) fn mark_context_audio_description_started(parent: HWND) {
    LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT.store(0, Ordering::SeqCst);
    LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT.store(0, Ordering::SeqCst);
    LA7_CONTEXT_AUDIO_DESCRIPTION_PENDING_PARENT.store(parent.0, Ordering::SeqCst);
    crate::log_debug(&format!(
        "La7 Play context audio description started parent={:?}",
        parent
    ));
}

pub(crate) fn finish_context_audio_description_download(parent: HWND, succeeded: bool) -> bool {
    if LA7_CONTEXT_AUDIO_DESCRIPTION_PENDING_PARENT
        .compare_exchange(parent.0, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_err()
    {
        return false;
    }
    if succeeded {
        LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT.store(parent.0, Ordering::SeqCst);
        LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT.store(parent.0, Ordering::SeqCst);
        crate::log_debug(&format!(
            "La7 Play context audio description completed: browser exit and focus protection armed parent={:?}",
            parent
        ));
    } else {
        LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT.store(0, Ordering::SeqCst);
        LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT.store(0, Ordering::SeqCst);
        crate::log_debug(&format!(
            "La7 Play context audio description did not complete: keeping browser open parent={:?}",
            parent
        ));
    }
    true
}

pub(crate) fn context_audio_description_keeps_window_foreground(parent: HWND) -> bool {
    LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT.load(Ordering::SeqCst) == parent.0
}

pub(crate) fn finish_context_audio_description_focus(parent: HWND) {
    if LA7_CONTEXT_AUDIO_DESCRIPTION_FOCUS_PARENT
        .compare_exchange(parent.0, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_ok()
    {
        crate::log_debug(&format!(
            "La7 Play context audio description: focus protection cleared parent={:?}",
            parent
        ));
    }
}

fn take_context_audio_description_browser_exit(parent: HWND) -> bool {
    if LA7_CONTEXT_AUDIO_DESCRIPTION_EXIT_PARENT
        .compare_exchange(parent.0, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_ok()
    {
        crate::log_debug(&format!(
            "La7 Play context audio description: closing navigation instead of restoring history parent={:?}",
            parent
        ));
        true
    } else {
        false
    }
}

fn context_menu_label(language: Language, key: &str) -> String {
    crate::i18n::tr(language, key)
        .replace('&', "")
        .split('\t')
        .next()
        .unwrap_or_default()
        .trim()
        .to_string()
}

fn optional_media_title(title: &str) -> Option<String> {
    let title = title.trim();
    if title.is_empty() {
        None
    } else {
        Some(title.to_string())
    }
}

fn cached_context_vod_url(
    items: &[BrowseItem],
    selected_id: &str,
    cache: &La7ContextTargetCache,
) -> Option<String> {
    if let Some(cached) = cache
        .lock()
        .unwrap_or_else(|err| err.into_inner())
        .get(selected_id)
        .cloned()
    {
        return cached.ok();
    }
    let item = items.iter().find(|item| item.id == selected_id)?;
    if item.kind != ItemKind::Media {
        return None;
    }
    let resolved = la7_play::resolve_vod(&item.target);
    cache
        .lock()
        .unwrap_or_else(|err| err.into_inner())
        .insert(selected_id.to_string(), resolved.clone());
    resolved.ok()
}
pub fn open(parent: HWND) {
    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }
    crate::screen_reader_speak(&crate::i18n::tr_la7_play("la7.loading"));
    with_state(parent, |s| {
        s.la7_play_navigation_stack.clear();
        s.last_la7_play_page_path = None;
        s.last_la7_play_item_id = None;
        s.last_la7_play_search_query.clear();
    });
    browse(
        parent,
        language,
        la7_play::root_page(),
        None,
        Vec::new(),
        String::new(),
    );
}

pub fn reopen_last(parent: HWND) {
    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    let (stack, path, item, query) = with_state(parent, |s| {
        (
            s.la7_play_navigation_stack.clone(),
            s.last_la7_play_page_path.clone(),
            s.last_la7_play_item_id.clone(),
            s.last_la7_play_search_query.clone(),
        )
    })
    .unwrap_or_default();
    let page = path
        .as_deref()
        .map(la7_play::load_page)
        .unwrap_or_else(|| Ok(la7_play::root_page()));
    match page {
        Ok(p) => {
            crate::set_foreground_window_safe(parent);
            browse(parent, language, p, item, stack, query)
        }
        Err(e) => show_error(parent, language, &e),
    }
}

fn browse(
    parent: HWND,
    language: Language,
    mut page: BrowsePage,
    mut selected: Option<String>,
    mut history: Vec<(String, Option<String>)>,
    mut query: String,
) {
    loop {
        if page.items.is_empty() {
            show_error(
                parent,
                language,
                &crate::i18n::tr_la7_play("la7.no_results"),
            );
            if let Some((src, id)) = history.pop() {
                match la7_play::load_page(&src) {
                    Ok(p) => {
                        page = p;
                        selected = id;
                        query.clear();
                        continue;
                    }
                    Err(e) => show_error(parent, language, &e),
                }
            }
            return;
        }
        let rows = page
            .items
            .iter()
            .map(|x| MultilineSelectionItem {
                id: x.id.clone(),
                title: x.title.clone(),
                description: x.description.clone(),
            })
            .collect();
        let context_target_cache: La7ContextTargetCache = Arc::new(Mutex::new(HashMap::new()));

        let save_items_for_enabled = page.items.clone();
        let save_cache_for_enabled = Arc::clone(&context_target_cache);
        let save_items_for_handler = page.items.clone();
        let save_page_title = page.title.clone();
        let save_cache_for_handler = Arc::clone(&context_target_cache);
        let save_context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: context_menu_label(language, "playback.download_episode"),
                ctrl_c_shortcut: false,
                delete_shortcut: false,
                children: Vec::new(),
                enabled: Arc::new(move |selected_id: &str| {
                    cached_context_vod_url(
                        &save_items_for_enabled,
                        selected_id,
                        &save_cache_for_enabled,
                    )
                    .is_some()
                }),
                handler: Arc::new(move |selected_id: String| {
                    let Some(item) = save_items_for_handler
                        .iter()
                        .find(|item| item.id == selected_id)
                    else {
                        return;
                    };
                    let Some(url) = cached_context_vod_url(
                        &save_items_for_handler,
                        &selected_id,
                        &save_cache_for_handler,
                    ) else {
                        return;
                    };
                    crate::save_la7_context_media(
                        parent,
                        language,
                        url,
                        optional_media_title(&save_page_title),
                        optional_media_title(&item.title),
                    );
                }),
            };

        let transcribe_items_for_enabled = page.items.clone();
        let transcribe_cache_for_enabled = Arc::clone(&context_target_cache);
        let transcribe_items_for_handler = page.items.clone();
        let transcribe_cache_for_handler = Arc::clone(&context_target_cache);
        let transcribe_context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: context_menu_label(language, "playback.transcribe_current"),
                ctrl_c_shortcut: false,
                delete_shortcut: false,
                children: Vec::new(),
                enabled: Arc::new(move |selected_id: &str| {
                    cached_context_vod_url(
                        &transcribe_items_for_enabled,
                        selected_id,
                        &transcribe_cache_for_enabled,
                    )
                    .is_some()
                }),
                handler: Arc::new(move |selected_id: String| {
                    let Some(item) = transcribe_items_for_handler
                        .iter()
                        .find(|item| item.id == selected_id)
                    else {
                        return;
                    };
                    let Some(url) = cached_context_vod_url(
                        &transcribe_items_for_handler,
                        &selected_id,
                        &transcribe_cache_for_handler,
                    ) else {
                        return;
                    };
                    crate::start_whisper_transcription_for_la7_context(
                        parent,
                        url,
                        optional_media_title(&item.title),
                    );
                }),
            };

        let ad_items_for_enabled = page.items.clone();
        let ad_cache_for_enabled = Arc::clone(&context_target_cache);
        let ad_items_for_handler = page.items.clone();
        let ad_page_title = page.title.clone();
        let ad_cache_for_handler = Arc::clone(&context_target_cache);
        let audio_description_context_action =
            crate::app_windows::interpreter_select_window::InterpreterContextAction {
                label: context_menu_label(language, "menu.create_audio_description"),
                ctrl_c_shortcut: false,
                delete_shortcut: false,
                children: Vec::new(),
                enabled: Arc::new(move |selected_id: &str| {
                    cached_context_vod_url(
                        &ad_items_for_enabled,
                        selected_id,
                        &ad_cache_for_enabled,
                    )
                    .is_some()
                }),
                handler: Arc::new(move |selected_id: String| {
                    let Some(item) = ad_items_for_handler
                        .iter()
                        .find(|item| item.id == selected_id)
                    else {
                        return;
                    };
                    let Some(url) = cached_context_vod_url(
                        &ad_items_for_handler,
                        &selected_id,
                        &ad_cache_for_handler,
                    ) else {
                        return;
                    };
                    crate::create_audio_description_from_la7_context(
                        parent,
                        language,
                        url,
                        optional_media_title(&ad_page_title),
                        optional_media_title(&item.title),
                    );
                }),
            };

        let result = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            page.title.clone(),
            rows,
            selected.clone(),
            MultilineSearchOptions {
                initial_query: query.clone(),
                search_button_label: crate::i18n::tr_la7_play("la7.search"),
                show_search_edit: true,
                secondary_action_label: None,
                context_actions: vec![
                    save_context_action,
                    transcribe_context_action,
                    audio_description_context_action,
                ],
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: false,
                refresh: None,
            },
        );
        let id = match result {
            MultilineSelectionResult::Selected(id) => id,
            MultilineSelectionResult::Search(q) => {
                let q = q.trim();
                if q.is_empty() {
                    continue;
                }
                crate::screen_reader_speak(&crate::i18n::tr_la7_play("la7.searching"));
                match la7_play::search(q) {
                    Ok(p) => {
                        history.push((page.source.clone(), selected.clone()));
                        page = p;
                        selected = None;
                        query = q.into();
                        continue;
                    }
                    Err(e) => {
                        show_error(parent, language, &e);
                        continue;
                    }
                }
            }
            MultilineSelectionResult::SecondaryAction => continue,
            MultilineSelectionResult::Cancelled => {
                if take_context_audio_description_browser_exit(parent)
                    || take_context_transcription_browser_exit(parent)
                {
                    return;
                }
                if let Some((src, id)) = history.pop() {
                    match la7_play::load_page(&src) {
                        Ok(p) => {
                            page = p;
                            selected = id;
                            query.clear();
                            continue;
                        }
                        Err(e) => show_error(parent, language, &e),
                    }
                }
                return;
            }
        };
        let Some(item) = page.items.iter().find(|x| x.id == id).cloned() else {
            continue;
        };
        selected = Some(item.id.clone());
        match item.kind {
            ItemKind::Page => match la7_play::load_page(&item.target) {
                Ok(p) => {
                    history.push((page.source.clone(), selected.clone()));
                    page = p;
                    selected = None;
                    query.clear()
                }
                Err(e) => show_error(parent, language, &e),
            },
            ItemKind::Media => {
                if play_vod(parent, language, &page, &item, &history, &query) {
                    return;
                }
            }
            ItemKind::Live => {
                if play_live(parent, language, &page, &item, &history, &query) {
                    return;
                }
            }
        }
    }
}

fn save_return(
    parent: HWND,
    page: &BrowsePage,
    item: &BrowseItem,
    history: &[(String, Option<String>)],
    query: &str,
) {
    with_state(parent, |s| {
        s.la7_play_navigation_stack = history.to_vec();
        s.last_la7_play_page_path = Some(page.source.clone());
        s.last_la7_play_item_id = Some(item.id.clone());
        s.last_la7_play_search_query = query.to_string();
    });
}
fn play_vod(
    parent: HWND,
    language: Language,
    page: &BrowsePage,
    item: &BrowseItem,
    history: &[(String, Option<String>)],
    query: &str,
) -> bool {
    let url = match la7_play::resolve_vod(&item.target) {
        Ok(x) => x,
        Err(e) => {
            show_error(parent, language, &e);
            return false;
        }
    };
    save_return(parent, page, item, history, query);
    match crate::launch_raiplay_in_mpv(
        parent,
        &url,
        Some(&page.title),
        Some(&item.title),
        RaiAudioOrigin::La7Play,
    ) {
        Ok(()) => true,
        Err(e) => {
            show_error(
                parent,
                language,
                &crate::i18n::tr_la7_play_f("la7.open_error", &[("error", &e)]),
            );
            false
        }
    }
}
fn play_live(
    parent: HWND,
    language: Language,
    page: &BrowsePage,
    item: &BrowseItem,
    history: &[(String, Option<String>)],
    query: &str,
) -> bool {
    let loaded = match crate::tools::tv::load_channels_with_cache() {
        Ok(x) => x,
        Err(e) => {
            show_error(parent, language, &e);
            return false;
        }
    };
    let needle = item.target.to_lowercase();
    let channel = loaded.channels.into_iter().find(|c| {
        let values = [c.name.as_str(), c.tvg_name.as_str(), c.tvg_id.as_str()];
        values.iter().any(|v| {
            let l = v.to_lowercase();
            if needle == "la7" {
                l == "la7" || l.contains("la 7") && !l.contains("cinema")
            } else {
                l.contains("la7") && l.contains("cinema")
                    || l.contains("la 7") && l.contains("cinema")
                    || l.contains("la7d")
            }
        })
    });
    let Some(channel) = channel else {
        show_error(parent, language, &crate::i18n::tr_la7_play("la7.no_media"));
        return false;
    };
    let url = match crate::tools::tv::resolve_stream_url(&channel) {
        Ok(x) => x,
        Err(e) => {
            show_error(parent, language, &e);
            return false;
        }
    };
    save_return(parent, page, item, history, query);
    // Le dirette usano lo stesso percorso della TV, che accetta sia HLS sia DASH
    // e applica l'eventuale User-Agent specifico del canale. In particolare La7
    // Cinema non deve passare dal risolutore RaiPlay, che considera i manifest
    // .mpd come contenuti DRM.
    match crate::launch_tv_stream_in_mpv(parent, &url, &item.title, &channel.http_user_agent, false)
    {
        Ok(()) => {
            with_state(parent, |s| {
                s.active_podcast_episode_from_rai = RaiAudioOrigin::La7Play
            });
            true
        }
        Err(e) => {
            show_error(
                parent,
                language,
                &crate::i18n::tr_la7_play_f("la7.open_error", &[("error", &e)]),
            );
            false
        }
    }
}
