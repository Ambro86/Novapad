use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::Language;
use crate::tools::la7_play::{self, BrowseItem, BrowsePage, ItemKind};
use crate::{RaiAudioOrigin, show_error, with_state};
use windows::Win32::Foundation::HWND;

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
                context_actions: Vec::new(),
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
