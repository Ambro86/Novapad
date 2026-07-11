use std::collections::HashMap;
use std::sync::{Arc, Mutex, OnceLock};

use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::WM_CLOSE;

use crate::app_windows::interpreter_select_window::InterpreterContextAction;
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::{Language, TvFavorite, load_settings, save_settings};
use crate::stream_recording::{self, StreamRecordingKind};
use crate::tools::tv::{self, TvChannel, TvProgram};
use crate::{show_error, with_state};

#[derive(Clone, Debug, PartialEq, Eq)]
enum TvPage {
    Root,
    Category(String),
    Regions,
    Region(String),
    Favorites,
    Search(String),
}

#[derive(Clone, Debug)]
enum TvEntryKind {
    Page(TvPage),
    Channel(usize),
    Favorite(Box<TvFavorite>),
    Recordings,
}

#[derive(Clone, Debug)]
enum TvDeferredAction {
    Refresh,
    Record {
        channel: Box<TvChannel>,
        selected_id: String,
    },
}

static TV_DEFERRED_ACTION: OnceLock<Mutex<Option<TvDeferredAction>>> = OnceLock::new();

fn set_tv_deferred_action(action: TvDeferredAction) {
    *TV_DEFERRED_ACTION
        .get_or_init(|| Mutex::new(None))
        .lock()
        .unwrap_or_else(|err| err.into_inner()) = Some(action);
}

fn take_tv_deferred_action() -> Option<TvDeferredAction> {
    TV_DEFERRED_ACTION
        .get_or_init(|| Mutex::new(None))
        .lock()
        .unwrap_or_else(|err| err.into_inner())
        .take()
}

fn close_tv_browser_dialog() {
    let dialog = crate::get_foreground_window_safe();
    if dialog.0 != 0
        && let Err(err) = crate::post_message_w_safe(dialog, WM_CLOSE, WPARAM(0), LPARAM(0))
    {
        crate::log_debug(&format!("TV: failed to close channel browser: {err}"));
    }
}

#[derive(Clone, Debug)]
struct TvEntry {
    id: String,
    title: String,
    description: Option<String>,
    kind: TvEntryKind,
}

#[derive(Clone, Debug)]
struct TvCatalog {
    channels: Vec<TvChannel>,
    categories: Vec<(String, Vec<usize>)>,
    regions: Vec<(String, Vec<usize>)>,
    current_programs: HashMap<String, TvProgram>,
}

impl TvCatalog {
    fn new(channels: Vec<TvChannel>, current_programs: HashMap<String, TvProgram>) -> Self {
        let mut categories = Vec::<(String, Vec<usize>)>::new();
        let mut regions = Vec::<(String, Vec<usize>)>::new();

        for (index, channel) in channels.iter().enumerate() {
            if channel.is_regional() {
                let region = channel.regional_name().unwrap_or("Altre regioni");
                push_group_index(&mut regions, region, index);
            } else {
                push_group_index(&mut categories, &channel.category, index);
            }
        }

        Self {
            channels,
            categories,
            regions,
            current_programs,
        }
    }

    fn entries_for_page(&self, page: &TvPage) -> Vec<TvEntry> {
        match page {
            TvPage::Root => self.root_entries(),
            TvPage::Category(category) => self.channel_entries(
                self.categories
                    .iter()
                    .find(|(name, _)| name == category)
                    .map(|(_, indices)| indices.as_slice())
                    .unwrap_or_default(),
            ),
            TvPage::Regions => self.region_entries(),
            TvPage::Region(region) => self.channel_entries(
                self.regions
                    .iter()
                    .find(|(name, _)| name == region)
                    .map(|(_, indices)| indices.as_slice())
                    .unwrap_or_default(),
            ),
            TvPage::Favorites => self.favorite_entries(),
            TvPage::Search(query) => {
                let indices = self
                    .channels
                    .iter()
                    .enumerate()
                    .filter_map(|(index, channel)| {
                        tv::matches_search(channel, query).then_some(index)
                    })
                    .collect::<Vec<_>>();
                self.channel_entries(&indices)
            }
        }
    }

    fn root_entries(&self) -> Vec<TvEntry> {
        let mut entries = self
            .categories
            .iter()
            .enumerate()
            .map(|(index, (name, channels))| TvEntry {
                id: format!("category:{index}"),
                title: name.clone(),
                description: Some(channel_count_description(channels.len())),
                kind: TvEntryKind::Page(TvPage::Category(name.clone())),
            })
            .collect::<Vec<_>>();

        if !self.regions.is_empty() {
            let channel_count = self
                .regions
                .iter()
                .map(|(_, channels)| channels.len())
                .sum::<usize>();
            entries.push(TvEntry {
                id: "regions".to_string(),
                title: "Regionali".to_string(),
                description: Some(format!(
                    "{} regioni, {}",
                    self.regions.len(),
                    channel_count_description(channel_count)
                )),
                kind: TvEntryKind::Page(TvPage::Regions),
            });
        }
        entries.push(TvEntry {
            id: "recordings".to_string(),
            title: "Registrazioni TV".to_string(),
            description: Some("Apri le registrazioni dei canali TV".to_string()),
            kind: TvEntryKind::Recordings,
        });
        entries
    }

    fn region_entries(&self) -> Vec<TvEntry> {
        self.regions
            .iter()
            .enumerate()
            .map(|(index, (name, channels))| TvEntry {
                id: format!("region:{index}"),
                title: name.clone(),
                description: Some(channel_count_description(channels.len())),
                kind: TvEntryKind::Page(TvPage::Region(name.clone())),
            })
            .collect()
    }

    fn channel_entries(&self, indices: &[usize]) -> Vec<TvEntry> {
        indices
            .iter()
            .filter_map(|index| {
                let channel = self.channels.get(*index)?;
                let description = if channel.tvg_name.trim().is_empty()
                    || channel.tvg_name.trim().eq_ignore_ascii_case(&channel.name)
                {
                    Some(channel.category.clone())
                } else {
                    Some(format!("{}; {}", channel.tvg_name.trim(), channel.category))
                };
                Some(TvEntry {
                    id: format!("channel:{index}"),
                    title: self.channel_accessible_title(channel),
                    description,
                    kind: TvEntryKind::Channel(*index),
                })
            })
            .collect()
    }

    fn favorite_entries(&self) -> Vec<TvEntry> {
        load_settings()
            .tv_favorites
            .into_iter()
            .enumerate()
            .map(|(index, favorite)| {
                let channel = channel_from_favorite(&favorite);
                TvEntry {
                    id: format!("favorite:{index}"),
                    title: self.channel_accessible_title(&channel),
                    description: Some(favorite.category.clone()),
                    kind: TvEntryKind::Favorite(Box::new(favorite)),
                }
            })
            .collect()
    }

    fn channel_accessible_title(&self, channel: &TvChannel) -> String {
        tv::current_program_for_channel(&self.current_programs, channel)
            .map(|program| format!("{}. Ora in onda: {}", channel.name, program.title))
            .unwrap_or_else(|| channel.name.clone())
    }
}

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access_with_title(
        parent,
        language,
        Some("TV"),
    ) {
        return;
    }

    crate::screen_reader_speak("Caricamento canali TV");
    let load_result = match tv::load_channels_with_cache() {
        Ok(result) => result,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };
    if load_result.channels.is_empty() {
        show_error(parent, language, "Nessun canale TV disponibile.");
        return;
    }
    if let Some(warning) = load_result.cache_warning.as_deref() {
        crate::screen_reader_speak(warning);
    }

    let current_programs = match tv::load_current_programs() {
        Ok(programs) => programs,
        Err(err) => {
            crate::log_debug(&format!(
                "TV: programmi correnti non disponibili, mostro comunque i canali: {err}"
            ));
            HashMap::new()
        }
    };

    let catalog = Arc::new(TvCatalog::new(load_result.channels, current_programs));
    browse_catalog(parent, language, catalog);
}

fn browse_catalog(parent: HWND, language: Language, catalog: Arc<TvCatalog>) {
    let mut page = TvPage::Root;
    let mut selected_id = None;
    let mut history = Vec::<(TvPage, Option<String>)>::new();
    let mut search_query = String::new();

    loop {
        let entries = catalog.entries_for_page(&page);
        if entries.is_empty() {
            match &page {
                TvPage::Search(query) => show_error(
                    parent,
                    language,
                    &format!("Nessun canale TV trovato per {query}."),
                ),
                _ => show_error(parent, language, "Questa sezione TV è vuota."),
            }
            if let Some((previous_page, previous_selected_id)) = history.pop() {
                page = previous_page;
                selected_id = previous_selected_id;
                search_query.clear();
                continue;
            }
            return;
        }

        let selection_items = entries
            .iter()
            .map(|entry| MultilineSelectionItem {
                id: entry.id.clone(),
                title: entry.title.clone(),
                description: entry.description.clone(),
            })
            .collect::<Vec<_>>();

        let selection = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            page_title(&page),
            selection_items,
            selected_id.clone(),
            MultilineSearchOptions {
                initial_query: search_query.clone(),
                search_button_label: "Cerca".to_string(),
                show_search_edit: true,
                secondary_action_label: Some(match page {
                    TvPage::Favorites => "Registrazioni TV".to_string(),
                    _ => "Canali preferiti".to_string(),
                }),
                context_actions: tv_context_actions(Arc::clone(&catalog)),
                right_arrow_accepts_selection: true,
                // Nelle sottopagine Freccia sinistra torna alla pagina TV
                // precedente. Nella radice non deve mai cadere nell'editor:
                // per chiudere l'intera funzione resta disponibile Esc.
                left_arrow_closes: page != TvPage::Root,
                escape_stops_active_player: true,
            },
        );

        match take_tv_deferred_action() {
            Some(TvDeferredAction::Refresh) => {
                selected_id = None;
                search_query.clear();
                continue;
            }
            Some(TvDeferredAction::Record {
                channel,
                selected_id: recorded_id,
            }) => {
                // Keep the TV browser flow alive exactly like normal playback.
                // The next loop iteration recreates the same list, selects the
                // recorded channel and hides it behind mpv. Esc can therefore
                // stop playback and return directly to the channel list.
                selected_id = Some(recorded_id);
                record_tv_channel(parent, language, channel.as_ref());
                continue;
            }
            None => {}
        }

        let selected_entry_id = match selection {
            MultilineSelectionResult::Selected(id) => id,
            MultilineSelectionResult::Search(query) => {
                let query = query.trim();
                if query.is_empty() {
                    continue;
                }
                history.push((page.clone(), selected_id.clone()));
                page = TvPage::Search(query.to_string());
                selected_id = None;
                search_query = query.to_string();
                continue;
            }
            MultilineSelectionResult::SecondaryAction => {
                if page == TvPage::Favorites {
                    stream_recording::open_recordings(parent, language, StreamRecordingKind::Tv);
                } else {
                    history.push((page.clone(), selected_id.clone()));
                    page = TvPage::Favorites;
                    selected_id = None;
                    search_query.clear();
                }
                continue;
            }
            MultilineSelectionResult::Cancelled => {
                if let Some((previous_page, previous_selected_id)) = history.pop() {
                    page = previous_page;
                    selected_id = previous_selected_id;
                    search_query.clear();
                    continue;
                }
                return;
            }
        };

        let Some(entry) = entries
            .into_iter()
            .find(|entry| entry.id == selected_entry_id)
        else {
            show_error(
                parent,
                language,
                "Impossibile aprire l'elemento TV selezionato.",
            );
            continue;
        };

        selected_id = Some(entry.id.clone());
        search_query.clear();
        match entry.kind {
            TvEntryKind::Page(next_page) => {
                history.push((page, selected_id.clone()));
                page = next_page;
                selected_id = None;
            }
            TvEntryKind::Channel(index) => {
                let Some(channel) = catalog.channels.get(index) else {
                    show_error(parent, language, "Canale TV non disponibile.");
                    continue;
                };
                play_channel(parent, language, channel);
            }
            TvEntryKind::Favorite(favorite) => {
                let channel = channel_from_favorite(favorite.as_ref());
                play_channel(parent, language, &channel);
            }
            TvEntryKind::Recordings => {
                stream_recording::open_recordings(parent, language, StreamRecordingKind::Tv);
            }
        }
    }
}

fn favorite_from_channel(channel: &TvChannel) -> TvFavorite {
    TvFavorite {
        name: channel.name.clone(),
        url: channel.url.clone(),
        category: channel.category.clone(),
        stream_resolver: channel.stream_resolver.clone(),
        resolver_endpoint: channel.resolver_endpoint.clone(),
        resolver_realm: channel.resolver_realm.clone(),
        resolver_channel_id: channel.resolver_channel_id.clone(),
        tvg_id: channel.tvg_id.clone(),
        tvg_name: channel.tvg_name.clone(),
        http_user_agent: channel.http_user_agent.clone(),
    }
}

fn channel_from_favorite(favorite: &TvFavorite) -> TvChannel {
    TvChannel {
        name: favorite.name.clone(),
        url: favorite.url.clone(),
        category: favorite.category.clone(),
        stream_resolver: favorite.stream_resolver.clone(),
        resolver_endpoint: favorite.resolver_endpoint.clone(),
        resolver_realm: favorite.resolver_realm.clone(),
        resolver_channel_id: favorite.resolver_channel_id.clone(),
        tvg_id: favorite.tvg_id.clone(),
        tvg_name: favorite.tvg_name.clone(),
        http_user_agent: favorite.http_user_agent.clone(),
    }
}

fn selected_channel_from_id(catalog: &TvCatalog, id: &str) -> Option<TvChannel> {
    if let Some(raw_index) = id.strip_prefix("channel:")
        && let Ok(index) = raw_index.parse::<usize>()
    {
        return catalog.channels.get(index).cloned();
    }
    if let Some(raw_index) = id.strip_prefix("favorite:")
        && let Ok(index) = raw_index.parse::<usize>()
    {
        return load_settings()
            .tv_favorites
            .get(index)
            .map(channel_from_favorite);
    }
    None
}

fn tv_favorite_matches_channel(favorite: &TvFavorite, channel: &TvChannel) -> bool {
    favorite.name.eq_ignore_ascii_case(&channel.name)
        && (favorite.url == channel.url
            || (!favorite.tvg_id.trim().is_empty()
                && favorite.tvg_id.eq_ignore_ascii_case(&channel.tvg_id)))
}

fn tv_context_actions(catalog: Arc<TvCatalog>) -> Vec<InterpreterContextAction> {
    let add_catalog = Arc::clone(&catalog);
    let add_enabled = Arc::new(move |id: &str| {
        let Some(channel) = selected_channel_from_id(&add_catalog, id) else {
            return false;
        };
        !load_settings()
            .tv_favorites
            .iter()
            .any(|favorite| tv_favorite_matches_channel(favorite, &channel))
    });
    let add_catalog_handler = Arc::clone(&catalog);
    let add_handler = Arc::new(move |id: String| {
        let Some(channel) = selected_channel_from_id(&add_catalog_handler, &id) else {
            return;
        };
        let mut settings = load_settings();
        if !settings
            .tv_favorites
            .iter()
            .any(|favorite| tv_favorite_matches_channel(favorite, &channel))
        {
            settings.tv_favorites.push(favorite_from_channel(&channel));
            settings.tv_favorites.sort_by(|a, b| {
                a.name
                    .to_lowercase()
                    .cmp(&b.name.to_lowercase())
                    .then_with(|| a.url.cmp(&b.url))
            });
            save_settings(settings);
            crate::screen_reader_speak(&format!("{} aggiunto ai preferiti", channel.name));
            set_tv_deferred_action(TvDeferredAction::Refresh);
            close_tv_browser_dialog();
        }
    });

    let remove_catalog = Arc::clone(&catalog);
    let remove_enabled = Arc::new(move |id: &str| {
        let Some(channel) = selected_channel_from_id(&remove_catalog, id) else {
            return false;
        };
        load_settings()
            .tv_favorites
            .iter()
            .any(|favorite| tv_favorite_matches_channel(favorite, &channel))
    });
    let remove_catalog_handler = Arc::clone(&catalog);
    let remove_handler = Arc::new(move |id: String| {
        let Some(channel) = selected_channel_from_id(&remove_catalog_handler, &id) else {
            return;
        };
        let mut settings = load_settings();
        let before = settings.tv_favorites.len();
        settings
            .tv_favorites
            .retain(|favorite| !tv_favorite_matches_channel(favorite, &channel));
        if settings.tv_favorites.len() != before {
            save_settings(settings);
            crate::screen_reader_speak(&format!("{} rimosso dai preferiti", channel.name));
            set_tv_deferred_action(TvDeferredAction::Refresh);
            close_tv_browser_dialog();
        }
    });

    let record_catalog = Arc::clone(&catalog);
    let record_enabled =
        Arc::new(move |id: &str| selected_channel_from_id(&record_catalog, id).is_some());
    let record_catalog_handler = Arc::clone(&catalog);
    let record_handler = Arc::new(move |id: String| {
        let Some(channel) = selected_channel_from_id(&record_catalog_handler, &id) else {
            return;
        };
        set_tv_deferred_action(TvDeferredAction::Record {
            channel: Box::new(channel),
            selected_id: id,
        });
        close_tv_browser_dialog();
    });

    vec![
        InterpreterContextAction {
            label: "Aggiungi ai preferiti".to_string(),
            ctrl_c_shortcut: false,
            enabled: add_enabled,
            handler: add_handler,
        },
        InterpreterContextAction {
            label: "Rimuovi dai preferiti".to_string(),
            ctrl_c_shortcut: false,
            enabled: remove_enabled,
            handler: remove_handler,
        },
        InterpreterContextAction {
            label: "Registra e riproduci canale".to_string(),
            ctrl_c_shortcut: false,
            enabled: record_enabled,
            handler: record_handler,
        },
    ]
}

fn record_tv_channel(parent: HWND, language: Language, channel: &TvChannel) {
    crate::screen_reader_speak(&format!("Avvio registrazione di {}", channel.name));
    let resolved_url = match tv::resolve_stream_url(channel) {
        Ok(url) if !url.trim().is_empty() => url,
        Ok(_) => {
            show_error(parent, language, "Lo stream TV risolto è vuoto.");
            return;
        }
        Err(err) => {
            show_error(
                parent,
                language,
                &format!("Impossibile registrare {}: {err}", channel.name),
            );
            return;
        }
    };

    match stream_recording::start_tv_recording_and_playback(
        parent,
        &resolved_url,
        &channel.name,
        channel.media_playback_user_agent(),
        tv::is_rai_audio_description_channel(channel),
    ) {
        Ok(path) => {
            crate::screen_reader_speak(&format!(
                "Registrazione di {} avviata. Il file sarà salvato in {}. La registrazione terminerà chiudendo il player.",
                channel.name,
                path.display()
            ));
        }
        Err(err) => show_error(parent, language, &err),
    }
}

fn play_channel(parent: HWND, language: Language, channel: &TvChannel) {
    crate::screen_reader_speak(&format!("Avvio {}", channel.name));
    let resolved_url = match tv::resolve_stream_url(channel) {
        Ok(url) if !url.trim().is_empty() => url,
        Ok(_) => {
            show_error(parent, language, "Lo stream TV risolto è vuoto.");
            return;
        }
        Err(err) => {
            show_error(
                parent,
                language,
                &format!("Impossibile avviare {}: {err}", channel.name),
            );
            return;
        }
    };

    crate::log_debug(&format!(
        "TV playback start: name={} category={} resolver={} user_agent={}",
        channel.name,
        channel.category,
        channel.stream_resolver.as_deref().unwrap_or("direct"),
        channel.playback_user_agent()
    ));
    if let Err(err) = crate::launch_tv_stream_in_mpv(
        parent,
        &resolved_url,
        &channel.name,
        channel.media_playback_user_agent(),
        tv::is_rai_audio_description_channel(channel),
    ) {
        show_error(
            parent,
            language,
            &format!("Impossibile riprodurre {}: {err}", channel.name),
        );
    }
}

fn page_title(page: &TvPage) -> String {
    match page {
        TvPage::Root => "TV in diretta".to_string(),
        TvPage::Category(category) => format!("TV - {category}"),
        TvPage::Regions => "TV regionali".to_string(),
        TvPage::Region(region) => format!("TV regionali - {region}"),
        TvPage::Favorites => "Canali TV preferiti".to_string(),
        TvPage::Search(query) => format!("Risultati TV per {query}"),
    }
}

fn push_group_index(groups: &mut Vec<(String, Vec<usize>)>, name: &str, index: usize) {
    if let Some((_, indices)) = groups.iter_mut().find(|(existing, _)| existing == name) {
        indices.push(index);
    } else {
        groups.push((name.to_string(), vec![index]));
    }
}

fn channel_count_description(count: usize) -> String {
    if count == 1 {
        "1 canale".to_string()
    } else {
        format!("{count} canali")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn channel(name: &str, category: &str) -> TvChannel {
        TvChannel {
            name: name.to_string(),
            url: "https://example.test/live.m3u8".to_string(),
            category: category.to_string(),
            stream_resolver: None,
            resolver_endpoint: None,
            resolver_realm: None,
            resolver_channel_id: None,
            tvg_id: String::new(),
            tvg_name: String::new(),
            http_user_agent: String::new(),
        }
    }

    #[test]
    fn catalog_groups_regional_channels_separately() {
        let catalog = TvCatalog::new(
            vec![
                channel("Rai 1", "Rai"),
                channel("TGR Piemonte", "Regionali - Piemonte"),
            ],
            HashMap::new(),
        );
        assert_eq!(catalog.categories.len(), 1);
        assert_eq!(catalog.regions.len(), 1);
        assert_eq!(catalog.regions[0].0, "Piemonte");
    }
}
