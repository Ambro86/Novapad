use windows::Win32::Foundation::HWND;

use crate::app_windows::prompt_window;
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::editor_manager;
use crate::settings::Language;
use crate::tools::italiaonline::{
    self, DetailResponse, DirectoryKind, SearchQuery, SearchResponse,
};
use crate::{show_error, with_state};

const ITALIAONLINE_PREVIOUS_PAGE_ID: &str = "__italiaonline_previous_page__";
const ITALIAONLINE_NEXT_PAGE_ID: &str = "__italiaonline_next_page__";

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }

    let initial_query = SearchQuery {
        kind: DirectoryKind::PagineBianche,
        what: String::new(),
        where_: String::new(),
        page: 1,
    };
    run_search_flow(parent, language, initial_query);
}

pub fn reopen_last(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }
    if !crate::app_windows::rai_audiodescrizioni_window::ensure_rai_luce_access(parent, language) {
        return;
    }
    let Some((query, selected_id)) = with_state(parent, |state| {
        Some((
            state.last_italiaonline_query.clone()?,
            state.last_italiaonline_result_id.clone(),
        ))
    })
    .flatten() else {
        open(parent);
        return;
    };
    crate::screen_reader_speak(&format!("Ricerca {} in corso", query.kind.label()));
    let response = match italiaonline::search(&query) {
        Ok(response) => response,
        Err(err) => {
            show_error(parent, language, &err);
            return;
        }
    };
    if response.results.is_empty() {
        show_error(parent, language, "Nessun risultato trovato.");
        return;
    }
    match browse_results(parent, language, &query, &response, selected_id) {
        BrowseResultsOutcome::OpenedDetail => {}
        BrowseResultsOutcome::NewSearch => {
            editor_manager::mark_current_document_from_italiaonline(parent, false);
            run_search_flow(parent, language, query);
        }
        BrowseResultsOutcome::Closed => {
            editor_manager::mark_current_document_from_italiaonline(parent, false);
        }
    }
}

fn run_search_flow(parent: HWND, language: Language, mut initial_query: SearchQuery) {
    loop {
        let Some(query) = prompt_search_query(parent, language, &initial_query) else {
            return;
        };
        initial_query = query.clone();
        crate::screen_reader_speak(&format!("Ricerca {} in corso", query.kind.label()));
        let response = match italiaonline::search(&query) {
            Ok(response) => response,
            Err(err) => {
                show_error(parent, language, &err);
                continue;
            }
        };
        if response.results.is_empty() {
            show_error(parent, language, "Nessun risultato trovato.");
            continue;
        }
        match browse_results(parent, language, &query, &response, None) {
            BrowseResultsOutcome::OpenedDetail => return,
            BrowseResultsOutcome::NewSearch => continue,
            BrowseResultsOutcome::Closed => return,
        }
    }
}

enum BrowseResultsOutcome {
    Closed,
    OpenedDetail,
    NewSearch,
}

fn prompt_search_query(
    parent: HWND,
    language: Language,
    initial: &SearchQuery,
) -> Option<SearchQuery> {
    prompt_search_query_with_focus(parent, language, initial, false)
}

fn prompt_search_query_with_focus(
    parent: HWND,
    language: Language,
    initial: &SearchQuery,
    focus_primary_field: bool,
) -> Option<SearchQuery> {
    let kind_options = vec![
        DirectoryKind::PagineBianche.label().to_string(),
        DirectoryKind::PagineGialle.label().to_string(),
    ];
    let default_kind = match initial.kind {
        DirectoryKind::PagineBianche => 0,
        DirectoryKind::PagineGialle => 1,
    };
    let prompt = prompt_window::prompt_directory_search(
        parent,
        prompt_window::PromptDirectoryOptions {
            title: "Ricerca nominativi".to_string(),
            type_label: "Tipo".to_string(),
            options: kind_options,
            default_selection: default_kind,
            focus_primary_field,
            primary_label: initial.kind.primary_field_label().to_string(),
            primary_labels: vec![
                DirectoryKind::PagineBianche
                    .primary_field_label()
                    .to_string(),
                DirectoryKind::PagineGialle
                    .primary_field_label()
                    .to_string(),
            ],
            primary_default: initial.what.clone(),
            secondary_label: "Dove".to_string(),
            secondary_default: initial.where_.clone(),
        },
        language,
    )?;
    let kind = if prompt.selected_index == 0 {
        DirectoryKind::PagineBianche
    } else {
        DirectoryKind::PagineGialle
    };
    let trimmed_what = prompt.primary_value.trim();
    if trimmed_what.is_empty() {
        show_error(
            parent,
            language,
            &format!("Il campo {} è vuoto.", kind.primary_field_name()),
        );
        return prompt_search_query_with_focus(
            parent,
            language,
            &SearchQuery {
                kind,
                what: prompt.primary_value,
                where_: prompt.secondary_value,
                page: 1,
            },
            true,
        );
    }

    Some(SearchQuery {
        kind,
        what: trimmed_what.to_string(),
        where_: prompt.secondary_value.trim().to_string(),
        page: 1,
    })
}

fn browse_results(
    parent: HWND,
    language: Language,
    query: &SearchQuery,
    response: &SearchResponse,
    initial_selected_id: Option<String>,
) -> BrowseResultsOutcome {
    let mut current_query = query.clone();
    let mut current_response = response.clone();
    let mut current_selected_id = initial_selected_id;

    loop {
        let title = results_title(&current_query, &current_response);
        let mut items = Vec::new();
        if current_response.current_page > 1 {
            items.push(MultilineSelectionItem {
                id: ITALIAONLINE_PREVIOUS_PAGE_ID.to_string(),
                title: "Risultati precedenti".to_string(),
                description: Some(format!(
                    "Torna alla pagina {} dei risultati",
                    current_response.current_page.saturating_sub(1)
                )),
            });
        }
        items.extend(
            current_response
                .results
                .iter()
                .map(|result| MultilineSelectionItem {
                    id: result.id.clone(),
                    title: result.name.clone(),
                    description: Some(format_result_description(result)),
                }),
        );
        if !current_response.is_last_page {
            items.push(MultilineSelectionItem {
                id: ITALIAONLINE_NEXT_PAGE_ID.to_string(),
                title: "Risultati successivi".to_string(),
                description: Some(format!(
                    "Vai alla pagina {} dei risultati",
                    current_response.current_page + 1
                )),
            });
        }

        match youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            title,
            items,
            current_selected_id.clone(),
            youtube_transcript_window::MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: "Nuova ricerca".to_string(),
                show_search_edit: false,
                context_action: None,
                right_arrow_accepts_selection: true,
                left_arrow_closes: false,
            },
        ) {
            MultilineSelectionResult::Selected(id) => {
                if id == ITALIAONLINE_PREVIOUS_PAGE_ID || id == ITALIAONLINE_NEXT_PAGE_ID {
                    let next_page = if id == ITALIAONLINE_PREVIOUS_PAGE_ID {
                        current_query.page.saturating_sub(1).max(1)
                    } else {
                        current_query.page + 1
                    };
                    crate::screen_reader_speak("Caricamento risultati");
                    let next_query = SearchQuery {
                        page: next_page,
                        ..current_query.clone()
                    };
                    match italiaonline::search(&next_query) {
                        Ok(next_response) => {
                            current_query = next_query;
                            current_response = next_response;
                            current_selected_id = None;
                            continue;
                        }
                        Err(err) => {
                            show_error(parent, language, &err);
                            continue;
                        }
                    }
                }
                crate::screen_reader_speak("Caricamento dettaglio");
                match italiaonline::load_detail(&current_query, &id) {
                    Ok(detail) => {
                        with_state(parent, |state| {
                            state.last_italiaonline_query = Some(current_query.clone());
                            state.last_italiaonline_result_id = Some(id.clone());
                        });
                        open_detail_document(parent, detail);
                        return BrowseResultsOutcome::OpenedDetail;
                    }
                    Err(err) => {
                        show_error(parent, language, &err);
                        continue;
                    }
                }
            }
            MultilineSelectionResult::Cancelled => return BrowseResultsOutcome::Closed,
            MultilineSelectionResult::Search(_) => return BrowseResultsOutcome::NewSearch,
        }
    }
}

fn results_title(query: &SearchQuery, response: &SearchResponse) -> String {
    let mut title = format!("{} - {}", query.kind.label(), query.what.trim());
    let where_display = response
        .display_where
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .or_else(|| {
            let raw = query.where_.trim();
            if raw.is_empty() { None } else { Some(raw) }
        });
    if let Some(where_display) = where_display {
        title.push_str(" - ");
        title.push_str(where_display);
    }
    if response.current_page > 1 {
        title.push_str(&format!(" - Pagina {}", response.current_page));
    }
    title
}

fn format_result_description(result: &italiaonline::SearchResult) -> String {
    let mut parts = Vec::new();
    if let Some(category) = result
        .category
        .as_deref()
        .filter(|value| !value.trim().is_empty())
    {
        parts.push(category.to_string());
    }
    if let Some(address) = result
        .address
        .as_deref()
        .filter(|value| !value.trim().is_empty())
    {
        parts.push(address.to_string());
    }
    let locality = match (
        result
            .city
            .as_deref()
            .map(str::trim)
            .filter(|value| !value.is_empty()),
        result
            .province
            .as_deref()
            .map(str::trim)
            .filter(|value| !value.is_empty()),
    ) {
        (Some(city), Some(province)) => Some(format!("{city} ({province})")),
        (Some(city), None) => Some(city.to_string()),
        (None, Some(province)) => Some(province.to_string()),
        (None, None) => None,
    };
    if let Some(locality) = locality {
        parts.push(locality);
    }
    if !result.phones.is_empty() {
        parts.push(result.phones.join(" - "));
    }
    parts.join(" - ")
}

fn open_detail_document(parent: HWND, detail: DetailResponse) {
    editor_manager::new_document(parent);
    editor_manager::set_current_document_title(parent, &detail.title);
    editor_manager::mark_current_document_from_italiaonline(parent, true);
    if with_state(parent, |state| {
        state
            .docs
            .get(state.current)
            .map(|doc| editor_manager::set_edit_text(doc.hwnd_edit, &detail.body))
    })
    .flatten()
    .is_none()
    {
        crate::log_debug("Failed to populate Italiaonline detail document");
    }
}
