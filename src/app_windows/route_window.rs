use windows::Win32::Foundation::HWND;

use crate::app_windows::prompt_window::{self, PromptDirectoryOptions};
use crate::app_windows::route_service::{
    GeocodeCandidate, RouteClient, RouteProfile, RouteRequestResult, RouteResult,
};
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::editor_manager;
use crate::settings::Language;
use crate::{show_error, with_state};

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    if language != Language::Italian {
        return;
    }

    run_route_flow(parent, language);
}

fn run_route_flow(parent: HWND, language: Language) {
    let client = RouteClient::default();

    // 1. Chiedi partenza, arrivo e profilo
    let Some(params) = prompt_route_params(parent, language) else {
        return;
    };

    crate::screen_reader_speak("Ricerca percorso in corso");

    // 2. Esegui la ricerca
    match client.route_from_addresses(&params.from, &params.to, params.profile) {
        Ok(RouteRequestResult::Ready(route)) => {
            display_route(parent, language, route);
        }
        Ok(RouteRequestResult::NeedsSelection {
            from_candidates,
            to_candidates,
            profile,
        }) => {
            handle_selection_and_route(
                parent,
                language,
                &client,
                from_candidates,
                to_candidates,
                profile,
            );
        }
        Err(err) => {
            show_error(parent, language, &err.to_string());
        }
    }
}

struct RouteParams {
    from: String,
    to: String,
    profile: RouteProfile,
}

fn prompt_route_params(parent: HWND, language: Language) -> Option<RouteParams> {
    let profile_options = vec![
        RouteProfile::Walking.label_it().to_string(),
        RouteProfile::Cycling.label_it().to_string(),
        RouteProfile::Driving.label_it().to_string(),
        RouteProfile::Wheelchair.label_it().to_string(),
    ];

    let options = PromptDirectoryOptions {
        title: "Percorsi e navigazione".to_string(),
        type_label: "Inserisci gli indirizzi e scegli il mezzo di trasporto.".to_string(),
        options: profile_options,
        default_selection: 0,
        focus_primary_field: true,
        primary_label: "Partenza:".to_string(),
        primary_labels: Vec::new(),
        primary_default: String::new(),
        secondary_label: "Arrivo:".to_string(),
        secondary_default: String::new(),
        tertiary_label: String::new(),
        tertiary_default: String::new(),
    };

    let result = prompt_window::prompt_directory_search(parent, options, language)?;

    let profile = match result.selected_index {
        0 => RouteProfile::Walking,
        1 => RouteProfile::Cycling,
        2 => RouteProfile::Driving,
        3 => RouteProfile::Wheelchair,
        _ => RouteProfile::Walking,
    };

    Some(RouteParams {
        from: result.primary_value,
        to: result.secondary_value,
        profile,
    })
}

fn handle_selection_and_route(
    parent: HWND,
    language: Language,
    client: &RouteClient,
    from_candidates: Vec<GeocodeCandidate>,
    to_candidates: Vec<GeocodeCandidate>,
    profile: RouteProfile,
) {
    let from_selected = if from_candidates.len() > 1 {
        let items = from_candidates
            .iter()
            .enumerate()
            .map(|(idx, c)| MultilineSelectionItem {
                id: idx.to_string(),
                title: c.display_label(),
                description: None,
            })
            .collect();

        match youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            "Scegli la partenza".to_string(),
            items,
            None,
            MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: String::new(),
                show_search_edit: false,
                secondary_action_label: None,
                context_action: None,
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
            },
        ) {
            MultilineSelectionResult::Selected(id_str) => {
                let idx = id_str.parse::<usize>().unwrap_or(0);
                from_candidates[idx].clone()
            }
            _ => return,
        }
    } else {
        from_candidates[0].clone()
    };

    let to_selected = if to_candidates.len() > 1 {
        let items = to_candidates
            .iter()
            .enumerate()
            .map(|(idx, c)| MultilineSelectionItem {
                id: idx.to_string(),
                title: c.display_label(),
                description: None,
            })
            .collect();

        match youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            "Scegli la destinazione".to_string(),
            items,
            None,
            MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: String::new(),
                show_search_edit: false,
                secondary_action_label: None,
                context_action: None,
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
            },
        ) {
            MultilineSelectionResult::Selected(id_str) => {
                let idx = id_str.parse::<usize>().unwrap_or(0);
                to_candidates[idx].clone()
            }
            _ => return,
        }
    } else {
        to_candidates[0].clone()
    };

    crate::screen_reader_speak("Calcolo percorso in corso");
    match client.route_between_coordinates(&from_selected, &to_selected, profile) {
        Ok(route) => display_route(parent, language, route),
        Err(err) => show_error(parent, language, &err.to_string()),
    }
}

fn display_route(parent: HWND, _language: Language, route: RouteResult) {
    let text = route.format_for_speech_or_text();
    editor_manager::new_document(parent);
    editor_manager::set_current_document_title(parent, "Percorso");
    if let Some(hwnd_edit) = crate::get_active_edit(parent) {
        editor_manager::set_edit_text(hwnd_edit, &text);
    }
}
