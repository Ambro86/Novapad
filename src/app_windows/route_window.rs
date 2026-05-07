use windows::Win32::Foundation::HWND;

use crate::app_windows::prompt_window::{self, PromptDirectoryOptions};
use crate::app_windows::route_service::{
    GeocodeCandidate, RouteAvoid, RouteClient, RoutePath, RoutePreference, RouteProfile,
    RouteRequestResult, RouteResult, format_distance, format_duration,
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
    match client.route_from_addresses(
        &params.from,
        &params.to,
        params.plan.profile,
        params.plan.preference,
        params.plan.avoid,
    ) {
        Ok(RouteRequestResult::Ready(route)) => {
            display_route(parent, language, route);
        }
        Ok(RouteRequestResult::NeedsSelection {
            from_candidates,
            to_candidates,
            profile,
            preference,
            avoid,
        }) => {
            let plan = RoutePlan {
                profile,
                preference,
                avoid,
            };
            handle_selection_and_route(
                parent,
                language,
                &client,
                from_candidates,
                to_candidates,
                plan,
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
    plan: RoutePlan,
}

#[derive(Clone, Copy)]
struct RoutePlan {
    profile: RouteProfile,
    preference: RoutePreference,
    avoid: RouteAvoid,
}

fn prompt_route_params(parent: HWND, language: Language) -> Option<RouteParams> {
    let profile_options = [
        RouteProfile::Walking,
        RouteProfile::Cycling,
        RouteProfile::Driving,
        RouteProfile::Wheelchair,
    ];
    let preference_options = [RoutePreference::Fastest, RoutePreference::Shortest];
    let avoid_options = [
        RouteAvoid::None,
        RouteAvoid::Highways,
        RouteAvoid::Tollways,
        RouteAvoid::HighwaysAndTollways,
    ];

    let options_list = profile_options
        .iter()
        .map(|profile| profile.label_it().to_string())
        .collect();
    let preference_list = preference_options
        .iter()
        .map(|preference| preference.label_it().to_string())
        .collect();
    let avoid_list = avoid_options
        .iter()
        .map(|avoid| {
            if *avoid == RouteAvoid::None {
                "nessuna".to_string()
            } else {
                avoid.label_it().to_string()
            }
        })
        .collect();

    let options = PromptDirectoryOptions {
        title: "Percorsi e navigazione".to_string(),
        type_label: "Mezzo".to_string(),
        options: options_list,
        default_selection: 0,
        secondary_type_label: "Tipo".to_string(),
        secondary_options: preference_list,
        secondary_default_selection: 0,
        tertiary_type_label: "Solo auto: evita".to_string(),
        tertiary_options: avoid_list,
        tertiary_default_selection: 0,
        focus_primary_field: false,
        primary_label: "Partenza:".to_string(),
        primary_labels: Vec::new(),
        primary_default: String::new(),
        secondary_label: "Arrivo:".to_string(),
        secondary_default: String::new(),
        tertiary_label: String::new(),
        tertiary_default: String::new(),
    };

    let result = prompt_window::prompt_directory_search(parent, options, language)?;

    let profile = profile_options
        .get(result.selected_index)
        .copied()
        .unwrap_or(RouteProfile::Walking);
    let preference = preference_options
        .get(result.secondary_selected_index)
        .copied()
        .unwrap_or(RoutePreference::Fastest);
    let avoid = if profile == RouteProfile::Driving {
        avoid_options
            .get(result.tertiary_selected_index)
            .copied()
            .unwrap_or(RouteAvoid::None)
    } else {
        RouteAvoid::None
    };

    Some(RouteParams {
        from: result.primary_value,
        to: result.secondary_value,
        plan: RoutePlan {
            profile,
            preference,
            avoid,
        },
    })
}

fn handle_selection_and_route(
    parent: HWND,
    language: Language,
    client: &RouteClient,
    from_candidates: Vec<GeocodeCandidate>,
    to_candidates: Vec<GeocodeCandidate>,
    plan: RoutePlan,
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
                context_actions: Vec::new(),
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
                context_actions: Vec::new(),
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
    match client.route_between_coordinates(
        &from_selected,
        &to_selected,
        plan.profile,
        plan.preference,
        plan.avoid,
    ) {
        Ok(route) => display_route(parent, language, route),
        Err(err) => show_error(parent, language, &err.to_string()),
    }
}

fn display_route(parent: HWND, _language: Language, route: RouteResult) {
    let Some(route) = select_route_path(parent, _language, route) else {
        return;
    };
    let text = route.format_for_speech_or_text();
    editor_manager::new_document(parent);
    editor_manager::set_current_document_title(parent, "Percorso");
    if let Some(hwnd_edit) = crate::get_active_edit(parent) {
        editor_manager::set_edit_text(hwnd_edit, &text);
    }
}

fn select_route_path(parent: HWND, language: Language, route: RouteResult) -> Option<RouteResult> {
    crate::log_debug(&format!("route_select_path: paths={}", route.paths.len()));
    if route.paths.len() <= 1 {
        return Some(route);
    }

    let items = route
        .paths
        .iter()
        .enumerate()
        .map(|(idx, path)| route_path_selection_item(idx, path))
        .collect();

    let selected = youtube_transcript_window::select_multiline_items_with_search(
        parent,
        language,
        "Scegli il percorso".to_string(),
        items,
        None,
        MultilineSearchOptions {
            initial_query: String::new(),
            search_button_label: String::new(),
            show_search_edit: false,
            secondary_action_label: None,
            context_actions: Vec::new(),
            right_arrow_accepts_selection: true,
            left_arrow_closes: true,
        },
    );

    match selected {
        MultilineSelectionResult::Selected(id_str) => {
            let idx = id_str.parse::<usize>().unwrap_or(0);
            let selected_path = route.paths.get(idx)?.clone();
            Some(RouteResult {
                paths: vec![selected_path],
                ..route
            })
        }
        _ => None,
    }
}

fn route_path_selection_item(index: usize, path: &RoutePath) -> MultilineSelectionItem {
    let title = if index == 0 {
        "Percorso principale".to_string()
    } else {
        format!("Alternativa {index}")
    };
    let description = format!(
        "{}; {}",
        format_distance(path.distance_meters),
        format_duration(path.duration_seconds)
    );

    MultilineSelectionItem {
        id: index.to_string(),
        title,
        description: Some(description),
    }
}
