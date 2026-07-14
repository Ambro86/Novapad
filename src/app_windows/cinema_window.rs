use chrono::{Datelike, NaiveDate};
use windows::Win32::Foundation::HWND;

use crate::app_windows::cinema_service::{CinemaClient, CinemaMovie};
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::i18n;
use crate::settings::Language;
use crate::{show_error, with_state};

#[derive(Clone, Copy, PartialEq, Eq)]
enum MovieListKind {
    NowPlaying,
    Upcoming,
}

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let client = CinemaClient::default();
    crate::screen_reader_speak(&i18n::tr(language, "cinema.loading"));

    let mut movies = match client.now_playing(language) {
        Ok(movies) => movies,
        Err(error) => {
            show_cinema_error(parent, language, &error);
            return;
        }
    };
    movies.sort_by(|left, right| right.release_date.cmp(&left.release_date));
    browse_movies(parent, language, &client, MovieListKind::NowPlaying, movies);
}

fn browse_movies(
    parent: HWND,
    language: Language,
    client: &CinemaClient,
    kind: MovieListKind,
    movies: Vec<CinemaMovie>,
) {
    if movies.is_empty() {
        show_error(parent, language, &i18n::tr(language, "cinema.no_movies"));
        return;
    }

    let mut query = String::new();
    let mut selected_id: Option<String> = None;

    loop {
        let filtered = filter_movies(&movies, &query);
        if filtered.is_empty() {
            show_error(parent, language, &i18n::tr(language, "cinema.no_movies"));
            query.clear();
            selected_id = None;
            continue;
        }

        let items = filtered
            .iter()
            .map(|movie| MultilineSelectionItem {
                id: movie.id.to_string(),
                title: movie.title.clone(),
                description: Some(release_text(movie, language)),
            })
            .collect();

        let title = match kind {
            MovieListKind::NowPlaying => i18n::tr(language, "cinema.title"),
            MovieListKind::Upcoming => i18n::tr(language, "cinema.upcoming"),
        };
        let result = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            title,
            items,
            selected_id.clone(),
            MultilineSearchOptions {
                initial_query: query.clone(),
                search_button_label: i18n::tr(language, "cinema.search"),
                show_search_edit: true,
                secondary_action_label: (kind == MovieListKind::NowPlaying)
                    .then(|| i18n::tr(language, "cinema.upcoming")),
                context_actions: Vec::new(),
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: true,
                refresh: None,
            },
        );

        match result {
            MultilineSelectionResult::Cancelled => return,
            MultilineSelectionResult::Search(value) => {
                query = value.trim().to_string();
                selected_id = None;
            }
            MultilineSelectionResult::SecondaryAction => {
                crate::screen_reader_speak(&i18n::tr(language, "cinema.loading"));
                let mut upcoming = match client.upcoming(language) {
                    Ok(movies) => movies,
                    Err(error) => {
                        show_cinema_error(parent, language, &error);
                        continue;
                    }
                };
                upcoming.sort_by(|left, right| left.release_date.cmp(&right.release_date));
                browse_movies(parent, language, client, MovieListKind::Upcoming, upcoming);
            }
            MultilineSelectionResult::Selected(id) => {
                selected_id = Some(id.clone());
                if let Some(movie) = movies.iter().find(|movie| movie.id.to_string() == id) {
                    show_movie_detail(parent, language, client, movie);
                }
            }
        }
    }
}

fn show_movie_detail(parent: HWND, language: Language, client: &CinemaClient, movie: &CinemaMovie) {
    let overview = if movie.overview.trim().is_empty() {
        i18n::tr(language, "cinema.no_overview")
    } else {
        movie.overview.trim().to_string()
    };
    let description = format!(
        "{}\n\n{}\n{}",
        release_text(movie, language),
        i18n::tr(language, "cinema.overview"),
        overview
    );

    loop {
        let result = youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            movie.title.clone(),
            vec![MultilineSelectionItem {
                id: movie.id.to_string(),
                title: movie.title.clone(),
                description: Some(description.clone()),
            }],
            Some(movie.id.to_string()),
            MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: String::new(),
                show_search_edit: false,
                secondary_action_label: Some(i18n::tr(language, "cinema.open_trailer")),
                context_actions: Vec::new(),
                right_arrow_accepts_selection: false,
                left_arrow_closes: true,
                escape_stops_active_player: false,
                refresh: None,
            },
        );

        match result {
            MultilineSelectionResult::SecondaryAction => {
                crate::screen_reader_speak(&i18n::tr(language, "cinema.loading_trailer"));
                match client.trailer_url(movie.id, language) {
                    Ok(Some(url)) => {
                        if let Err(error) = youtube_transcript_window::play_youtube_video_in_mpv(
                            parent,
                            &url,
                            &movie.title,
                        ) {
                            let message = i18n::tr_f(
                                language,
                                "cinema.trailer_error",
                                &[("err", error.as_str())],
                            );
                            show_error(parent, language, &message);
                        }
                    }
                    Ok(None) => show_error(
                        parent,
                        language,
                        &i18n::tr(language, "cinema.trailer_unavailable"),
                    ),
                    Err(error) => show_cinema_error(parent, language, &error),
                }
                return;
            }
            MultilineSelectionResult::Cancelled | MultilineSelectionResult::Selected(_) => return,
            MultilineSelectionResult::Search(_) => {}
        }
    }
}

fn filter_movies<'a>(movies: &'a [CinemaMovie], query: &str) -> Vec<&'a CinemaMovie> {
    let normalized = query.trim().to_lowercase();
    if normalized.is_empty() {
        return movies.iter().collect();
    }
    movies
        .iter()
        .filter(|movie| {
            movie.title.to_lowercase().contains(&normalized)
                || movie.overview.to_lowercase().contains(&normalized)
        })
        .collect()
}

fn release_text(movie: &CinemaMovie, language: Language) -> String {
    if movie.release_date.trim().is_empty() {
        return String::new();
    }
    let date = format_date(&movie.release_date, language);
    let future = NaiveDate::parse_from_str(&movie.release_date, "%Y-%m-%d")
        .map(|value| value > chrono::Local::now().date_naive())
        .unwrap_or(false);
    let key = if future {
        "cinema.will_release"
    } else {
        "cinema.released"
    };
    i18n::tr_f(language, key, &[("date", date.as_str())])
}

fn format_date(raw: &str, language: Language) -> String {
    let Ok(date) = NaiveDate::parse_from_str(raw, "%Y-%m-%d") else {
        return raw.to_string();
    };
    match language {
        Language::English => format!("{:02}/{:02}/{}", date.month(), date.day(), date.year()),
        Language::Chinese => format!("{}年{}月{}日", date.year(), date.month(), date.day()),
        _ => format!("{:02}/{:02}/{}", date.day(), date.month(), date.year()),
    }
}

fn show_cinema_error(parent: HWND, language: Language, error: &str) {
    crate::log_debug(&format!("Cinema error: {error}"));
    let message = format!("{}\n{}", i18n::tr(language, "cinema.error"), error);
    show_error(parent, language, &message);
}
