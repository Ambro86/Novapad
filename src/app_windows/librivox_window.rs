use reqwest::blocking::Client;
use serde_json::Value;
use std::collections::{HashMap, HashSet};
use std::sync::{Mutex, OnceLock};
use std::time::Duration;
use url::Url;
use windows::Win32::Foundation::HWND;

use crate::app_windows::prompt_window;
use crate::app_windows::youtube_transcript_window::{
    self, MultilineSearchOptions, MultilineSelectionItem, MultilineSelectionResult,
};
use crate::settings::Language;
use crate::{RaiAudioOrigin, i18n, show_error, with_state};

const API_BASE: &str = "https://librivox.org/api/feed/audiobooks";
const USER_AGENT: &str = concat!("Sonarpad/", env!("CARGO_PKG_VERSION"));
const PAGE_LIMIT: usize = 50;
const SEARCH_CANDIDATE_LIMIT: usize = 100;
const FALLBACK_PAGE_LIMIT: usize = 500;
const FALLBACK_MAX_BOOKS: usize = 2_000;

#[derive(Clone)]
struct LibrivoxClient {
    client: Client,
}

#[derive(Clone)]
struct LibrivoxBook {
    id: i64,
    title: String,
    description: String,
    language: String,
    total_time: String,
    authors: Vec<String>,
    sections: Vec<LibrivoxTrack>,
}

#[derive(Clone)]
struct LibrivoxTrack {
    id: i64,
    number: i64,
    title: String,
    listen_url: String,
    play_time: String,
}

#[derive(Clone)]
struct LibrivoxParentListContext {
    query: String,
    books: Vec<LibrivoxBook>,
    has_more: bool,
    offset: usize,
    selected_id: String,
}

#[derive(Clone)]
struct LibrivoxPlayerReturnContext {
    book: LibrivoxBook,
    selected_id: String,
    selected_url: String,
    parent_list: LibrivoxParentListContext,
}

static LIBRIVOX_PLAYER_RETURN_CONTEXT: OnceLock<Mutex<Option<LibrivoxPlayerReturnContext>>> =
    OnceLock::new();

fn player_return_context() -> &'static Mutex<Option<LibrivoxPlayerReturnContext>> {
    LIBRIVOX_PLAYER_RETURN_CONTEXT.get_or_init(|| Mutex::new(None))
}

fn remember_player_return_context(context: LibrivoxPlayerReturnContext) {
    if let Ok(mut stored) = player_return_context().lock() {
        *stored = Some(context);
    }
}

fn clear_player_return_context() {
    if let Ok(mut stored) = player_return_context().lock() {
        *stored = None;
    }
}

pub(crate) fn restore_chapter_list_after_stop(parent: HWND, stopped_url: Option<&str>) -> bool {
    let Some(stopped_url) = stopped_url else {
        return false;
    };
    let context = player_return_context()
        .lock()
        .ok()
        .and_then(|stored| stored.clone())
        .filter(|context| context.selected_url == stopped_url);
    let Some(context) = context else {
        return false;
    };

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    clear_player_return_context();
    let parent_list = context.parent_list.clone();
    let started_playback = browse_loaded_book(
        parent,
        language,
        context.book,
        Some(context.selected_id),
        parent_list.clone(),
    );
    if !started_playback {
        browse_parent_list(parent, language, parent_list);
    }
    true
}

struct LibrivoxPage {
    books: Vec<LibrivoxBook>,
    has_more: bool,
}

impl Default for LibrivoxClient {
    fn default() -> Self {
        let client = Client::builder()
            .user_agent(USER_AGENT)
            .timeout(Duration::from_secs(30))
            .build()
            .unwrap_or_else(|error| {
                crate::log_debug(&format!("LibriVox HTTP client setup failed: {error}"));
                Client::new()
            });
        Self { client }
    }
}

impl LibrivoxClient {
    fn search(&self, query: &str, offset: usize, limit: usize) -> Result<LibrivoxPage, String> {
        let normalized_query = query.trim();
        if normalized_query.is_empty() {
            let books = self.fetch_books(limit, offset, None, None)?;
            let has_more = books.len() >= limit;
            return Ok(LibrivoxPage { books, has_more });
        }

        let terms = search_terms(normalized_query);
        if terms.is_empty() {
            return Ok(LibrivoxPage {
                books: Vec::new(),
                has_more: false,
            });
        }

        let mut candidates = HashMap::<i64, LibrivoxBook>::new();
        let mut api_terms = vec![normalized_query.to_string()];
        for term in &terms {
            if !api_terms.iter().any(|value| value == term) {
                api_terms.push(term.clone());
            }
        }

        for term in api_terms {
            for field in ["title", "author"] {
                if let Ok(books) =
                    self.fetch_books(SEARCH_CANDIDATE_LIMIT, 0, Some(field), Some(&term))
                {
                    for book in books {
                        candidates.insert(book.id, book);
                    }
                }
            }
        }

        let target_count = offset.saturating_add(limit).saturating_add(1);
        let mut matches = sorted_matches(candidates.values(), &terms);
        let mut reached_end = true;
        if matches.len() < target_count {
            reached_end = false;
            let mut scan_offset = 0usize;
            let mut scanned = 0usize;
            while scan_offset < FALLBACK_MAX_BOOKS {
                let books = match self.fetch_books(FALLBACK_PAGE_LIMIT, scan_offset, None, None) {
                    Ok(books) => books,
                    Err(_) => {
                        reached_end = true;
                        break;
                    }
                };
                if books.is_empty() {
                    reached_end = true;
                    break;
                }
                for book in books.iter().filter(|book| matches_all_terms(book, &terms)) {
                    candidates.insert(book.id, book.clone());
                }
                if candidates
                    .values()
                    .filter(|book| matches_all_terms(book, &terms))
                    .count()
                    >= target_count
                {
                    break;
                }
                scanned = scanned.saturating_add(books.len());
                scan_offset = scan_offset.saturating_add(books.len());
                if books.len() < FALLBACK_PAGE_LIMIT || scanned >= FALLBACK_MAX_BOOKS {
                    reached_end = true;
                    break;
                }
            }
            matches = sorted_matches(candidates.values(), &terms);
        }

        let start = offset.min(matches.len());
        let end = start.saturating_add(limit).min(matches.len());
        Ok(LibrivoxPage {
            books: matches[start..end].to_vec(),
            has_more: end < matches.len() || !reached_end,
        })
    }

    fn fetch_book(&self, id: i64) -> Result<LibrivoxBook, String> {
        let id_value = id.to_string();
        let books = self.fetch_books_with_extra_query(1, 0, &[("id", id_value.as_str())])?;
        books
            .into_iter()
            .next()
            .ok_or_else(|| "Audiobook not found".to_string())
    }

    fn fetch_books(
        &self,
        limit: usize,
        offset: usize,
        search_field: Option<&str>,
        search_value: Option<&str>,
    ) -> Result<Vec<LibrivoxBook>, String> {
        let mut url = Url::parse(API_BASE).map_err(|error| error.to_string())?;
        if let (Some(field), Some(value)) = (search_field, search_value)
            && matches!(field, "title" | "author" | "genre")
            && !value.trim().is_empty()
        {
            let mut segments = url
                .path_segments_mut()
                .map_err(|_| "Invalid LibriVox URL".to_string())?;
            segments.push(field);
            segments.push(value.trim());
        }
        let limit_value = limit.to_string();
        let offset_value = offset.to_string();
        self.fetch_books_from_url(
            url,
            &[
                ("format", "json"),
                ("extended", "1"),
                ("coverart", "1"),
                ("limit", limit_value.as_str()),
                ("offset", offset_value.as_str()),
            ],
        )
    }

    fn fetch_books_with_extra_query(
        &self,
        limit: usize,
        offset: usize,
        extra: &[(&str, &str)],
    ) -> Result<Vec<LibrivoxBook>, String> {
        let url = Url::parse(API_BASE).map_err(|error| error.to_string())?;
        let limit_value = limit.to_string();
        let offset_value = offset.to_string();
        let mut query = vec![
            ("format", "json"),
            ("extended", "1"),
            ("coverart", "1"),
            ("limit", limit_value.as_str()),
            ("offset", offset_value.as_str()),
        ];
        query.extend_from_slice(extra);
        self.fetch_books_from_url(url, &query)
    }

    fn fetch_books_from_url(
        &self,
        url: Url,
        query: &[(&str, &str)],
    ) -> Result<Vec<LibrivoxBook>, String> {
        let response = self
            .client
            .get(url)
            .query(query)
            .header("Accept", "application/json")
            .send()
            .map_err(|error| format!("LibriVox network error: {error}"))?;
        if response.status().as_u16() == 404 {
            return Ok(Vec::new());
        }
        let response = response
            .error_for_status()
            .map_err(|error| format!("LibriVox server error: {error}"))?;
        let root = response
            .json::<Value>()
            .map_err(|error| format!("Invalid LibriVox response: {error}"))?;
        Ok(root
            .get("books")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
            .filter_map(parse_book)
            .collect())
    }
}

pub fn open(parent: HWND) {
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let Some(mut query) = prompt_window::prompt_user(
        parent,
        &i18n::tr(language, "librivox.title"),
        &i18n::tr(language, "librivox.search_prompt"),
        "",
        language,
    ) else {
        return;
    };
    query = query.trim().to_string();
    let client = LibrivoxClient::default();

    loop {
        crate::screen_reader_speak(&i18n::tr(language, "librivox.loading"));
        let first_page = match client.search(&query, 0, PAGE_LIMIT) {
            Ok(page) => page,
            Err(error) => {
                show_librivox_error(parent, language, &error);
                return;
            }
        };
        let mut books = first_page.books;
        let mut has_more = first_page.has_more;
        let mut offset = books.len();
        let mut selected_id = None;

        if books.is_empty() {
            show_error(parent, language, &i18n::tr(language, "librivox.no_results"));
            let Some(new_query) = prompt_window::prompt_user(
                parent,
                &i18n::tr(language, "librivox.title"),
                &i18n::tr(language, "librivox.search_prompt"),
                &query,
                language,
            ) else {
                return;
            };
            query = new_query.trim().to_string();
            continue;
        }

        loop {
            let list = books
                .iter()
                .map(|book| MultilineSelectionItem {
                    id: book.id.to_string(),
                    title: book.title.clone(),
                    description: Some(book_description(book, language)),
                })
                .collect();
            let result = youtube_transcript_window::select_multiline_items_with_search(
                parent,
                language,
                i18n::tr(language, "librivox.title"),
                list,
                selected_id.clone(),
                MultilineSearchOptions {
                    initial_query: query.clone(),
                    search_button_label: i18n::tr(language, "podcasts.search.button"),
                    show_search_edit: true,
                    secondary_action_label: has_more
                        .then(|| i18n::tr(language, "podcasts.categories.load_more_results")),
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
                    break;
                }
                MultilineSelectionResult::SecondaryAction => {
                    if !has_more {
                        continue;
                    }
                    crate::screen_reader_speak(&i18n::tr(language, "librivox.loading"));
                    match client.search(&query, offset, PAGE_LIMIT) {
                        Ok(mut page) => {
                            offset = offset.saturating_add(page.books.len());
                            books.append(&mut page.books);
                            has_more = page.has_more;
                        }
                        Err(error) => show_librivox_error(parent, language, &error),
                    }
                }
                MultilineSelectionResult::Selected(id) => {
                    selected_id = Some(id.clone());
                    if let Some(book) = books.iter().find(|book| book.id.to_string() == id) {
                        let parent_list = LibrivoxParentListContext {
                            query: query.clone(),
                            books: books.clone(),
                            has_more,
                            offset,
                            selected_id: id,
                        };
                        if browse_book(parent, language, &client, book, parent_list) {
                            return;
                        }
                    }
                }
            }
        }
    }
}

fn browse_book(
    parent: HWND,
    language: Language,
    client: &LibrivoxClient,
    summary: &LibrivoxBook,
    parent_list: LibrivoxParentListContext,
) -> bool {
    crate::screen_reader_speak(&i18n::tr(language, "librivox.loading_tracks"));
    let book = if summary.sections.is_empty() {
        match client.fetch_book(summary.id) {
            Ok(book) => book,
            Err(error) => {
                show_librivox_error(parent, language, &error);
                return false;
            }
        }
    } else {
        summary.clone()
    };
    if book.sections.is_empty() {
        show_error(parent, language, &i18n::tr(language, "librivox.no_tracks"));
        return false;
    }
    browse_loaded_book(parent, language, book, None, parent_list)
}

fn browse_loaded_book(
    parent: HWND,
    language: Language,
    book: LibrivoxBook,
    mut selected_id: Option<String>,
    parent_list: LibrivoxParentListContext,
) -> bool {
    loop {
        let list = book
            .sections
            .iter()
            .map(|track| MultilineSelectionItem {
                id: track_identity(track),
                title: track_title(track),
                description: (!track.play_time.trim().is_empty()).then(|| track.play_time.clone()),
            })
            .collect();
        match youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            book.title.clone(),
            list,
            selected_id.clone(),
            MultilineSearchOptions {
                initial_query: String::new(),
                search_button_label: String::new(),
                show_search_edit: false,
                secondary_action_label: None,
                context_actions: Vec::new(),
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: true,
                refresh: None,
            },
        ) {
            MultilineSelectionResult::Cancelled => return false,
            MultilineSelectionResult::Selected(id) => {
                selected_id = Some(id.clone());
                if let Some(track) = book
                    .sections
                    .iter()
                    .find(|track| track_identity(track) == id)
                {
                    remember_player_return_context(LibrivoxPlayerReturnContext {
                        book: book.clone(),
                        selected_id: id,
                        selected_url: track.listen_url.clone(),
                        parent_list: parent_list.clone(),
                    });
                    crate::play_named_remote_audio_from_url_with_rai_origin(
                        parent,
                        track.listen_url.clone(),
                        Some(format!("{} - {}", book.title, track_title(track))),
                        Some(audio_mime(&track.listen_url)),
                        RaiAudioOrigin::None,
                    );
                    return true;
                }
            }
            MultilineSelectionResult::Search(_) | MultilineSelectionResult::SecondaryAction => {}
        }
    }
}

fn browse_parent_list(parent: HWND, language: Language, mut context: LibrivoxParentListContext) {
    let client = LibrivoxClient::default();
    loop {
        let list = context
            .books
            .iter()
            .map(|book| MultilineSelectionItem {
                id: book.id.to_string(),
                title: book.title.clone(),
                description: Some(book_description(book, language)),
            })
            .collect();
        match youtube_transcript_window::select_multiline_items_with_search(
            parent,
            language,
            i18n::tr(language, "librivox.title"),
            list,
            Some(context.selected_id.clone()),
            MultilineSearchOptions {
                initial_query: context.query.clone(),
                search_button_label: i18n::tr(language, "podcasts.search.button"),
                show_search_edit: true,
                secondary_action_label: context
                    .has_more
                    .then(|| i18n::tr(language, "podcasts.categories.load_more_results")),
                context_actions: Vec::new(),
                right_arrow_accepts_selection: true,
                left_arrow_closes: true,
                escape_stops_active_player: true,
                refresh: None,
            },
        ) {
            MultilineSelectionResult::Cancelled => return,
            MultilineSelectionResult::Search(query) => {
                context.query = query.trim().to_string();
                crate::screen_reader_speak(&i18n::tr(language, "librivox.loading"));
                match client.search(&context.query, 0, PAGE_LIMIT) {
                    Ok(page) if !page.books.is_empty() => {
                        context.books = page.books;
                        context.has_more = page.has_more;
                        context.offset = context.books.len();
                        context.selected_id = context.books[0].id.to_string();
                    }
                    Ok(_) => {
                        show_error(parent, language, &i18n::tr(language, "librivox.no_results"))
                    }
                    Err(error) => show_librivox_error(parent, language, &error),
                }
            }
            MultilineSelectionResult::SecondaryAction => {
                if !context.has_more {
                    continue;
                }
                crate::screen_reader_speak(&i18n::tr(language, "librivox.loading"));
                match client.search(&context.query, context.offset, PAGE_LIMIT) {
                    Ok(mut page) => {
                        context.offset = context.offset.saturating_add(page.books.len());
                        context.books.append(&mut page.books);
                        context.has_more = page.has_more;
                    }
                    Err(error) => show_librivox_error(parent, language, &error),
                }
            }
            MultilineSelectionResult::Selected(id) => {
                context.selected_id = id.clone();
                if let Some(book) = context
                    .books
                    .iter()
                    .find(|book| book.id.to_string() == id)
                    .cloned()
                    && browse_book(parent, language, &client, &book, context.clone())
                {
                    return;
                }
            }
        }
    }
}

fn parse_book(value: &Value) -> Option<LibrivoxBook> {
    let id = value_as_i64(value.get("id"));
    if id <= 0 {
        return None;
    }
    let title = value_as_string(value.get("title"));
    let authors = value
        .get("authors")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(|author| {
            let first = value_as_string(author.get("first_name"));
            let last = value_as_string(author.get("last_name"));
            let name = format!("{first} {last}").trim().to_string();
            (!name.is_empty()).then_some(name)
        })
        .collect();
    let sections = value
        .get("sections")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(parse_track)
        .collect();
    Some(LibrivoxBook {
        id,
        title: if title.is_empty() {
            "LibriVox".to_string()
        } else {
            title
        },
        description: value_as_string(value.get("description")),
        language: value_as_string(value.get("language")),
        total_time: value_as_string(value.get("totaltime")),
        authors,
        sections,
    })
}

fn parse_track(value: &Value) -> Option<LibrivoxTrack> {
    let listen_url = value_as_string(value.get("listen_url"));
    if listen_url.is_empty() {
        return None;
    }
    Some(LibrivoxTrack {
        id: value_as_i64(value.get("id")),
        number: value_as_i64(value.get("section_number")),
        title: {
            let title = value_as_string(value.get("title"));
            if title.is_empty() {
                "Chapter".to_string()
            } else {
                title
            }
        },
        listen_url,
        play_time: value_as_string(value.get("playtime")),
    })
}

fn book_description(book: &LibrivoxBook, language: Language) -> String {
    let authors = if book.authors.is_empty() {
        i18n::tr(language, "librivox.unknown_author")
    } else {
        book.authors.join(", ")
    };
    let language_value = (!book.language.trim().is_empty()).then(|| {
        format!(
            "{} {}",
            i18n::tr(language, "librivox.language_value"),
            book.language
        )
    });
    let duration = (!book.total_time.trim().is_empty()).then(|| {
        format!(
            "{} {}",
            i18n::tr(language, "librivox.duration_value"),
            book.total_time
        )
    });
    [
        Some(authors),
        language_value,
        duration,
        (!book.description.trim().is_empty()).then(|| book.description.clone()),
    ]
    .into_iter()
    .flatten()
    .collect::<Vec<_>>()
    .join("\n\n")
}

fn sorted_matches<'a>(
    books: impl Iterator<Item = &'a LibrivoxBook>,
    terms: &[String],
) -> Vec<LibrivoxBook> {
    let mut matches = books
        .filter(|book| matches_all_terms(book, terms))
        .cloned()
        .collect::<Vec<_>>();
    matches.sort_by(|left, right| {
        match_score(right, terms)
            .cmp(&match_score(left, terms))
            .then_with(|| left.title.to_lowercase().cmp(&right.title.to_lowercase()))
            .then_with(|| left.id.cmp(&right.id))
    });
    matches
}

fn search_terms(query: &str) -> Vec<String> {
    let ignored: HashSet<&str> = [
        "a", "an", "and", "de", "del", "della", "di", "e", "el", "gli", "i", "il", "in", "la",
        "le", "les", "of", "on", "the", "un", "una", "und",
    ]
    .into_iter()
    .collect();
    normalize_for_search(query)
        .split_whitespace()
        .filter(|term| term.len() > 1 && !ignored.contains(*term))
        .map(str::to_string)
        .collect()
}

fn matches_all_terms(book: &LibrivoxBook, terms: &[String]) -> bool {
    if terms.is_empty() {
        return true;
    }
    let haystack = normalize_for_search(&format!(
        "{} {} {} {}",
        book.title,
        book.authors.join(" "),
        book.description,
        book.language
    ));
    let matched = terms
        .iter()
        .filter(|term| haystack.contains(term.as_str()))
        .count();
    matched == terms.len() || (terms.len() >= 3 && matched + 1 >= terms.len())
}

fn match_score(book: &LibrivoxBook, terms: &[String]) -> usize {
    let title = normalize_for_search(&book.title);
    let authors = normalize_for_search(&book.authors.join(" "));
    let description = normalize_for_search(&book.description);
    terms
        .iter()
        .map(|term| {
            usize::from(title.contains(term.as_str())) * 4
                + usize::from(authors.contains(term.as_str())) * 3
                + usize::from(description.contains(term.as_str()))
        })
        .sum()
}

fn normalize_for_search(value: &str) -> String {
    value
        .to_lowercase()
        .replace('&', " and ")
        .chars()
        .map(|character| match character {
            'à' | 'á' | 'â' | 'ã' | 'ä' | 'å' | 'ā' | 'ă' | 'ą' => 'a',
            'ç' | 'ć' | 'ĉ' | 'ċ' | 'č' => 'c',
            'ď' | 'đ' => 'd',
            'è' | 'é' | 'ê' | 'ë' | 'ē' | 'ĕ' | 'ė' | 'ę' | 'ě' => 'e',
            'ì' | 'í' | 'î' | 'ï' | 'ĩ' | 'ī' | 'ĭ' | 'į' | 'ı' => 'i',
            'ñ' | 'ń' | 'ņ' | 'ň' => 'n',
            'ò' | 'ó' | 'ô' | 'õ' | 'ö' | 'ø' | 'ō' | 'ŏ' | 'ő' => 'o',
            'ŕ' | 'ŗ' | 'ř' => 'r',
            'ś' | 'ŝ' | 'ş' | 'š' => 's',
            'ù' | 'ú' | 'û' | 'ü' | 'ũ' | 'ū' | 'ŭ' | 'ů' | 'ű' | 'ų' => 'u',
            'ý' | 'ÿ' | 'ŷ' => 'y',
            'ź' | 'ż' | 'ž' => 'z',
            character if character.is_ascii_alphanumeric() => character,
            _ => ' ',
        })
        .collect::<String>()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

fn track_identity(track: &LibrivoxTrack) -> String {
    format!("{}:{}", track.id, track.listen_url)
}

fn track_title(track: &LibrivoxTrack) -> String {
    if track.number > 0 {
        format!("{}. {}", track.number, track.title)
    } else {
        track.title.clone()
    }
}

fn audio_mime(url: &str) -> &'static str {
    let lower = url.to_lowercase();
    if lower.contains(".ogg") {
        "audio/ogg"
    } else if lower.contains(".m4a") || lower.contains(".mp4") {
        "audio/mp4"
    } else {
        "audio/mpeg"
    }
}

fn value_as_string(value: Option<&Value>) -> String {
    match value {
        Some(Value::String(text)) => text.trim().to_string(),
        Some(Value::Number(number)) => number.to_string(),
        _ => String::new(),
    }
}

fn value_as_i64(value: Option<&Value>) -> i64 {
    match value {
        Some(Value::Number(number)) => number.as_i64().unwrap_or(0),
        Some(Value::String(text)) => text.parse().unwrap_or(0),
        _ => 0,
    }
}

fn show_librivox_error(parent: HWND, language: Language, error: &str) {
    crate::log_debug(&format!("LibriVox error: {error}"));
    let message = i18n::tr_f(language, "librivox.error", &[("err", error)]);
    show_error(parent, language, &message);
}
