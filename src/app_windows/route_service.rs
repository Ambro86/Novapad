use reqwest::blocking::Client;
use serde::{Deserialize, Serialize};
use std::fmt;

use crate::i18n;
use crate::settings::Language;

const DEFAULT_BASE_URL: &str = "https://sonarpad.com/api";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RouteProfile {
    Walking,
    Cycling,
    Driving,
    Wheelchair,
}

impl RouteProfile {
    pub fn api_value(self) -> &'static str {
        match self {
            RouteProfile::Walking => "foot-walking",
            RouteProfile::Cycling => "cycling-regular",
            RouteProfile::Driving => "driving-car",
            RouteProfile::Wheelchair => "wheelchair",
        }
    }

    pub fn label(self, language: Language) -> String {
        let key = match self {
            RouteProfile::Walking => "route.profile.walking",
            RouteProfile::Cycling => "route.profile.cycling",
            RouteProfile::Driving => "route.profile.driving",
            RouteProfile::Wheelchair => "route.profile.wheelchair",
        };
        i18n::tr(language, key)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RoutePreference {
    Fastest,
    Shortest,
}

impl RoutePreference {
    pub fn api_value(self) -> &'static str {
        match self {
            RoutePreference::Fastest => "fastest",
            RoutePreference::Shortest => "shortest",
        }
    }

    pub fn label(self, language: Language) -> String {
        let key = match self {
            RoutePreference::Fastest => "route.preference.fastest",
            RoutePreference::Shortest => "route.preference.shortest",
        };
        i18n::tr(language, key)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RouteAvoid {
    None,
    Highways,
    Tollways,
    HighwaysAndTollways,
}

impl RouteAvoid {
    pub fn api_value(self) -> &'static str {
        match self {
            RouteAvoid::None => "",
            RouteAvoid::Highways => "highways",
            RouteAvoid::Tollways => "tollways",
            RouteAvoid::HighwaysAndTollways => "highways,tollways",
        }
    }

    pub fn label(self, language: Language) -> String {
        let key = match self {
            RouteAvoid::None => "route.avoid.none",
            RouteAvoid::Highways => "route.avoid.highways",
            RouteAvoid::Tollways => "route.avoid.tollways",
            RouteAvoid::HighwaysAndTollways => "route.avoid.highways_tollways",
        };
        i18n::tr(language, key)
    }
}

#[derive(Clone, Copy)]
pub struct RouteOptions {
    pub profile: RouteProfile,
    pub preference: RoutePreference,
    pub avoid: RouteAvoid,
    pub include_municipalities: bool,
    pub language: Language,
}

#[derive(Debug, Clone)]
pub struct RouteClient {
    client: Client,
    base_url: String,
}

impl Default for RouteClient {
    fn default() -> Self {
        Self::new(DEFAULT_BASE_URL)
    }
}

impl RouteClient {
    pub fn new(base_url: &str) -> Self {
        Self {
            client: Client::new(),
            base_url: base_url.trim_end_matches('/').to_string(),
        }
    }

    pub fn geocode(
        &self,
        address: &str,
        language: Language,
    ) -> Result<Vec<GeocodeCandidate>, RouteError> {
        let address = address.trim();

        if address.is_empty() {
            return Err(RouteError::InvalidInput(i18n::tr(
                language,
                "route.error.address_search",
            )));
        }

        // Prova la query originale
        let results = self.fetch_geocode(address, language)?;

        // Se i risultati sono solo la città (fallback del geocoder), prova a semplificare la query
        if is_all_city_fallback(&results, address) {
            if let Some(simplified) = simplify_query(address)
                && let Ok(fallback_results) = self.fetch_geocode(&simplified, language)
                && !fallback_results.is_empty()
                && !is_all_city_fallback(&fallback_results, &simplified)
            {
                return Ok(fallback_results);
            }

            // Prova un secondo livello di semplificazione (rimuovendo un'altra parola)
            if let Some(more_simplified) = simplify_query_more(address)
                && let Ok(fallback_results) = self.fetch_geocode(&more_simplified, language)
                && !fallback_results.is_empty()
                && !is_all_city_fallback(&fallback_results, &more_simplified)
            {
                return Ok(fallback_results);
            }
        }

        Ok(results)
    }

    fn fetch_geocode(
        &self,
        q: &str,
        language: Language,
    ) -> Result<Vec<GeocodeCandidate>, RouteError> {
        let url = format!("{}/ors_geocode.php", self.base_url);

        let response: GeocodeApiResponse = self
            .client
            .get(url)
            .query(&[
                ("q", q),
                ("size", "20"),
                ("layers", "address,street,venue"),
                ("sources", "osm,oa"),
                ("boundary.country", "ITA"),
            ])
            .send()
            .map_err(|error| route_network_error(language, error.to_string()))?
            .json()
            .map_err(|error| route_invalid_response_error(language, error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(response.error.unwrap_or_else(|| {
                i18n::tr(language, "route.error.address_search")
            })));
        }

        Ok(response.results)
    }

    pub fn route_between_coordinates(
        &self,
        from: &GeocodeCandidate,
        to: &GeocodeCandidate,
        options: RouteOptions,
    ) -> Result<RouteResult, RouteError> {
        let url = format!("{}/ors_route.php", self.base_url);

        let response: RouteApiResponse = self
            .client
            .get(url)
            .query(&[
                ("from_lat", from.latitude.to_string()),
                ("from_lon", from.longitude.to_string()),
                ("to_lat", to.latitude.to_string()),
                ("to_lon", to.longitude.to_string()),
                ("profile", options.profile.api_value().to_string()),
                ("preference", options.preference.api_value().to_string()),
                ("avoid", options.avoid.api_value().to_string()),
                (
                    "include_municipalities",
                    if options.include_municipalities {
                        "1"
                    } else {
                        "0"
                    }
                    .to_string(),
                ),
                (
                    "language",
                    route_backend_language(options.language).to_string(),
                ),
            ])
            .send()
            .map_err(|error| route_network_error(options.language, error.to_string()))?
            .json()
            .map_err(|error| route_invalid_response_error(options.language, error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(response.error.unwrap_or_else(|| {
                i18n::tr(options.language, "route.error.calculation")
            })));
        }

        Ok(RouteResult {
            profile: options.profile,
            preference: options.preference,
            avoid: options.avoid,
            from_label: from.label.clone(),
            to_label: to.label.clone(),
            paths: response.into_paths(),
        })
    }

    pub fn route_from_addresses(
        &self,
        from_address: &str,
        to_address: &str,
        options: RouteOptions,
    ) -> Result<RouteRequestResult, RouteError> {
        let from_candidates = self.geocode(from_address, options.language)?;
        let to_candidates = self.geocode(to_address, options.language)?;

        if from_candidates.is_empty() {
            return Err(RouteError::NotFound(i18n::tr(
                options.language,
                "route.error.from_not_found",
            )));
        }

        if to_candidates.is_empty() {
            return Err(RouteError::NotFound(i18n::tr(
                options.language,
                "route.error.to_not_found",
            )));
        }

        if from_candidates.len() > 1 || to_candidates.len() > 1 {
            return Ok(RouteRequestResult::NeedsSelection {
                from_candidates,
                to_candidates,
                profile: options.profile,
                preference: options.preference,
                avoid: options.avoid,
                include_municipalities: options.include_municipalities,
            });
        }

        let route =
            self.route_between_coordinates(&from_candidates[0], &to_candidates[0], options)?;

        Ok(RouteRequestResult::Ready(route))
    }
}

#[derive(Debug, Clone)]
pub enum RouteRequestResult {
    Ready(RouteResult),
    NeedsSelection {
        from_candidates: Vec<GeocodeCandidate>,
        to_candidates: Vec<GeocodeCandidate>,
        profile: RouteProfile,
        preference: RoutePreference,
        avoid: RouteAvoid,
        include_municipalities: bool,
    },
}

#[derive(Debug, Clone, Deserialize)]
pub struct GeocodeCandidate {
    pub label: String,
    pub name: String,
    pub country: String,
    #[allow(dead_code)]
    pub region: String,
    #[allow(dead_code)]
    pub county: String,
    pub locality: String,
    #[allow(dead_code)]
    pub postalcode: String,
    pub latitude: f64,
    pub longitude: f64,
}

impl GeocodeCandidate {
    pub fn display_label(&self) -> String {
        if !self.label.trim().is_empty() {
            return self.label.clone();
        }

        let mut parts = Vec::new();

        if !self.name.trim().is_empty() {
            parts.push(self.name.as_str());
        }

        if !self.locality.trim().is_empty() {
            parts.push(self.locality.as_str());
        }

        if !self.country.trim().is_empty() {
            parts.push(self.country.as_str());
        }

        parts.join(", ")
    }
}

#[derive(Debug, Clone)]
pub struct RouteResult {
    pub profile: RouteProfile,
    pub preference: RoutePreference,
    pub avoid: RouteAvoid,
    pub from_label: String,
    pub to_label: String,
    pub paths: Vec<RoutePath>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct RoutePath {
    pub distance_meters: f64,
    pub duration_seconds: f64,
    #[serde(default)]
    pub steps: Vec<RouteStep>,
    #[serde(default)]
    pub geometry: Vec<[f64; 2]>,
    #[serde(default)]
    pub municipality_changes: Vec<MunicipalityChange>,
}

#[derive(Debug, Clone, Serialize)]
pub struct RouteMapData {
    pub from_label: String,
    pub to_label: String,
    pub geometry: Vec<[f64; 2]>,
}

impl RouteResult {
    pub fn map_data(&self) -> Option<RouteMapData> {
        let path = self.paths.first()?;
        if path.geometry.len() < 2 {
            return None;
        }

        Some(RouteMapData {
            from_label: self.from_label.clone(),
            to_label: self.to_label.clone(),
            geometry: path.geometry.clone(),
        })
    }

    pub fn format_for_speech_or_text(&self, language: Language) -> String {
        let mut output = String::new();

        output.push_str(&i18n::tr(language, "route.found_prefix"));
        output.push(' ');
        output.push_str(&self.profile.label(language));
        output.push(' ');
        output.push_str(&self.preference.label(language));
        if self.avoid != RouteAvoid::None {
            output.push_str(", ");
            output.push_str(&self.avoid.label(language));
        }
        output.push(' ');
        output.push_str(&i18n::tr(language, "route.found_suffix"));
        output.push_str("\n\n");

        output.push_str(&i18n::tr(language, "route.from_result"));
        output.push(' ');
        output.push_str(&self.from_label);
        output.push('\n');

        output.push_str(&i18n::tr(language, "route.to_result"));
        output.push(' ');
        output.push_str(&self.to_label);
        output.push('\n');

        if self.paths.is_empty() {
            output.push('\n');
            output.push_str(&i18n::tr(language, "route.no_route_available"));
            return output;
        }

        if self.paths.len() > 1 {
            output.push('\n');
            output.push_str(&i18n::tr(language, "route.alternative_routes_found"));
            output.push(' ');
            output.push_str(&self.paths.len().to_string());
            output.push_str(".\n");
        }

        for (path_index, path) in self.paths.iter().enumerate() {
            output.push('\n');
            if path_index == 0 {
                output.push_str(&i18n::tr(language, "route.main_route"));
            } else {
                output.push_str(&i18n::tr(language, "route.alternative_route"));
                output.push(' ');
                output.push_str(&path_index.to_string());
            }
            output.push_str(".\n");

            output.push_str(&i18n::tr(language, "route.distance"));
            output.push(' ');
            output.push_str(&format_distance_for_language(
                path.distance_meters,
                language,
            ));
            output.push_str(".\n");

            output.push_str(&i18n::tr(language, "route.estimated_duration"));
            output.push(' ');
            output.push_str(&format_duration_for_language(
                path.duration_seconds,
                language,
            ));
            output.push_str(".\n");

            output.push('\n');
            output.push_str(&i18n::tr(language, "route.directions"));
            output.push('\n');

            append_route_steps_with_municipality_changes(&mut output, path, language);
        }

        output
    }
}

#[derive(Debug, Clone, Deserialize)]
pub struct MunicipalityChange {
    pub name: String,
    #[serde(default)]
    pub distance_meters: f64,
    #[allow(dead_code)]
    pub coordinate: Option<[f64; 2]>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct RouteStep {
    pub instruction: String,
    pub distance_meters: f64,
    #[allow(dead_code)]
    pub duration_seconds: f64,
}

#[derive(Debug, Deserialize)]
struct GeocodeApiResponse {
    ok: bool,
    error: Option<String>,
    #[serde(default)]
    results: Vec<GeocodeCandidate>,
}

#[derive(Debug, Deserialize)]
struct RouteApiResponse {
    ok: bool,
    error: Option<String>,
    distance_meters: Option<f64>,
    duration_seconds: Option<f64>,
    #[serde(default)]
    steps: Vec<RouteStep>,
    #[serde(default)]
    routes: Vec<RoutePath>,
}

impl RouteApiResponse {
    fn into_paths(self) -> Vec<RoutePath> {
        if !self.routes.is_empty() {
            return self.routes;
        }

        vec![RoutePath {
            distance_meters: self.distance_meters.unwrap_or(0.0),
            duration_seconds: self.duration_seconds.unwrap_or(0.0),
            steps: self.steps,
            geometry: Vec::new(),
            municipality_changes: Vec::new(),
        }]
    }
}

fn route_backend_language(language: Language) -> &'static str {
    match language {
        Language::Italian => "it",
        Language::English => "en",
        Language::Spanish => "es",
        Language::Portuguese => "pt",
        Language::French => "fr",
        Language::Russian => "ru",
        Language::Chinese => "zh",
        _ => "en",
    }
}

fn route_network_error(language: Language, message: String) -> RouteError {
    RouteError::Network(format!(
        "{} {}",
        i18n::tr(language, "route.error.network"),
        message
    ))
}

fn route_invalid_response_error(language: Language, message: String) -> RouteError {
    RouteError::InvalidResponse(format!(
        "{} {}",
        i18n::tr(language, "route.error.invalid_response"),
        message
    ))
}

#[derive(Debug, Clone)]
pub enum RouteError {
    InvalidInput(String),
    Network(String),
    InvalidResponse(String),
    Server(String),
    NotFound(String),
}

impl fmt::Display for RouteError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RouteError::InvalidInput(message) => write!(f, "{message}"),
            RouteError::Network(message) => write!(f, "{message}"),
            RouteError::InvalidResponse(message) => write!(f, "{message}"),
            RouteError::Server(message) => write!(f, "{message}"),
            RouteError::NotFound(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for RouteError {}

fn append_route_steps_with_municipality_changes(
    output: &mut String,
    path: &RoutePath,
    language: Language,
) {
    let changes = sorted_municipality_changes(path);

    if let Some(start_municipality) = changes
        .first()
        .filter(|change| change.distance_meters <= 1.0)
    {
        output.push_str(&i18n::tr(language, "route.start_municipality"));
        output.push(' ');
        output.push_str(start_municipality.name.trim());
        output.push_str(".\n");
    }

    let mut next_change = first_route_municipality_change_index(&changes);

    if path.steps.is_empty() {
        output.push_str(&i18n::tr(language, "route.no_instruction_available"));
        output.push('\n');
        append_remaining_municipality_changes(output, &changes, next_change, 1, language);
        return;
    }

    let mut instruction_number = 1usize;
    let mut cumulative_distance = 0.0f64;

    for step in &path.steps {
        let instruction = clean_route_instruction(&step.instruction);
        let distance = format_distance_for_language(step.distance_meters, language);

        output.push_str(&instruction_number.to_string());
        output.push_str(". ");
        output.push_str(&instruction);

        if step.distance_meters > 0.0 {
            output.push_str(", ");
            output.push_str(&i18n::tr(language, "route.for_distance"));
            output.push(' ');
            output.push_str(&distance);
        }

        output.push_str(".\n");
        instruction_number += 1;

        cumulative_distance += step.distance_meters.max(0.0);

        while next_change < changes.len()
            && changes[next_change].distance_meters <= cumulative_distance + 10.0
        {
            append_municipality_instruction(
                output,
                instruction_number,
                changes[next_change],
                language,
            );
            instruction_number += 1;
            next_change += 1;
        }
    }

    append_remaining_municipality_changes(
        output,
        &changes,
        next_change,
        instruction_number,
        language,
    );
}

fn sorted_municipality_changes(path: &RoutePath) -> Vec<&MunicipalityChange> {
    let mut changes: Vec<&MunicipalityChange> = path
        .municipality_changes
        .iter()
        .filter(|change| !change.name.trim().is_empty())
        .collect();

    changes.sort_by(|left, right| left.distance_meters.total_cmp(&right.distance_meters));

    let mut seen = std::collections::HashSet::new();
    changes
        .into_iter()
        .filter(|change| seen.insert(municipality_dedupe_key(&change.name)))
        .collect()
}

fn municipality_dedupe_key(name: &str) -> String {
    name.split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

fn first_route_municipality_change_index(changes: &[&MunicipalityChange]) -> usize {
    changes
        .iter()
        .position(|change| change.distance_meters > 1.0)
        .unwrap_or(changes.len())
}

fn append_municipality_instruction(
    output: &mut String,
    instruction_number: usize,
    change: &MunicipalityChange,
    language: Language,
) {
    output.push_str(&instruction_number.to_string());
    output.push_str(". ");
    output.push_str(&i18n::tr(language, "route.enter_municipality"));
    output.push(' ');
    output.push_str(change.name.trim());
    output.push_str(".\n");
}

fn append_remaining_municipality_changes(
    output: &mut String,
    changes: &[&MunicipalityChange],
    mut next_change: usize,
    mut instruction_number: usize,
    language: Language,
) {
    while next_change < changes.len() {
        append_municipality_instruction(output, instruction_number, changes[next_change], language);
        next_change += 1;
        instruction_number += 1;
    }
}

pub fn clean_route_instruction(instruction: &str) -> String {
    instruction
        .replace("Gira svolta sinistra", "Gira a sinistra")
        .replace("Gira svolta destra", "Gira a destra")
        .replace("Gira sulla sinistra", "Gira a sinistra")
        .replace("Gira sulla destra", "Gira a destra")
        .replace("  ", " ")
        .trim()
        .to_string()
}

pub fn format_distance_for_language(meters: f64, language: Language) -> String {
    if meters < 1.0 {
        return i18n::tr(language, "route.distance.few_meters");
    }

    if meters < 1000.0 {
        return format!(
            "{} {}",
            meters.round() as i64,
            i18n::tr(language, "route.distance.meters")
        );
    }

    let kilometers = meters / 1000.0;

    if kilometers < 10.0 {
        return format!("{kilometers:.1} km").replace('.', ",");
    }

    format!("{} km", kilometers.round() as i64)
}

pub fn format_duration_for_language(seconds: f64, language: Language) -> String {
    let total_seconds = seconds.round() as i64;

    if total_seconds < 60 {
        return format!(
            "{total_seconds} {}",
            i18n::tr(language, "route.duration.seconds")
        );
    }

    let total_minutes = ((total_seconds as f64) / 60.0).round() as i64;

    if total_minutes < 60 {
        if total_minutes == 1 {
            return i18n::tr(language, "route.duration.minute_one");
        }

        return format!(
            "{total_minutes} {}",
            i18n::tr(language, "route.duration.minutes")
        );
    }

    let hours = total_minutes / 60;
    let minutes = total_minutes % 60;

    if minutes == 0 {
        if hours == 1 {
            return i18n::tr(language, "route.duration.hour_one");
        }

        return format!("{hours} {}", i18n::tr(language, "route.duration.hours"));
    }

    if hours == 1 {
        return i18n::tr_f(
            language,
            "route.duration.hour_and_minutes",
            &[("minutes", &minutes.to_string())],
        );
    }

    i18n::tr_f(
        language,
        "route.duration.hours_and_minutes",
        &[
            ("hours", &hours.to_string()),
            ("minutes", &minutes.to_string()),
        ],
    )
}

fn is_all_city_fallback(results: &[GeocodeCandidate], original_query: &str) -> bool {
    if results.is_empty() {
        return false;
    }

    // Se la query originale è corta, non consideriamolo un fallback
    if original_query.split_whitespace().count() <= 1 {
        return false;
    }

    // Se tutti i risultati hanno postalcode vuoto e il nome è uguale alla località, è probabilmente un fallback del geocoder alla città
    results.iter().all(|c| {
        c.postalcode.is_empty()
            && (c.name == c.locality || c.name == c.region || c.name == c.country)
    })
}

fn simplify_query(query: &str) -> Option<String> {
    let mut words: Vec<&str> = query.split_whitespace().collect();
    if words.len() <= 1 {
        return None;
    }

    // Rimuovi prefissi comuni (via, corso, etc.)
    let prefixes = [
        "via", "corso", "viale", "piazza", "vicolo", "largo", "strada", "v.", "c.so", "p.zza",
        "p.za",
    ];
    if prefixes.contains(&words[0].to_lowercase().as_str()) {
        words.remove(0);
    }

    if words.is_empty() {
        return None;
    }

    Some(words.join(" "))
}

fn simplify_query_more(query: &str) -> Option<String> {
    let mut words: Vec<&str> = query.split_whitespace().collect();

    // Rimuovi il prefisso se esiste
    let prefixes = [
        "via", "corso", "viale", "piazza", "vicolo", "largo", "strada", "v.", "c.so", "p.zza",
        "p.za",
    ];
    if !words.is_empty() && prefixes.contains(&words[0].to_lowercase().as_str()) {
        words.remove(0);
    }

    // Se rimangono più di 2 parole (es. "Pierluigi Palestrina 45 Torino"), prova a rimuovere la prima parola del nome
    // Spesso i geocoder falliscono sui nomi propri composti o con refusi
    if words.len() > 2 {
        words.remove(0);
        return Some(words.join(" "));
    }

    None
}
