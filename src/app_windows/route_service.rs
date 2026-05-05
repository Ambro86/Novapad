use reqwest::blocking::Client;
use serde::Deserialize;
use std::fmt;

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

    pub fn label_it(self) -> &'static str {
        match self {
            RouteProfile::Walking => "a piedi",
            RouteProfile::Cycling => "in bici",
            RouteProfile::Driving => "in auto",
            RouteProfile::Wheelchair => "in sedia a rotelle",
        }
    }
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

    pub fn geocode(&self, address: &str) -> Result<Vec<GeocodeCandidate>, RouteError> {
        let address = address.trim();

        if address.is_empty() {
            return Err(RouteError::InvalidInput(
                "Inserisci un indirizzo.".to_string(),
            ));
        }

        // Prova la query originale
        let results = self.fetch_geocode(address)?;

        // Se i risultati sono solo la città (fallback del geocoder), prova a semplificare la query
        if is_all_city_fallback(&results, address) {
            if let Some(simplified) = simplify_query(address)
                && let Ok(fallback_results) = self.fetch_geocode(&simplified)
                && !fallback_results.is_empty()
                && !is_all_city_fallback(&fallback_results, &simplified)
            {
                return Ok(fallback_results);
            }

            // Prova un secondo livello di semplificazione (rimuovendo un'altra parola)
            if let Some(more_simplified) = simplify_query_more(address)
                && let Ok(fallback_results) = self.fetch_geocode(&more_simplified)
                && !fallback_results.is_empty()
                && !is_all_city_fallback(&fallback_results, &more_simplified)
            {
                return Ok(fallback_results);
            }
        }

        Ok(results)
    }

    fn fetch_geocode(&self, q: &str) -> Result<Vec<GeocodeCandidate>, RouteError> {
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
            .map_err(|error| RouteError::Network(error.to_string()))?
            .json()
            .map_err(|error| RouteError::InvalidResponse(error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(response.error.unwrap_or_else(|| {
                "Errore durante la ricerca dell’indirizzo.".to_string()
            })));
        }

        Ok(response.results)
    }

    pub fn route_between_coordinates(
        &self,
        from: &GeocodeCandidate,
        to: &GeocodeCandidate,
        profile: RouteProfile,
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
                ("profile", profile.api_value().to_string()),
            ])
            .send()
            .map_err(|error| RouteError::Network(error.to_string()))?
            .json()
            .map_err(|error| RouteError::InvalidResponse(error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(response.error.unwrap_or_else(|| {
                "Errore durante il calcolo del percorso.".to_string()
            })));
        }

        Ok(RouteResult {
            profile,
            from_label: from.label.clone(),
            to_label: to.label.clone(),
            distance_meters: response.distance_meters.unwrap_or(0.0),
            duration_seconds: response.duration_seconds.unwrap_or(0.0),
            steps: response.steps,
        })
    }

    pub fn route_from_addresses(
        &self,
        from_address: &str,
        to_address: &str,
        profile: RouteProfile,
    ) -> Result<RouteRequestResult, RouteError> {
        let from_candidates = self.geocode(from_address)?;
        let to_candidates = self.geocode(to_address)?;

        if from_candidates.is_empty() {
            return Err(RouteError::NotFound(
                "Nessun risultato trovato per l’indirizzo di partenza.".to_string(),
            ));
        }

        if to_candidates.is_empty() {
            return Err(RouteError::NotFound(
                "Nessun risultato trovato per l’indirizzo di arrivo.".to_string(),
            ));
        }

        if from_candidates.len() > 1 || to_candidates.len() > 1 {
            return Ok(RouteRequestResult::NeedsSelection {
                from_candidates,
                to_candidates,
                profile,
            });
        }

        let route =
            self.route_between_coordinates(&from_candidates[0], &to_candidates[0], profile)?;

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
    pub from_label: String,
    pub to_label: String,
    pub distance_meters: f64,
    pub duration_seconds: f64,
    pub steps: Vec<RouteStep>,
}

impl RouteResult {
    pub fn format_for_speech_or_text(&self) -> String {
        let mut output = String::new();

        output.push_str("Percorso ");
        output.push_str(self.profile.label_it());
        output.push_str(" trovato.\n\n");

        output.push_str("Partenza: ");
        output.push_str(&self.from_label);
        output.push('\n');

        output.push_str("Arrivo: ");
        output.push_str(&self.to_label);
        output.push_str("\n\n");

        output.push_str("Distanza: ");
        output.push_str(&format_distance(self.distance_meters));
        output.push_str(".\n");

        output.push_str("Durata stimata: ");
        output.push_str(&format_duration(self.duration_seconds));
        output.push_str(".\n\n");

        output.push_str("Indicazioni:\n");

        if self.steps.is_empty() {
            output.push_str("Nessuna istruzione disponibile.");
            return output;
        }

        for (index, step) in self.steps.iter().enumerate() {
            let instruction = clean_route_instruction(&step.instruction);
            let distance = format_distance(step.distance_meters);

            output.push_str(&(index + 1).to_string());
            output.push_str(". ");
            output.push_str(&instruction);

            if step.distance_meters > 0.0 {
                output.push_str(", per ");
                output.push_str(&distance);
            }

            output.push_str(".\n");
        }

        output
    }
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
            RouteError::Network(message) => write!(f, "Errore di rete: {message}"),
            RouteError::InvalidResponse(message) => {
                write!(f, "Risposta non valida dal server: {message}")
            }
            RouteError::Server(message) => write!(f, "{message}"),
            RouteError::NotFound(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for RouteError {}

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

pub fn format_distance(meters: f64) -> String {
    if meters < 1.0 {
        return "pochi metri".to_string();
    }

    if meters < 1000.0 {
        return format!("{} metri", meters.round() as i64);
    }

    let kilometers = meters / 1000.0;

    if kilometers < 10.0 {
        return format!("{kilometers:.1} km").replace('.', ",");
    }

    format!("{} km", kilometers.round() as i64)
}

pub fn format_duration(seconds: f64) -> String {
    let total_seconds = seconds.round() as i64;

    if total_seconds < 60 {
        return format!("{total_seconds} secondi");
    }

    let total_minutes = ((total_seconds as f64) / 60.0).round() as i64;

    if total_minutes < 60 {
        if total_minutes == 1 {
            return "1 minuto".to_string();
        }

        return format!("{total_minutes} minuti");
    }

    let hours = total_minutes / 60;
    let minutes = total_minutes % 60;

    if minutes == 0 {
        if hours == 1 {
            return "1 ora".to_string();
        }

        return format!("{hours} ore");
    }

    if hours == 1 {
        return format!("1 ora e {minutes} minuti");
    }

    format!("{hours} ore e {minutes} minuti")
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
