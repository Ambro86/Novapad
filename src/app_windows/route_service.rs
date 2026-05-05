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

        let url = format!("{}/ors_geocode.php", self.base_url);

        let response: GeocodeApiResponse = self
            .client
            .get(url)
            .query(&[("q", address)])
            .send()
            .map_err(|error| RouteError::Network(error.to_string()))?
            .json()
            .map_err(|error| RouteError::InvalidResponse(error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(
                response
                    .error
                    .unwrap_or_else(|| "Errore durante la ricerca dell’indirizzo.".to_string()),
            ));
        }

        Ok(response.results)
    }

    pub fn route_between_coordinates(
        &self,
        from: &GeocodeCandidate,
        to: &GeocodeCandidate,
        profile: RouteProfile,
    ) -> Result<RouteResult, RouteError> {
        let from_lat = from.latitude.ok_or_else(|| {
            RouteError::InvalidInput("La partenza non contiene la latitudine.".to_string())
        })?;

        let from_lon = from.longitude.ok_or_else(|| {
            RouteError::InvalidInput("La partenza non contiene la longitudine.".to_string())
        })?;

        let to_lat = to.latitude.ok_or_else(|| {
            RouteError::InvalidInput("La destinazione non contiene la latitudine.".to_string())
        })?;

        let to_lon = to.longitude.ok_or_else(|| {
            RouteError::InvalidInput("La destinazione non contiene la longitudine.".to_string())
        })?;

        let url = format!("{}/ors_route.php", self.base_url);

        let response: RouteApiResponse = self
            .client
            .get(url)
            .query(&[
                ("from_lat", from_lat.to_string()),
                ("from_lon", from_lon.to_string()),
                ("to_lat", to_lat.to_string()),
                ("to_lon", to_lon.to_string()),
                ("profile", profile.api_value().to_string()),
            ])
            .send()
            .map_err(|error| RouteError::Network(error.to_string()))?
            .json()
            .map_err(|error| RouteError::InvalidResponse(error.to_string()))?;

        if !response.ok {
            return Err(RouteError::Server(
                response
                    .error
                    .unwrap_or_else(|| "Errore durante il calcolo del percorso.".to_string()),
            ));
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

    /*
     * Questa funzione fa tutto:
     * 1. geocoding partenza
     * 2. geocoding arrivo
     * 3. se ci sono più risultati, restituisce NeedsSelection
     * 4. se c’è un solo risultato per parte, calcola direttamente il percorso
     */
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

    /*
     * Se ci sono più indirizzi possibili, la UI deve mostrare queste liste.
     * L’utente sceglie una partenza e un arrivo.
     * Poi richiami route_between_coordinates().
     */
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
    pub region: String,
    pub county: String,
    pub locality: String,
    pub postalcode: String,
    pub longitude: Option<f64>,
    pub latitude: Option<f64>,
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
    pub duration_seconds: f64,
}

#[derive(Debug, Deserialize)]
struct GeocodeApiResponse {
    ok: bool,
    error: Option<String>,
    results: Vec<GeocodeCandidate>,
}

#[derive(Debug, Deserialize)]
struct RouteApiResponse {
    ok: bool,
    error: Option<String>,
    distance_meters: Option<f64>,
    duration_seconds: Option<f64>,
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