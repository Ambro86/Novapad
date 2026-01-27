use crate::i18n;
use crate::settings::Language;
use platform_windows::{WindowHandle, about_window};

pub fn show(parent: Option<WindowHandle>, language: Language) {
    let title = i18n::tr(language, "about.title");
    let version = env!("CARGO_PKG_VERSION");
    let message = i18n::tr_f(language, "about.message", &[("version", version)]);

    if let Err(e) = about_window::open(parent, &title, &message) {
        crate::log_debug(&format!("Error opening about window: {}", e));
    }
}
