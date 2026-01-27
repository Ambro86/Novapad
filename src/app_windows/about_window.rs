use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::Language;
use crate::with_state;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{MB_ICONINFORMATION, MB_OK, MessageBoxW};
use windows::core::PCWSTR;

pub unsafe fn show(hwnd: HWND) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let message = to_wide(&about_message(language));
    let title = to_wide(&about_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
}

fn about_title(language: Language) -> String {
    i18n::tr(language, "about.title")
}

fn about_message(language: Language) -> String {
    let version = env!("CARGO_PKG_VERSION");
    i18n::tr_f(language, "about.message", &[("version", version)])
}
