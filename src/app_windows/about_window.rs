use windows::core::PCWSTR;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{MessageBoxW, MB_OK, MB_ICONINFORMATION};
use crate::with_state;
use crate::settings::Language;
use crate::accessibility::to_wide;

pub unsafe fn show(hwnd: HWND) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let message = to_wide(about_message(language));
    let title = to_wide(about_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
}

fn about_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Informazioni sul programma",
        Language::English => "About the program",
    }
}

fn about_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Questo programma è un piccolo Notepad, creato da Ambrogio Riili, che permette di aprire i files più comuni, tra cui anche pdf, e di creare degli audiolibri.",
        Language::English => "This program is a small Notepad, created by Ambrogio Riili, that can open common files, including PDF, and can create audiobooks.",
    }
}
