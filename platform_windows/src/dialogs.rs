use crate::{WindowHandle, to_wide};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    IDCANCEL, IDCONTINUE, IDIGNORE, IDNO, IDOK, IDRETRY, IDTRYAGAIN, IDYES, MB_ICONERROR,
    MB_ICONINFORMATION, MB_ICONQUESTION, MB_OK, MB_SETFOREGROUND, MB_YESNO, MB_YESNOCANCEL,
    MessageBoxW,
};
use windows::core::PCWSTR;

#[derive(Debug, Clone, Copy)]
pub enum DialogStyle {
    Info,
    Error,
    Question,
    YesNo,
    YesNoCancel,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DialogResult {
    Ok,
    Cancel,
    Yes,
    No,
    Retry,
    TryAgain,
    Continue,
    Ignore,
    Unknown,
}

impl DialogResult {
    pub fn as_int(&self) -> i32 {
        match self {
            DialogResult::Ok => IDOK.0,
            DialogResult::Cancel => IDCANCEL.0,
            DialogResult::Yes => IDYES.0,
            DialogResult::No => IDNO.0,
            DialogResult::Retry => IDRETRY.0,
            DialogResult::TryAgain => IDTRYAGAIN.0,
            DialogResult::Continue => IDCONTINUE.0,
            DialogResult::Ignore => IDIGNORE.0,
            DialogResult::Unknown => 0,
        }
    }
}

/// Shows a message box dialog.
///
/// # Arguments
/// * `parent` - Optional parent window handle. If `None`, the dialog has no parent.
/// * `title` - The dialog title.
/// * `message` - The dialog message.
/// * `style` - The dialog style (determines icon and buttons).
/// * `set_foreground` - If `true`, the dialog will be brought to the foreground.
pub fn show_message_box(
    parent: Option<WindowHandle>,
    title: &str,
    message: &str,
    style: DialogStyle,
    set_foreground: bool,
) -> DialogResult {
    let parent_hwnd = parent.map(|h| h.raw()).unwrap_or(HWND(0));
    let title_wide = to_wide(title);
    let message_wide = to_wide(message);

    let mut flags = match style {
        DialogStyle::Info => MB_OK | MB_ICONINFORMATION,
        DialogStyle::Error => MB_OK | MB_ICONERROR,
        DialogStyle::Question => MB_OK | MB_ICONQUESTION,
        DialogStyle::YesNo => MB_YESNO | MB_ICONQUESTION,
        DialogStyle::YesNoCancel => MB_YESNOCANCEL | MB_ICONQUESTION,
    };

    if set_foreground {
        flags |= MB_SETFOREGROUND;
    }

    // SAFETY: MessageBoxW is a standard Win32 API. We provide valid null-terminated strings.
    let result = unsafe {
        MessageBoxW(
            parent_hwnd,
            PCWSTR(message_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            flags,
        )
    };

    match result {
        IDOK => DialogResult::Ok,
        IDCANCEL => DialogResult::Cancel,
        IDYES => DialogResult::Yes,
        IDNO => DialogResult::No,
        IDRETRY => DialogResult::Retry,
        IDTRYAGAIN => DialogResult::TryAgain,
        IDCONTINUE => DialogResult::Continue,
        IDIGNORE => DialogResult::Ignore,
        _ => DialogResult::Unknown,
    }
}
