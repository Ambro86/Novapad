//! Interpreter selection dialog.
//!
//! Shows a list of available interpreters and returns the selected one.

use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Gdi::HFONT;

use crate::i18n;
use crate::settings::Language;
use crate::with_state;

/// Shows a dialog to select an interpreter from a list.
///
/// Returns `Some(selected_path)` if user selected an item, `None` if cancelled or error.
pub fn select_interpreter(parent: HWND, items: Vec<String>, language: Language) -> Option<String> {
    if items.is_empty() {
        return None;
    }

    // Get font from parent state
    // SAFETY: parent is a valid HWND from the caller
    let font = unsafe { with_state(parent, |state| state.hfont) }.unwrap_or(HFONT(0));
    let font_handle = if font.0 != 0 {
        Some(platform_windows::FontHandle::from_isize(font.0))
    } else {
        None
    };

    let params = platform_windows::ListSelectDialogParams {
        title: &i18n::tr(language, "options.interpreter_search.title"),
        items: &items,
        ok_label: &i18n::tr(language, "options.ok"),
        cancel_label: &i18n::tr(language, "options.cancel"),
        font: font_handle,
    };

    let parent_handle = platform_windows::WindowHandle::from_isize(parent.0);

    match platform_windows::show_list_select_dialog(Some(parent_handle), params) {
        Ok(result) => result,
        Err(e) => {
            crate::log_debug(&format!("Error showing interpreter select dialog: {}", e));
            None
        }
    }
}
