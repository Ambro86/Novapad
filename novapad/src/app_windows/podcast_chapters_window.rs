//! Podcast chapters selection dialog.
//!
//! Shows a list of podcast chapters and returns the selected index.

use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Gdi::HFONT;

use crate::i18n;
use crate::podcast::chapters::Chapter;
use crate::settings::Language;
use crate::with_state;

/// Shows a dialog to select a chapter from a list.
///
/// Returns `Some(selected_index)` if user selected an item, `None` if cancelled or error.
pub fn select_chapter(parent: HWND, chapters: &[Chapter], language: Language) -> Option<usize> {
    if chapters.is_empty() {
        return None;
    }

    let items: Vec<String> = chapters
        .iter()
        .map(crate::podcast::chapters::chapter_label)
        .collect();

    // Get font from parent state
    // SAFETY: parent is a valid HWND from the caller
    let font = unsafe { with_state(parent, |state| state.hfont) }.unwrap_or(HFONT(0));
    let font_handle = if font.0 != 0 {
        Some(platform_windows::FontHandle::from_isize(font.0))
    } else {
        None
    };

    let params = platform_windows::ListSelectDialogParams {
        title: &i18n::tr(language, "podcasts.chapters.title"),
        items: &items,
        ok_label: &i18n::tr(language, "marker_select.ok"),
        cancel_label: &i18n::tr(language, "marker_select.cancel"),
        font: font_handle,
    };

    let parent_handle = platform_windows::WindowHandle::from_isize(parent.0);

    match platform_windows::show_list_select_dialog_index(Some(parent_handle), params) {
        Ok(result) => result,
        Err(e) => {
            crate::log_debug(&format!("Error showing chapter select dialog: {}", e));
            None
        }
    }
}
