//! Go to time dialog for audiobook seeking.
//!
//! Shows a text input dialog to enter a time position, then seeks the audiobook.

use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Gdi::HFONT;
use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;

use crate::accessibility::nvda_speak;
use crate::audio_player::{audiobook_duration_secs, parse_time_input, seek_audiobook_to};
use crate::i18n;
use crate::with_state;

/// Opens the "Go to Time" dialog for seeking in an audiobook.
///
/// This function shows a text input dialog, parses the entered time,
/// and seeks to that position. If parsing fails, it announces the error
/// via NVDA and re-shows the dialog.
///
/// # Safety
/// Requires valid parent HWND with associated app state.
pub unsafe fn open(parent: HWND) {
    // Check if there's an active audiobook
    let has_player = with_state(parent, |state| state.active_audiobook.is_some()).unwrap_or(false);
    if !has_player {
        return;
    }

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();

    // Get font from parent state
    let font = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));
    let font_handle = if font.0 != 0 {
        Some(platform_windows::FontHandle::from_isize(font.0))
    } else {
        None
    };

    let parent_handle = platform_windows::WindowHandle::from_isize(parent.0);

    // Loop to re-show dialog on parse error
    loop {
        let params = platform_windows::TextInputDialogParams {
            title: &i18n::tr(language, "go_to_time.title"),
            label: &i18n::tr(language, "go_to_time.label_time"),
            hint: &i18n::tr(language, "go_to_time.hint"),
            ok_label: &i18n::tr(language, "go_to_time.ok"),
            cancel_label: &i18n::tr(language, "go_to_time.cancel"),
            font: font_handle,
        };

        let result = match platform_windows::show_text_input_dialog(Some(parent_handle), params) {
            Ok(r) => r,
            Err(e) => {
                crate::log_debug(&format!("Error showing go to time dialog: {}", e));
                return;
            }
        };

        // User cancelled
        let Some(text) = result else {
            return;
        };

        // Empty input - treat as cancel
        if text.trim().is_empty() {
            return;
        }

        // Parse the time input
        let target = match parse_time_input(&text) {
            Ok(v) => v,
            Err(_) => {
                // Announce error via NVDA and re-show dialog
                let msg = i18n::tr(language, "go_to_time.invalid_time");
                nvda_speak(&msg);
                continue;
            }
        };

        // Get audiobook path and duration
        let (path, duration) = with_state(parent, |state| {
            state.active_audiobook.as_ref().map(|p| {
                let duration = audiobook_duration_secs(&p.path);
                (p.path.clone(), duration)
            })
        })
        .unwrap_or(None)
        .unwrap_or((std::path::PathBuf::new(), None));

        if path.as_os_str().is_empty() {
            return;
        }

        // Clamp to duration if needed
        let mut seek_target = target as u64;
        if let Some(max_secs) = duration
            && seek_target > max_secs
        {
            seek_target = max_secs;
            let msg = i18n::tr(language, "go_to_time.clamped");
            nvda_speak(&msg);
        }

        // Perform the seek
        crate::log_if_err!(seek_audiobook_to(parent, seek_target));

        // Restore focus to parent
        SetForegroundWindow(parent);

        return;
    }
}

/// Handles keyboard navigation for the go to time window.
/// Note: This is now handled internally by platform_windows::show_text_input_dialog
#[allow(dead_code)]
pub unsafe fn handle_navigation(
    _hwnd: HWND,
    _msg: &windows::Win32::UI::WindowsAndMessaging::MSG,
) -> bool {
    // Navigation is now handled by the platform_windows text input dialog
    false
}
