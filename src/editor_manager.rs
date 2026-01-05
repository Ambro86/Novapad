#![allow(clippy::if_same_then_else, clippy::collapsible_else_if)]
use crate::accessibility::{EM_REPLACESEL, from_wide, to_wide, to_wide_normalized};
use crate::file_handler::decode_text_with_encoding;
use crate::file_handler::*;
use crate::settings::{
    FileFormat, ModifiedMarkerPosition, TextEncoding, confirm_save_message, confirm_title,
    untitled_base, untitled_title,
};
use crate::{log_debug, with_state};
use std::collections::HashSet;
use std::path::{Path, PathBuf};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::HFONT;
use windows::Win32::UI::Controls::RichEdit::{
    CFM_COLOR, CFM_SIZE, CHARFORMAT2W, CHARRANGE, EM_EXGETSEL, EM_EXSETSEL, EM_SETCHARFORMAT,
    EM_SETEVENTMASK, ENM_CHANGE, MSFTEDIT_CLASS, SCF_ALL,
};
use windows::Win32::UI::Controls::{
    EM_GETMODIFY, EM_SETMODIFY, EM_SETREADONLY, TCIF_TEXT, TCITEMW, TCM_ADJUSTRECT,
    TCM_INSERTITEMW, TCM_SETCURSEL, TCM_SETITEMW,
};
use windows::Win32::UI::Input::KeyboardAndMouse::SetFocus;
use windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, ES_AUTOHSCROLL, ES_AUTOVSCROLL, ES_MULTILINE, ES_WANTRETURN, GetClientRect,
    GetWindowTextLengthW, GetWindowTextW, HMENU, IDNO, IDYES, MB_ICONWARNING, MB_YESNOCANCEL,
    MessageBoxW, MoveWindow, SW_HIDE, SW_SHOW, SendMessageW, SetWindowTextW, ShowWindow,
    WM_SETFONT, WS_CHILD, WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, WS_GROUP, WS_HSCROLL, WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR};

const EM_LIMITTEXT: u32 = 0x00C5;
const EM_BEGINUNDOACTION: u32 = 0x0459;
const EM_ENDUNDOACTION: u32 = 0x045A;
const EM_STOPGROUPTYPING: u32 = 0x0477;
const EM_SETTEXTEX: u32 = 0x0461;
const EM_GETTEXTLENGTHEX: u32 = 0x045F;
const ST_KEEPUNDO: u32 = 0x0001;
const ST_SELECTION: u32 = 0x0002;
const GTL_NUMCHARS: u32 = 0x0008;
const CP_UNICODE: u32 = 1200;
const VOICE_PANEL_PADDING: i32 = 6;
const VOICE_PANEL_ROW_HEIGHT: i32 = 22;
const VOICE_PANEL_SPACING: i32 = 6;
const VOICE_PANEL_LABEL_WIDTH: i32 = 140;
const VOICE_PANEL_COMBO_HEIGHT: i32 = 140;

fn apply_text_limit(hwnd_edit: HWND) {
    unsafe {
        if hwnd_edit.0 != 0 {
            SendMessageW(hwnd_edit, EM_LIMITTEXT, WPARAM(0x7FFFFFFE), LPARAM(0));
        }
    }
}

pub unsafe fn apply_text_limit_to_all_edits(hwnd: HWND) {
    let edits = with_state(hwnd, |state| {
        state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>()
    })
    .unwrap_or_default();

    for hwnd_edit in edits {
        apply_text_limit(hwnd_edit);
    }
}

pub struct Document {
    pub title: String,
    pub path: Option<PathBuf>,
    pub hwnd_edit: HWND,
    pub dirty: bool,
    pub format: FileFormat,
    pub opened_text_encoding: Option<TextEncoding>,
    pub current_save_text_encoding: Option<TextEncoding>,
}

#[derive(Clone)]
pub struct NormalizeUndo {
    pub hwnd_edit: HWND,
    pub text: String,
    pub sel_start: i32,
    pub sel_end: i32,
    pub was_dirty: bool,
}

impl Default for Document {
    fn default() -> Self {
        Document {
            title: String::new(),
            path: None,
            hwnd_edit: HWND(0),
            dirty: false,
            format: FileFormat::Text(TextEncoding::Utf8),
            opened_text_encoding: None,
            current_save_text_encoding: None,
        }
    }
}

// --- Editor Helpers ---

pub unsafe fn set_edit_text(hwnd_edit: HWND, text: &str) {
    let wide = to_wide_normalized(text);
    if hwnd_edit.0 != 0 {
        // Prevent programmatic loads from marking the document as modified.
        SendMessageW(hwnd_edit, EM_SETEVENTMASK, WPARAM(0), LPARAM(0));
    }
    let _ = SetWindowTextW(hwnd_edit, PCWSTR(wide.as_ptr()));
    if hwnd_edit.0 != 0 {
        SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
        SendMessageW(
            hwnd_edit,
            EM_SETEVENTMASK,
            WPARAM(0),
            LPARAM(ENM_CHANGE as isize),
        );
    }
}

pub unsafe fn get_edit_text(hwnd_edit: HWND) -> String {
    let len = GetWindowTextLengthW(hwnd_edit);
    if len == 0 {
        return String::new();
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    GetWindowTextW(hwnd_edit, &mut buf);
    from_wide(buf.as_ptr())
}

pub unsafe fn send_to_active_edit(hwnd: HWND, msg: u32) {
    if let Some(hwnd_edit) = crate::get_active_edit(hwnd) {
        SendMessageW(hwnd_edit, msg, WPARAM(0), LPARAM(0));
    }
}

pub unsafe fn select_all_active_edit(hwnd: HWND) {
    if let Some(hwnd_edit) = crate::get_active_edit(hwnd) {
        let cr = CHARRANGE {
            cpMin: 0,
            cpMax: -1,
        };
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&cr as *const _ as isize),
        );
    }
}

pub unsafe fn apply_word_wrap_to_all_edits(hwnd: HWND, word_wrap: bool) {
    let edits = with_state(hwnd, |state| {
        state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>()
    })
    .unwrap_or_default();

    for hwnd_edit in edits {
        if hwnd_edit.0 == 0 {
            continue;
        }
        log_debug(&format!(
            "Word wrap toggle for {:?}: {}",
            hwnd_edit, word_wrap
        ));
        apply_text_limit(hwnd_edit);
    }
}

pub unsafe fn apply_text_appearance_to_all_edits(hwnd: HWND, text_color: u32, text_size: i32) {
    let edits = with_state(hwnd, |state| {
        state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>()
    })
    .unwrap_or_default();

    for hwnd_edit in edits {
        if hwnd_edit.0 == 0 {
            continue;
        }
        apply_text_appearance(hwnd_edit, text_color, text_size);
    }
}

fn apply_text_appearance(hwnd_edit: HWND, text_color: u32, text_size: i32) {
    let mut format = CHARFORMAT2W::default();
    format.Base.cbSize = std::mem::size_of::<CHARFORMAT2W>() as u32;
    format.Base.dwMask = CFM_COLOR | CFM_SIZE;
    format.Base.crTextColor = windows::Win32::Foundation::COLORREF(text_color);
    if text_size > 0 {
        format.Base.yHeight = text_size.saturating_mul(20);
    }
    unsafe {
        SendMessageW(
            hwnd_edit,
            EM_SETCHARFORMAT,
            WPARAM(SCF_ALL as usize),
            LPARAM(&mut format as *mut _ as isize),
        );
    }
}

#[repr(C)]
struct SetTextEx {
    flags: u32,
    codepage: u32,
}

#[repr(C)]
struct GetTextLengthEx {
    flags: u32,
    codepage: u32,
}

fn begin_single_undo_action(hwnd_edit: HWND) {
    unsafe {
        SendMessageW(hwnd_edit, EM_STOPGROUPTYPING, WPARAM(0), LPARAM(0));
        SendMessageW(hwnd_edit, EM_BEGINUNDOACTION, WPARAM(0), LPARAM(0));
    }
}

fn end_single_undo_action(hwnd_edit: HWND) {
    unsafe {
        SendMessageW(hwnd_edit, EM_ENDUNDOACTION, WPARAM(0), LPARAM(0));
    }
}

pub unsafe fn try_normalize_undo(hwnd: HWND) -> bool {
    let mut undo = None;
    let _ = with_state(hwnd, |state| {
        undo = state.normalize_undo.clone();
        state.normalize_undo = None;
    });
    let Some(undo) = undo else {
        return false;
    };
    if undo.hwnd_edit.0 == 0 {
        return false;
    }
    set_edit_text(undo.hwnd_edit, &undo.text);
    let mut cr = CHARRANGE {
        cpMin: undo.sel_start,
        cpMax: undo.sel_end,
    };
    SendMessageW(
        undo.hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );
    let _ = with_state(hwnd, |state| {
        for (idx, doc) in state.docs.iter_mut().enumerate() {
            if doc.hwnd_edit == undo.hwnd_edit {
                doc.dirty = undo.was_dirty;
                update_tab_title(state.hwnd_tab, idx, &doc.title, doc.dirty);
                if state.current == idx {
                    update_window_title(hwnd);
                }
                break;
            }
        }
    });
    SetFocus(undo.hwnd_edit);
    true
}

pub unsafe fn handle_normalize_edit_change(hwnd: HWND, hwnd_edit: HWND) {
    let _ = with_state(hwnd, |state| {
        if state.normalize_skip_change {
            state.normalize_skip_change = false;
            return;
        }
        if let Some(pending) = &state.normalize_undo {
            if pending.hwnd_edit == hwnd_edit {
                state.normalize_undo = None;
            }
        }
    });
}

pub unsafe fn strip_markdown_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }
    let cleaned = strip_markdown_text(&text);
    if cleaned == text {
        return;
    }
    let mut replace_range = CHARRANGE {
        cpMin: 0,
        cpMax: byte_index_to_utf16(&text, text.len()),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&cleaned);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn normalize_whitespace_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let mut length_info = GetTextLengthEx {
        flags: GTL_NUMCHARS,
        codepage: CP_UNICODE,
    };
    let total_chars = SendMessageW(
        hwnd_edit,
        EM_GETTEXTLENGTHEX,
        WPARAM(&mut length_info as *mut _ as usize),
        LPARAM(0),
    )
    .0 as i32;
    let mut sel_start = selection.cpMin;
    let mut sel_end = selection.cpMax;
    if sel_start < 0 {
        sel_start = 0;
    }
    if sel_end < 0 {
        sel_end = total_chars;
    }
    if sel_end > total_chars {
        sel_end = total_chars;
    }
    let near_end = sel_end >= total_chars.saturating_sub(1);
    if near_end {
        sel_end = total_chars;
    }

    let has_selection = sel_start != sel_end;
    let whole_doc_selected = has_selection && sel_start == 0 && near_end;
    let (start_byte, end_byte) = if has_selection {
        (
            utf16_index_to_byte(&text, sel_start),
            utf16_index_to_byte(&text, sel_end),
        )
    } else {
        (0, text.len())
    };

    let (affected_start, affected_end) = if has_selection {
        if whole_doc_selected {
            (0, text.len())
        } else {
            let mut effective_end = end_byte;
            if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
                effective_end = end_byte.saturating_sub(1);
            }
            let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
            let line_end = text[effective_end..]
                .find('\n')
                .map(|i| effective_end + i + 1)
                .unwrap_or(text.len());
            (line_start.min(start_byte), line_end.max(end_byte))
        }
    } else {
        (0, text.len())
    };

    let affected = &text[affected_start..affected_end];
    let normalized = normalize_whitespace_block(affected, line_ending);
    if normalized == affected {
        return;
    }
    let was_dirty = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .find(|doc| doc.hwnd_edit == hwnd_edit)
            .map(|doc| doc.dirty)
            .unwrap_or(false)
    })
    .unwrap_or(false);

    let mut replace_range = if whole_doc_selected {
        CHARRANGE {
            cpMin: 0,
            cpMax: -1,
        }
    } else {
        CHARRANGE {
            cpMin: byte_index_to_utf16(&text, affected_start),
            cpMax: byte_index_to_utf16(&text, affected_end),
        }
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    SendMessageW(hwnd_edit, EM_STOPGROUPTYPING, WPARAM(0), LPARAM(0));
    let _ = with_state(hwnd, |state| {
        state.normalize_undo = Some(NormalizeUndo {
            hwnd_edit,
            text: text.clone(),
            sel_start,
            sel_end,
            was_dirty,
        });
        state.normalize_skip_change = true;
    });
    let mut set_text = SetTextEx {
        flags: ST_KEEPUNDO | ST_SELECTION,
        codepage: CP_UNICODE,
    };
    let replace_wide = to_wide(&normalized);
    SendMessageW(
        hwnd_edit,
        EM_SETTEXTEX,
        WPARAM(&mut set_text as *mut _ as usize),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn hard_line_break_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }
    let wrap_width = with_state(hwnd, |state| state.settings.wrap_width).unwrap_or(80);
    let wrap_width = wrap_width.max(1) as usize;
    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };

    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (range_start, range_end, has_trailing_newline) = if selection.cpMin != selection.cpMax {
        let start = utf16_index_to_byte(&text, selection.cpMin);
        let end = utf16_index_to_byte(&text, selection.cpMax);
        let selected = &text[start..end];
        (start, end, selected.ends_with('\n'))
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let target = &text[range_start..range_end];
    let reformatted = reflow_block_text(target, wrap_width, line_ending, has_trailing_newline);
    if reformatted == target {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, range_start),
        cpMax: byte_index_to_utf16(&text, range_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&reformatted);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn order_items_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let ordered = order_lines_block(affected, line_ending, has_trailing_newline);
    if ordered == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&ordered);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn keep_unique_items_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let cleaned = keep_unique_lines_block(affected, line_ending, has_trailing_newline);
    if cleaned == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&cleaned);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn reverse_items_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let reversed = reverse_lines_block(affected, line_ending, has_trailing_newline);
    if reversed == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&reversed);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn quote_lines_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let quote_prefix = with_state(hwnd, |state| state.settings.quote_prefix.clone())
        .unwrap_or_else(|| "> ".to_string());
    if quote_prefix.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let quoted = quote_lines_block(affected, line_ending, has_trailing_newline, &quote_prefix);
    if quoted == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&quoted);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn unquote_lines_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let quote_prefix = with_state(hwnd, |state| state.settings.quote_prefix.clone())
        .unwrap_or_else(|| "> ".to_string());
    if quote_prefix.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let unquoted = unquote_lines_block(affected, line_ending, has_trailing_newline, &quote_prefix);
    if unquoted == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&unquoted);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn join_lines_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        let caret = utf16_index_to_byte(&text, selection.cpMin);
        let Some((start, end, trailing)) = paragraph_range_bytes(&text, caret) else {
            return;
        };
        (start, end, trailing)
    };

    let affected = &text[affected_start..affected_end];
    let joined = join_lines_block(affected, line_ending, has_trailing_newline);
    if joined == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&joined);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn clean_end_of_line_hyphens_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return;
    }

    let line_ending = if text.contains("\r\n") { "\r\n" } else { "\n" };
    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let (affected_start, affected_end, has_trailing_newline) = if selection.cpMin != selection.cpMax
    {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        let mut effective_end = end_byte;
        if end_byte > start_byte && end_byte > 0 && text.as_bytes()[end_byte - 1] == b'\n' {
            effective_end = end_byte.saturating_sub(1);
        }
        let line_start = text[..start_byte].rfind('\n').map(|i| i + 1).unwrap_or(0);
        let line_end = text[effective_end..]
            .find('\n')
            .map(|i| effective_end + i + 1)
            .unwrap_or(text.len());
        (
            line_start,
            line_end,
            text[line_start..line_end].ends_with('\n'),
        )
    } else {
        (0, text.len(), text.ends_with('\n'))
    };

    let affected = &text[affected_start..affected_end];
    let cleaned = clean_end_of_line_hyphens_block(affected, line_ending, has_trailing_newline);
    if cleaned == affected {
        return;
    }

    let mut replace_range = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, affected_start),
        cpMax: byte_index_to_utf16(&text, affected_end),
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut replace_range as *mut _ as isize),
    );
    // Single-undo guarantee.
    begin_single_undo_action(hwnd_edit);
    let replace_wide = to_wide(&cleaned);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    end_single_undo_action(hwnd_edit);
    mark_dirty_from_edit(hwnd, hwnd_edit);
    SetFocus(hwnd_edit);
}

pub unsafe fn text_stats_active_edit(hwnd: HWND) {
    let Some(hwnd_edit) = crate::get_active_edit(hwnd) else {
        return;
    };
    let text = get_edit_text(hwnd_edit);
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if text.is_empty() {
        let message = build_text_stats_message(language, 0, 0, 0, 0);
        crate::show_info(hwnd, language, &message);
        return;
    }

    let mut selection = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut selection as *mut _ as isize),
    );

    let target = if selection.cpMin != selection.cpMax {
        let start_byte = utf16_index_to_byte(&text, selection.cpMin);
        let end_byte = utf16_index_to_byte(&text, selection.cpMax);
        &text[start_byte..end_byte]
    } else {
        &text[..]
    };

    let chars_with_spaces = target.chars().count();
    let chars_without_spaces = target.chars().filter(|c| !c.is_whitespace()).count();
    let words = target.split_whitespace().count();
    let lines = if target.is_empty() {
        0
    } else {
        target.as_bytes().iter().filter(|b| **b == b'\n').count() + 1
    };
    let message = build_text_stats_message(
        language,
        chars_with_spaces,
        chars_without_spaces,
        words,
        lines,
    );
    crate::show_info(hwnd, language, &message);
    SetFocus(hwnd_edit);
}

fn normalize_whitespace_block(text: &str, line_ending: &str) -> String {
    let mut out_lines = Vec::new();
    let mut blank_run = 0usize;
    for raw_line in text.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        let trimmed = line.trim();
        if trimmed.is_empty() {
            blank_run += 1;
            if blank_run <= 2 {
                out_lines.push(String::new());
            }
        } else {
            blank_run = 0;
            out_lines.push(trimmed.to_string());
        }
    }
    out_lines.join(line_ending)
}

fn order_lines_block(text: &str, line_ending: &str, has_trailing_newline: bool) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut lines: Vec<String> = content
        .split('\n')
        .map(|raw_line| raw_line.strip_suffix('\r').unwrap_or(raw_line).to_string())
        .collect();

    let mut nonblank_indices = Vec::new();
    let mut nonblank_lines = Vec::new();
    for (idx, line) in lines.iter().enumerate() {
        if !line.trim().is_empty() {
            nonblank_indices.push(idx);
            nonblank_lines.push((line.clone(), idx));
        }
    }

    if nonblank_lines.len() > 1 {
        nonblank_lines.sort_by_key(|(line, idx)| (line.to_ascii_lowercase(), *idx));
    }

    for (slot, (line, _)) in nonblank_indices.into_iter().zip(nonblank_lines.into_iter()) {
        lines[slot] = line;
    }

    let mut out = lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn keep_unique_lines_block(text: &str, line_ending: &str, has_trailing_newline: bool) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut seen: HashSet<String> = HashSet::new();
    let mut out_lines: Vec<String> = Vec::new();

    for raw_line in content.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        if line.trim().is_empty() {
            out_lines.push(line.to_string());
            continue;
        }
        let key = line.to_ascii_lowercase();
        if seen.insert(key) {
            out_lines.push(line.to_string());
        }
    }

    let mut out = out_lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn reverse_lines_block(text: &str, line_ending: &str, has_trailing_newline: bool) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut lines: Vec<String> = content
        .split('\n')
        .map(|raw_line| raw_line.strip_suffix('\r').unwrap_or(raw_line).to_string())
        .collect();

    let mut nonblank_indices = Vec::new();
    let mut nonblank_lines = Vec::new();
    for (idx, line) in lines.iter().enumerate() {
        if !line.trim().is_empty() {
            nonblank_indices.push(idx);
            nonblank_lines.push(line.clone());
        }
    }

    if nonblank_lines.len() > 1 {
        nonblank_lines.reverse();
    }

    for (slot, line) in nonblank_indices.into_iter().zip(nonblank_lines.into_iter()) {
        lines[slot] = line;
    }

    let mut out = lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn quote_lines_block(
    text: &str,
    line_ending: &str,
    has_trailing_newline: bool,
    quote_prefix: &str,
) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut out_lines: Vec<String> = Vec::new();

    for raw_line in content.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        if line.trim().is_empty() {
            out_lines.push(line.to_string());
        } else {
            let mut quoted = String::with_capacity(quote_prefix.len() + line.len());
            quoted.push_str(quote_prefix);
            quoted.push_str(line);
            out_lines.push(quoted);
        }
    }

    let mut out = out_lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn unquote_lines_block(
    text: &str,
    line_ending: &str,
    has_trailing_newline: bool,
    quote_prefix: &str,
) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut out_lines: Vec<String> = Vec::new();

    for raw_line in content.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        if line.trim().is_empty() {
            out_lines.push(line.to_string());
        } else if let Some(rest) = line.strip_prefix(quote_prefix) {
            out_lines.push(rest.to_string());
        } else {
            out_lines.push(line.to_string());
        }
    }

    let mut out = out_lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn join_lines_block(text: &str, line_ending: &str, has_trailing_newline: bool) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut out = String::new();
    let mut has_content = false;

    for raw_line in content.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        if line.trim().is_empty() {
            continue;
        }
        if !has_content {
            out.push_str(line);
            has_content = true;
            continue;
        }

        let prev_ends_ws = out.chars().last().is_some_and(|c| c.is_whitespace());
        let next_starts_ws = line.chars().next().is_some_and(|c| c.is_whitespace());
        if !prev_ends_ws && !next_starts_ws {
            let prev_is_word = out
                .chars()
                .last()
                .is_some_and(|c| c.is_alphanumeric() || c == '_');
            let next_is_word = line
                .chars()
                .next()
                .is_some_and(|c| c.is_alphanumeric() || c == '_');
            if prev_is_word && next_is_word {
                out.push(' ');
            }
        }
        out.push_str(line);
    }

    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn clean_end_of_line_hyphens_block(
    text: &str,
    line_ending: &str,
    has_trailing_newline: bool,
) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut out = String::with_capacity(content.len());
    let chars: Vec<char> = content.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '-' {
            // Check char before '-'
            let mut valid_before = false;
            if i > 0 {
                let prev = chars[i - 1];
                if prev.is_alphanumeric() && !prev.is_whitespace() {
                    valid_before = true;
                }
            }

            if valid_before {
                // Look ahead for line breaks (up to 3)
                let mut temp_j = i + 1;
                let mut line_breaks = 0;

                while line_breaks < 3 && temp_j < chars.len() {
                    if chars[temp_j] == '\r' {
                        temp_j += 1;
                        continue;
                    }
                    if chars[temp_j] == '\n' {
                        line_breaks += 1;
                        temp_j += 1;
                        // Skip any optional \r\n that might follow (if we allow more line breaks)
                        continue;
                    }
                    // If we reach here, it's not a line break char.
                    break;
                }

                if line_breaks > 0 && temp_j < chars.len() {
                    let next_char = chars[temp_j];
                    if next_char.is_alphabetic() {
                        // Join! Skip the hyphen and the line breaks.
                        i = temp_j;
                        continue;
                    }
                }
            }
        }
        out.push(chars[i]);
        i += 1;
    }

    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn build_text_stats_message(
    language: crate::settings::Language,
    chars_with_spaces: usize,
    chars_without_spaces: usize,
    words: usize,
    lines: usize,
) -> String {
    let with_spaces = crate::i18n::tr_f(
        language,
        "text_stats.characters_with_spaces",
        &[("count", &chars_with_spaces.to_string())],
    );
    let without_spaces = crate::i18n::tr_f(
        language,
        "text_stats.characters_without_spaces",
        &[("count", &chars_without_spaces.to_string())],
    );
    let words = crate::i18n::tr_f(
        language,
        "text_stats.words",
        &[("count", &words.to_string())],
    );
    let lines = crate::i18n::tr_f(
        language,
        "text_stats.lines",
        &[("count", &lines.to_string())],
    );
    format!("{with_spaces}.\n{without_spaces}.\n{words}.\n{lines}.")
}

fn reflow_block_text(
    text: &str,
    wrap_width: usize,
    line_ending: &str,
    has_trailing_newline: bool,
) -> String {
    let (content, trailing_newline) = split_trailing_newline(text, has_trailing_newline);
    let mut out_lines: Vec<String> = Vec::new();
    let mut current_words: Vec<&str> = Vec::new();

    for raw_line in content.split('\n') {
        let line = raw_line.strip_suffix('\r').unwrap_or(raw_line);
        if line.trim().is_empty() {
            if !current_words.is_empty() {
                out_lines.extend(wrap_words(current_words.drain(..), wrap_width));
            }
            out_lines.push(String::new());
        } else {
            current_words.extend(line.split_whitespace());
        }
    }

    if !current_words.is_empty() {
        out_lines.extend(wrap_words(current_words.drain(..), wrap_width));
    }

    let mut out = out_lines.join(line_ending);
    if trailing_newline {
        out.push_str(line_ending);
    }
    out
}

fn split_trailing_newline(text: &str, prefer_trailing: bool) -> (&str, bool) {
    if prefer_trailing && text.ends_with("\r\n") {
        return (&text[..text.len().saturating_sub(2)], true);
    }
    if prefer_trailing && text.ends_with('\n') {
        return (&text[..text.len().saturating_sub(1)], true);
    }
    (text, false)
}

fn wrap_words<'a, I>(words: I, wrap_width: usize) -> Vec<String>
where
    I: Iterator<Item = &'a str>,
{
    let mut lines = Vec::new();
    let mut current = String::new();
    let mut current_len = 0usize;

    for word in words {
        let word_len = word.chars().count();
        if current_len == 0 {
            current.push_str(word);
            current_len = word_len;
            continue;
        }
        if current_len + 1 + word_len <= wrap_width {
            current.push(' ');
            current.push_str(word);
            current_len += 1 + word_len;
        } else {
            lines.push(current);
            current = word.to_string();
            current_len = word_len;
        }
    }

    if current_len > 0 {
        lines.push(current);
    }
    lines
}

fn paragraph_range_bytes(text: &str, caret: usize) -> Option<(usize, usize, bool)> {
    if text.is_empty() {
        return None;
    }
    let mut lines: Vec<(usize, usize, usize, bool)> = Vec::new();
    let mut start = 0usize;
    for (idx, byte) in text.as_bytes().iter().enumerate() {
        if *byte == b'\n' {
            let end = idx;
            let line = &text[start..end];
            let line = line.strip_suffix('\r').unwrap_or(line);
            let is_blank = line.trim().is_empty();
            lines.push((start, end, idx + 1, is_blank));
            start = idx + 1;
        }
    }
    if start <= text.len() {
        let end = text.len();
        let line = &text[start..end];
        let line = line.strip_suffix('\r').unwrap_or(line);
        let is_blank = line.trim().is_empty();
        lines.push((start, end, end, is_blank));
    }

    let mut line_idx = lines.len().saturating_sub(1);
    for (idx, line) in lines.iter().enumerate() {
        if caret < line.2 {
            line_idx = idx;
            break;
        }
    }
    if lines[line_idx].3 {
        return None;
    }
    let mut start_idx = line_idx;
    while start_idx > 0 && !lines[start_idx - 1].3 {
        start_idx = start_idx.saturating_sub(1);
    }
    let mut end_idx = line_idx;
    while end_idx + 1 < lines.len() && !lines[end_idx + 1].3 {
        end_idx += 1;
    }
    let range_start = lines[start_idx].0;
    let range_end = lines[end_idx].2;
    let has_trailing_newline = lines[end_idx].2 > lines[end_idx].1;
    Some((range_start, range_end, has_trailing_newline))
}

fn utf16_index_to_byte(text: &str, target: i32) -> usize {
    if target <= 0 {
        return 0;
    }
    let target = target as usize;
    let mut utf16_count = 0usize;
    for (byte_idx, ch) in text.char_indices() {
        let units = ch.len_utf16();
        let next = utf16_count + units;
        if target <= next {
            if target == next {
                return byte_idx + ch.len_utf8();
            }
            return byte_idx;
        }
        utf16_count = next;
    }
    text.len()
}

fn byte_index_to_utf16(text: &str, byte_idx: usize) -> i32 {
    let mut utf16_count = 0usize;
    for (idx, ch) in text.char_indices() {
        if idx >= byte_idx {
            break;
        }
        utf16_count += ch.len_utf16();
    }
    utf16_count as i32
}

fn strip_markdown_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for line in text.split_inclusive('\n') {
        let (content, line_end) = if let Some(pos) = line.find('\n') {
            (&line[..pos], &line[pos..])
        } else {
            (line, "")
        };
        let mut trimmed = content.trim_start();
        if trimmed.starts_with("```") {
            trimmed = trimmed.trim_start_matches('`').trim_start();
        }
        if trimmed.starts_with('#') {
            trimmed = trimmed.trim_start_matches('#').trim_start();
        }
        if trimmed.starts_with('>') {
            trimmed = trimmed.trim_start_matches('>').trim_start();
        }
        if trimmed.starts_with("- ") || trimmed.starts_with("* ") || trimmed.starts_with("+ ") {
            trimmed = trimmed[2..].trim_start();
        } else if let Some(rest) = trimmed.strip_prefix(|c: char| c.is_ascii_digit()) {
            let rest = rest.trim_start();
            if let Some(rest) = rest.strip_prefix('.') {
                trimmed = rest.trim_start();
            }
        }
        let mut cleaned = strip_markdown_inline(trimmed);
        cleaned.push_str(line_end);
        out.push_str(&cleaned);
    }
    out
}

fn strip_markdown_inline(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut chars = text.chars().peekable();
    while let Some(&ch) = chars.peek() {
        let _ = chars.next();
        if ch == '`' {
            continue;
        }
        if ch == '*' || ch == '_' {
            if let Some(next) = chars.peek() {
                if *next == ch {
                    let _ = chars.next();
                    continue;
                }
            }
        }
        if ch == '~' {
            if let Some(next) = chars.peek() {
                if *next == '~' {
                    let _ = chars.next();
                    continue;
                }
            }
        }
        if ch == '!' && chars.peek() == Some(&'[') {
            let _ = chars.next();
            let alt = collect_bracket_text(&mut chars, ']');
            if chars.peek() == Some(&'(') {
                let _ = chars.next();
                let _ = collect_bracket_text(&mut chars, ')');
            }
            out.push_str(&alt);
            continue;
        }
        if ch == '[' {
            let label = collect_bracket_text(&mut chars, ']');
            if chars.peek() == Some(&'(') {
                let _ = chars.next();
                let _ = collect_bracket_text(&mut chars, ')');
                out.push_str(&label);
                continue;
            }
            out.push('[');
            out.push_str(&label);
            continue;
        }
        out.push(ch);
    }
    out
}

fn collect_bracket_text<I: Iterator<Item = char>>(
    chars: &mut std::iter::Peekable<I>,
    end: char,
) -> String {
    let mut out = String::new();
    for ch in chars.by_ref() {
        if ch == end {
            break;
        }
        out.push(ch);
    }
    out
}

// --- Document Management ---

pub unsafe fn new_document(hwnd: HWND) {
    let new_index = with_state(hwnd, |state| {
        state.untitled_count += 1;
        let language = state.settings.language;
        let title = untitled_title(language, state.untitled_count);
        let hwnd_edit = create_edit(
            hwnd,
            state.hfont,
            state.settings.word_wrap,
            state.settings.text_color,
            state.settings.text_size,
        );
        let doc = Document {
            title: title.clone(),
            path: None,
            hwnd_edit,
            dirty: false,
            format: FileFormat::Text(TextEncoding::Utf8),
            opened_text_encoding: None,
            current_save_text_encoding: None,
        };
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, &title, (state.docs.len() - 1) as i32);
        state.docs.len() - 1
    })
    .unwrap_or(0);
    select_tab(hwnd, new_index);
}

pub unsafe fn open_document_with_encoding(
    hwnd: HWND,
    path: &Path,
    user_encoding: Option<TextEncoding>,
) {
    log_debug(&format!(
        "Open document: {} (encoding: {:?})",
        path.display(),
        user_encoding
    ));

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if is_pdf_path(path) {
        crate::open_pdf_document_async(hwnd, path);
        return;
    }
    let (content, format, opened_text_encoding) = if is_docx_path(path) {
        match read_docx_text(path, language) {
            Ok(text) => (text, FileFormat::Docx, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_pptx_path(path) {
        match read_ppt_text(path, language) {
            Ok(text) => (text, FileFormat::Pptx, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_ppt_path(path) {
        match read_ppt_text(path, language) {
            Ok(text) => (text, FileFormat::Ppt, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_epub_path(path) {
        match read_epub_text(path, language) {
            Ok(text) => (text, FileFormat::Epub, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_html_path(path) {
        match read_html_text(path, language) {
            Ok((text, _encoding)) => (text, FileFormat::Html, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_mp3_path(path) {
        (String::new(), FileFormat::Audiobook, None)
    } else if is_doc_path(path) {
        match read_doc_text(path, language) {
            Ok(text) => (text, FileFormat::Doc, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_spreadsheet_path(path) {
        match read_spreadsheet_text(path, language) {
            Ok(text) => (text, FileFormat::Spreadsheet, None),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else {
        match std::fs::read(path) {
            Ok(bytes) => {
                if let Some(encoding) = user_encoding {
                    // User specified encoding
                    match decode_text_with_encoding(&bytes, encoding, language) {
                        Ok(text) => (text, FileFormat::Text(encoding), Some(encoding)),
                        Err(message) => {
                            crate::show_error(hwnd, language, &message);
                            return;
                        }
                    }
                } else {
                    // Auto-detect encoding
                    match decode_text(&bytes, language) {
                        Ok((text, encoding)) => (text, FileFormat::Text(encoding), Some(encoding)),
                        Err(message) => {
                            crate::show_error(hwnd, language, &message);
                            return;
                        }
                    }
                }
            }
            Err(err) => {
                crate::show_error(
                    hwnd,
                    language,
                    &crate::settings::error_open_file_message(language, err),
                );
                return;
            }
        }
    };

    let new_index = with_state(hwnd, |state| {
        let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
        let hwnd_edit = create_edit(
            hwnd,
            state.hfont,
            state.settings.word_wrap,
            state.settings.text_color,
            state.settings.text_size,
        );
        set_edit_text(hwnd_edit, &content);

        let doc = Document {
            title: title.to_string(),
            path: Some(path.to_path_buf()),
            hwnd_edit,
            dirty: false,
            format,
            opened_text_encoding,
            current_save_text_encoding: None,
        };
        if matches!(format, FileFormat::Audiobook) {
            unsafe {
                SendMessageW(hwnd_edit, EM_SETREADONLY, WPARAM(1), LPARAM(0));
                ShowWindow(hwnd_edit, SW_HIDE);
            }
        }
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, title, (state.docs.len() - 1) as i32);
        crate::goto_first_bookmark(hwnd_edit, path, &state.bookmarks, format);
        state.docs.len() - 1
    })
    .unwrap_or(0);
    select_tab(hwnd, new_index);
    if matches!(format, FileFormat::Audiobook) {
        unsafe {
            crate::audio_player::start_audiobook_playback(hwnd, path);
        }
    }
    crate::push_recent_file(hwnd, path);
}

pub unsafe fn open_document(hwnd: HWND, path: &Path) {
    open_document_with_encoding(hwnd, path, None);
}

pub unsafe fn select_tab(hwnd: HWND, index: usize) {
    let result = with_state(hwnd, |state| {
        if index >= state.docs.len() {
            return None;
        }
        let prev = state.current;
        let prev_edit = state.docs.get(prev).map(|doc| doc.hwnd_edit);
        let new_doc = state.docs.get(index);
        let new_edit = new_doc.map(|doc| doc.hwnd_edit);
        let is_audiobook = new_doc
            .map(|doc| matches!(doc.format, FileFormat::Audiobook))
            .unwrap_or(false);
        state.current = index;
        Some((state.hwnd_tab, prev_edit, new_edit, is_audiobook))
    })
    .flatten();

    let Some((hwnd_tab, prev_edit, new_edit, is_audiobook)) = result else {
        return;
    };

    if let Some(hwnd_edit) = prev_edit {
        ShowWindow(hwnd_edit, SW_HIDE);
    }
    SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(index), LPARAM(0));
    if let Some(hwnd_edit) = new_edit {
        if is_audiobook {
            ShowWindow(hwnd_edit, SW_HIDE);
            SetFocus(hwnd_tab);
        } else {
            ShowWindow(hwnd_edit, SW_SHOW);
            SetFocus(hwnd_edit);
        }
    }
    update_window_title(hwnd);
    layout_children(hwnd);
}

pub unsafe fn insert_tab(hwnd_tab: HWND, title: &str, index: i32) {
    let mut text = to_wide(title);
    let mut item = TCITEMW {
        mask: TCIF_TEXT,
        pszText: PWSTR(text.as_mut_ptr()),
        ..Default::default()
    };
    SendMessageW(
        hwnd_tab,
        TCM_INSERTITEMW,
        WPARAM(index as usize),
        LPARAM(&mut item as *mut _ as isize),
    );
}

pub unsafe fn update_tab_title(hwnd_tab: HWND, index: usize, title: &str, dirty: bool) {
    let label = if dirty {
        format!("{title}*")
    } else {
        title.to_string()
    };
    let mut text = to_wide(&label);
    let mut item = TCITEMW {
        mask: TCIF_TEXT,
        pszText: PWSTR(text.as_mut_ptr()),
        ..Default::default()
    };
    SendMessageW(
        hwnd_tab,
        TCM_SETITEMW,
        WPARAM(index),
        LPARAM(&mut item as *mut _ as isize),
    );
}

pub unsafe fn mark_dirty_from_edit(hwnd: HWND, hwnd_edit: HWND) {
    let _ = with_state(hwnd, |state| {
        for (i, doc) in state.docs.iter_mut().enumerate() {
            if doc.hwnd_edit == hwnd_edit && !doc.dirty {
                doc.dirty = true;
                update_tab_title(state.hwnd_tab, i, &doc.title, true);
                update_window_title(hwnd);
                break;
            }
        }
    });
}

pub unsafe fn update_window_title(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(state.current) {
            let untitled = untitled_base(state.settings.language);
            let display_title = if doc.title.starts_with(&untitled) {
                &doc.title
            } else {
                &doc.title
            };
            let base_title = if display_title.trim().is_empty() {
                "Novapad".to_string()
            } else {
                format!("{display_title} - Novapad")
            };
            let full_title = apply_modified_marker(
                &base_title,
                doc.dirty,
                state.settings.modified_marker_position,
            );
            let wide = to_wide(&full_title);
            let _ = SetWindowTextW(hwnd, PCWSTR(wide.as_ptr()));
        }
    });
}

fn apply_modified_marker(title: &str, dirty: bool, position: ModifiedMarkerPosition) -> String {
    if !dirty {
        return title.to_string();
    }
    match position {
        ModifiedMarkerPosition::Beginning => format!("* {title}"),
        _ => format!("{title} *"),
    }
}

pub unsafe fn layout_children(hwnd: HWND) {
    let state_data = with_state(hwnd, |state| {
        (
            state.hwnd_tab,
            state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>(),
            state.voice_panel_visible,
            state.voice_favorites_visible,
            state.settings.tts_engine,
            state.voice_label_engine,
            state.voice_combo_engine,
            state.voice_label_voice,
            state.voice_combo_voice,
            state.voice_checkbox_multilingual,
            state.voice_label_favorites,
            state.voice_combo_favorites,
        )
    });

    let Some((
        hwnd_tab,
        edit_handles,
        voice_panel_visible,
        favorites_visible,
        tts_engine,
        label_engine,
        combo_engine,
        label_voice,
        combo_voice,
        checkbox_multilingual,
        label_favorites,
        combo_favorites,
    )) = state_data
    else {
        return;
    };

    let mut rc = RECT::default();
    if GetClientRect(hwnd, &mut rc).is_err() {
        return;
    }

    let width = rc.right - rc.left;
    let height = rc.bottom - rc.top;

    let _ = MoveWindow(hwnd_tab, 0, 0, width, height, true);

    let mut tab_rc = rc;
    SendMessageW(
        hwnd_tab,
        TCM_ADJUSTRECT,
        WPARAM(0),
        LPARAM(&mut tab_rc as *mut _ as isize),
    );

    let mut panel_height = 0;
    let panel_visible = voice_panel_visible || favorites_visible;
    if panel_visible {
        let show_multilingual =
            voice_panel_visible && matches!(tts_engine, crate::settings::TtsEngine::Edge);
        let mut rows = 0;
        if voice_panel_visible {
            rows += 2;
            if show_multilingual {
                rows += 1;
            }
        }
        if favorites_visible {
            rows += 1;
        }
        panel_height = VOICE_PANEL_PADDING * 2
            + VOICE_PANEL_ROW_HEIGHT * rows
            + VOICE_PANEL_SPACING * (rows - 1);
        let label_x = tab_rc.left + VOICE_PANEL_PADDING;
        let combo_x = label_x + VOICE_PANEL_LABEL_WIDTH + VOICE_PANEL_PADDING;
        let combo_width = (tab_rc.right - VOICE_PANEL_PADDING) - combo_x;
        let combo_width = if combo_width < 120 { 120 } else { combo_width };
        let row1_top = tab_rc.top + VOICE_PANEL_PADDING;
        let row2_top = row1_top + VOICE_PANEL_ROW_HEIGHT + VOICE_PANEL_SPACING;
        let row3_top = row2_top + VOICE_PANEL_ROW_HEIGHT + VOICE_PANEL_SPACING;
        let row4_top = row3_top + VOICE_PANEL_ROW_HEIGHT + VOICE_PANEL_SPACING;

        if voice_panel_visible {
            let _ = MoveWindow(
                label_engine,
                label_x,
                row1_top,
                VOICE_PANEL_LABEL_WIDTH,
                VOICE_PANEL_ROW_HEIGHT,
                true,
            );
            let _ = MoveWindow(
                combo_engine,
                combo_x,
                row1_top - 2,
                combo_width,
                VOICE_PANEL_COMBO_HEIGHT,
                true,
            );
            let _ = MoveWindow(
                label_voice,
                label_x,
                row2_top,
                VOICE_PANEL_LABEL_WIDTH,
                VOICE_PANEL_ROW_HEIGHT,
                true,
            );
            let _ = MoveWindow(
                combo_voice,
                combo_x,
                row2_top - 2,
                combo_width,
                VOICE_PANEL_COMBO_HEIGHT,
                true,
            );
            if show_multilingual {
                let _ = MoveWindow(
                    checkbox_multilingual,
                    label_x,
                    row3_top,
                    combo_width + VOICE_PANEL_LABEL_WIDTH + VOICE_PANEL_PADDING,
                    VOICE_PANEL_ROW_HEIGHT,
                    true,
                );
                if favorites_visible {
                    let _ = MoveWindow(
                        label_favorites,
                        label_x,
                        row4_top,
                        VOICE_PANEL_LABEL_WIDTH,
                        VOICE_PANEL_ROW_HEIGHT,
                        true,
                    );
                    let _ = MoveWindow(
                        combo_favorites,
                        combo_x,
                        row4_top - 2,
                        combo_width,
                        VOICE_PANEL_COMBO_HEIGHT,
                        true,
                    );
                }
            } else if favorites_visible {
                let _ = MoveWindow(
                    label_favorites,
                    label_x,
                    row3_top,
                    VOICE_PANEL_LABEL_WIDTH,
                    VOICE_PANEL_ROW_HEIGHT,
                    true,
                );
                let _ = MoveWindow(
                    combo_favorites,
                    combo_x,
                    row3_top - 2,
                    combo_width,
                    VOICE_PANEL_COMBO_HEIGHT,
                    true,
                );
            }
        } else if favorites_visible {
            let _ = MoveWindow(
                label_favorites,
                label_x,
                row1_top,
                VOICE_PANEL_LABEL_WIDTH,
                VOICE_PANEL_ROW_HEIGHT,
                true,
            );
            let _ = MoveWindow(
                combo_favorites,
                combo_x,
                row1_top - 2,
                combo_width,
                VOICE_PANEL_COMBO_HEIGHT,
                true,
            );
        }
    }

    let panel_offset = panel_height;
    for hwnd_edit in edit_handles {
        if hwnd_edit.0 != 0 {
            let _ = MoveWindow(
                hwnd_edit,
                tab_rc.left,
                tab_rc.top + panel_offset,
                tab_rc.right - tab_rc.left,
                tab_rc.bottom - tab_rc.top - panel_offset,
                true,
            );
        }
    }
}

pub unsafe fn create_edit(
    parent: HWND,
    hfont: HFONT,
    word_wrap: bool,
    text_color: u32,
    text_size: i32,
) -> HWND {
    let mut style = WS_CHILD
        | WS_CLIPCHILDREN
        | WS_VSCROLL
        | WS_GROUP
        | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_MULTILINE as u32)
        | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_AUTOVSCROLL as u32)
        | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_WANTRETURN as u32);
    if !word_wrap {
        style |= WS_HSCROLL
            | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_AUTOHSCROLL as u32);
    }

    let hwnd_edit = windows::Win32::UI::WindowsAndMessaging::CreateWindowExW(
        WS_EX_CLIENTEDGE,
        MSFTEDIT_CLASS,
        PCWSTR::null(),
        style,
        0,
        0,
        0,
        0,
        parent,
        HMENU(0),
        HINSTANCE(0),
        None,
    );

    if hwnd_edit.0 != 0 {
        if hfont.0 != 0 {
            SendMessageW(hwnd_edit, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }
        // Allow large pastes (default edit limit is ~32K).
        apply_text_limit(hwnd_edit);
        apply_text_appearance(hwnd_edit, text_color, text_size);
        SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
        SendMessageW(
            hwnd_edit,
            EM_SETEVENTMASK,
            WPARAM(0),
            LPARAM(ENM_CHANGE as isize),
        );
    }
    hwnd_edit
}

pub unsafe fn save_current_document(hwnd: HWND) -> bool {
    save_document_at(hwnd, get_current_index(hwnd), false)
}

pub unsafe fn save_current_document_as(hwnd: HWND) -> bool {
    save_document_at(hwnd, get_current_index(hwnd), true)
}

pub unsafe fn save_all_documents(hwnd: HWND) -> bool {
    let dirty_indices = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .enumerate()
            .filter_map(|(i, doc)| if doc.dirty { Some(i) } else { None })
            .collect::<Vec<_>>()
    })
    .unwrap_or_default();
    for index in dirty_indices {
        if !save_document_at(hwnd, index, false) {
            return false;
        }
    }
    true
}

pub unsafe fn save_document_at(hwnd: HWND, index: usize, force_dialog: bool) -> bool {
    let result = with_state(hwnd, |state| {
        if state.docs.is_empty() || index >= state.docs.len() {
            return None;
        }
        let language = state.settings.language;
        let text = get_edit_text(state.docs[index].hwnd_edit);
        let is_lossy_doc = matches!(
            state.docs[index].format,
            FileFormat::Docx
                | FileFormat::Doc
                | FileFormat::Pdf
                | FileFormat::Spreadsheet
                | FileFormat::Epub
                | FileFormat::Html
                | FileFormat::Ppt
                | FileFormat::Pptx
        );
        let mut suggested_name = crate::suggested_filename_from_text(&text)
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| state.docs[index].title.clone());
        if is_lossy_doc {
            let mut name_path = PathBuf::from(&suggested_name);
            name_path.set_extension("txt");
            suggested_name = name_path
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or("document.txt")
                .to_string();
        }

        let path_info = if !force_dialog && !is_lossy_doc {
            state.docs[index].path.clone().map(|p| (p, None))
        } else {
            None
        };

        let (path, user_selected_encoding) = match path_info {
            Some((path, enc)) => (path, enc),
            None => {
                let initial_encoding = state.docs[index]
                    .current_save_text_encoding
                    .or(state.docs[index].opened_text_encoding)
                    .unwrap_or(state.settings.default_text_encoding);
                match crate::save_file_dialog_with_encoding(
                    hwnd,
                    Some(&suggested_name),
                    initial_encoding,
                ) {
                    Some((path, enc)) => (path, Some(enc)),
                    None => return None,
                }
            }
        };

        let mut path = path;
        if is_lossy_doc {
            path.set_extension("txt");
        }

        let is_docx = is_docx_path(&path);
        let is_pdf = is_pdf_path(&path);
        if !is_lossy_doc && is_docx {
            if let Err(message) = write_docx_text(&path, &text, language) {
                crate::show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Docx;
        } else if !is_lossy_doc && is_pdf {
            let pdf_title = path
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("Documento");
            if let Err(message) = write_pdf_text(&path, pdf_title, &text, language) {
                crate::show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Pdf;
        } else {
            let encoding = if let Some(enc) = user_selected_encoding {
                state.docs[index].current_save_text_encoding = Some(enc);
                enc
            } else {
                state.docs[index]
                    .current_save_text_encoding
                    .or(state.docs[index].opened_text_encoding)
                    .unwrap_or(state.settings.default_text_encoding)
            };
            let bytes = encode_text(&text, encoding);
            if let Err(err) = std::fs::write(&path, bytes) {
                crate::show_error(
                    hwnd,
                    language,
                    &crate::settings::error_save_file_message(language, err),
                );
                return None;
            }
            state.docs[index].format = FileFormat::Text(encoding);
        }

        let hwnd_edit = state.docs[index].hwnd_edit;
        state.docs[index].path = Some(path.clone());
        state.docs[index].dirty = false;
        SendMessageW(hwnd_edit, EM_SETMODIFY, WPARAM(0), LPARAM(0));
        let title = path.file_name().and_then(|s| s.to_str()).unwrap_or("File");
        state.docs[index].title = title.to_string();
        update_tab_title(state.hwnd_tab, index, &state.docs[index].title, false);
        if index == state.current {
            update_window_title(hwnd);
        }
        Some(path)
    });

    if let Some(Some(path)) = result {
        crate::push_recent_file(hwnd, &path);
        true
    } else {
        false
    }
}

pub unsafe fn close_current_document(hwnd: HWND) {
    let index = match with_state(hwnd, |state| state.current) {
        Some(i) => i,
        None => return,
    };
    let _ = close_document_at(hwnd, index);
}

pub unsafe fn close_document_at(hwnd: HWND, index: usize) -> bool {
    let result = with_state(hwnd, |state| {
        if index >= state.docs.len() {
            return None;
        }
        Some((
            state.current,
            state.hwnd_tab,
            state.docs.len(),
            state.docs[index].title.clone(),
        ))
    });

    let (_current, hwnd_tab, _count, title) = match result {
        Some(Some(values)) => values,
        _ => return true,
    };

    if !confirm_save_if_dirty_entry(hwnd, index, &title) {
        return false;
    }

    let mut closing_hwnd_edit = HWND(0);
    let mut new_hwnd_edit = None;
    let mut was_current = false;
    let mut was_empty = false;
    let mut update_title = false;

    let _ = with_state(hwnd, |state| {
        was_current = state.current == index;
        let doc = state.docs.remove(index);
        closing_hwnd_edit = doc.hwnd_edit;
        let _ = SendMessageW(
            hwnd_tab,
            windows::Win32::UI::Controls::TCM_DELETEITEM,
            WPARAM(index),
            LPARAM(0),
        );

        if state.docs.is_empty() {
            state.untitled_count = 0;
            state.current = 0;
            was_empty = true;
        } else {
            if was_current {
                let idx = if index >= state.docs.len() {
                    state.docs.len() - 1
                } else {
                    index
                };
                state.current = idx;
                SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(idx), LPARAM(0));
                new_hwnd_edit = state.docs.get(idx).map(|doc| doc.hwnd_edit);
                update_title = true;
            } else if index < state.current {
                state.current -= 1;
                SendMessageW(hwnd_tab, TCM_SETCURSEL, WPARAM(state.current), LPARAM(0));
            }
        }
    });

    if closing_hwnd_edit.0 != 0 {
        let _ = DestroyWindow(closing_hwnd_edit);
    }

    if was_empty {
        new_document(hwnd);
    } else {
        if let Some(hwnd_edit) = new_hwnd_edit {
            let is_audiobook = with_state(hwnd, |state| {
                state
                    .docs
                    .get(state.current)
                    .map(|d| matches!(d.format, FileFormat::Audiobook))
                    .unwrap_or(false)
            })
            .unwrap_or(false);
            if is_audiobook {
                let hwnd_tab = with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0));
                if hwnd_tab.0 != 0 {
                    SetFocus(hwnd_tab);
                }
            } else {
                ShowWindow(hwnd_edit, SW_SHOW);
                SetFocus(hwnd_edit);
            }
        }
        if update_title {
            update_window_title(hwnd);
        }
    }
    layout_children(hwnd);
    true
}

pub unsafe fn try_close_app(hwnd: HWND) -> bool {
    let result = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .enumerate()
            .map(|(i, d)| (i, d.title.clone()))
            .collect::<Vec<_>>()
    });

    if let Some(entries) = result {
        for (index, title) in entries {
            if !confirm_save_if_dirty_entry(hwnd, index, &title) {
                return false;
            }
        }
    }
    let _ = DestroyWindow(hwnd);
    true
}

pub unsafe fn sync_dirty_from_edit(hwnd: HWND, index: usize) -> bool {
    let mut hwnd_edit = HWND(0);
    let mut is_dirty = false;
    let mut is_current = false;
    let _ = with_state(hwnd, |state| {
        if let Some(doc) = state.docs.get(index) {
            hwnd_edit = doc.hwnd_edit;
            is_dirty = doc.dirty;
            is_current = state.current == index;
        }
    });

    if hwnd_edit.0 == 0 {
        return is_dirty;
    }

    let modified = SendMessageW(hwnd_edit, EM_GETMODIFY, WPARAM(0), LPARAM(0)).0 != 0;
    if modified && !is_dirty {
        let _ = with_state(hwnd, |state| {
            if let Some(doc) = state.docs.get_mut(index) {
                doc.dirty = true;
                update_tab_title(state.hwnd_tab, index, &doc.title, true);
                if is_current {
                    update_window_title(hwnd);
                }
            }
        });
        return true;
    }
    is_dirty
}

pub unsafe fn confirm_save_if_dirty_entry(hwnd: HWND, index: usize, title: &str) -> bool {
    if !sync_dirty_from_edit(hwnd, index) {
        return true;
    }

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let msg = confirm_save_message(language, title);
    let title_w = confirm_title(language);

    let result = MessageBoxW(
        hwnd,
        PCWSTR(to_wide(&msg).as_ptr()),
        PCWSTR(to_wide(&title_w).as_ptr()),
        MB_YESNOCANCEL | MB_ICONWARNING,
    );

    match result {
        IDYES => save_document_at(hwnd, index, false),
        IDNO => true,
        _ => false,
    }
}

pub unsafe fn get_current_index(hwnd: HWND) -> usize {
    with_state(hwnd, |state| state.current).unwrap_or(0)
}

pub unsafe fn get_tab(hwnd: HWND) -> HWND {
    with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_clean_end_of_line_hyphens() {
        // Simple join
        assert_eq!(
            clean_end_of_line_hyphens_block("inter-\nnational", "\n", false),
            "international"
        );
        // Windows EOL
        assert_eq!(
            clean_end_of_line_hyphens_block("inter-\r\nnational", "\r\n", false),
            "international"
        );
        // Page gap (2 line breaks)
        assert_eq!(
            clean_end_of_line_hyphens_block("inter-\n\nnational", "\n", false),
            "international"
        );
        // 3 line breaks
        assert_eq!(
            clean_end_of_line_hyphens_block("inter-\n\n\nnational", "\n", false),
            "international"
        );
        // 4 line breaks (too many, shouldn't join)
        assert_eq!(
            clean_end_of_line_hyphens_block("inter-\n\n\n\nnational", "\n", false),
            "inter-\n\n\n\nnational"
        );

        // Non-join for dashes (whitespace before hyphen)
        assert_eq!(
            clean_end_of_line_hyphens_block("word -\nnext", "\n", false),
            "word -\nnext"
        );

        // Non-join for hyphenated compounds (next char not alphabetic)
        assert_eq!(
            clean_end_of_line_hyphens_block("state-\n123", "\n", false),
            "state-\n123"
        );
        assert_eq!(
            clean_end_of_line_hyphens_block("state-\n. Next", "\n", false),
            "state-\n. Next"
        );

        // Digit before hyphen (should join)
        assert_eq!(
            clean_end_of_line_hyphens_block("Section3-\npart", "\n", false),
            "Section3part"
        );

        // Preserve paragraph structure
        let input = "This is a test-\ncase.\n\nNew paragraph.";
        let expected = "This is a testcase.\n\nNew paragraph.";
        assert_eq!(
            clean_end_of_line_hyphens_block(input, "\n", false),
            expected
        );

        // Case sensitivity (lowercase/uppercase after join)
        assert_eq!(
            clean_end_of_line_hyphens_block("Inter-\nNational", "\n", false),
            "InterNational"
        );
    }
}
