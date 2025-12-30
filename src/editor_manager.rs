use std::path::{Path, PathBuf};
use windows::core::{PCWSTR, PWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM, RECT, HINSTANCE};
use windows::Win32::UI::WindowsAndMessaging::{
    ShowWindow, SW_HIDE, SW_SHOW, SendMessageW, GetClientRect, MoveWindow, SetWindowTextW,
    WS_CHILD, WS_CLIPCHILDREN, WS_EX_CLIENTEDGE, WS_VSCROLL, WS_HSCROLL, HMENU,
    WM_SETFONT, IDYES, IDNO, MessageBoxW, MB_YESNOCANCEL, MB_ICONWARNING,
    GetWindowTextLengthW, GetWindowTextW, DestroyWindow, WS_GROUP,
    ES_MULTILINE, ES_AUTOVSCROLL, ES_AUTOHSCROLL, ES_WANTRETURN
};
use windows::Win32::UI::Controls::{
    TCITEMW, TCIF_TEXT, TCM_INSERTITEMW, TCM_SETITEMW, TCM_SETCURSEL, TCM_ADJUSTRECT,
    EM_GETMODIFY, EM_SETMODIFY, EM_SETREADONLY
};
use windows::Win32::UI::Controls::RichEdit::{
    CHARFORMAT2W, CFM_COLOR, CFM_SIZE, EM_SETCHARFORMAT, ENM_CHANGE, MSFTEDIT_CLASS, SCF_ALL,
    CHARRANGE, EM_EXSETSEL, EM_SETEVENTMASK
};
use windows::Win32::UI::Input::KeyboardAndMouse::{SetFocus};
use windows::Win32::Graphics::Gdi::{HFONT};
use crate::settings::{FileFormat, TextEncoding, untitled_title, untitled_base, confirm_save_message, confirm_title};
use crate::accessibility::{to_wide, from_wide, to_wide_normalized};
use crate::{with_state, log_debug};
use crate::file_handler::*;

const EM_LIMITTEXT: u32 = 0x00C5;
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
    }).unwrap_or_default();

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
}

impl Default for Document {
    fn default() -> Self {
        Document {
            title: String::new(),
            path: None,
            hwnd_edit: HWND(0),
            dirty: false,
            format: FileFormat::Text(TextEncoding::Utf8),
        }
    }
}

// --- Editor Helpers ---

pub unsafe fn set_edit_text(hwnd_edit: HWND, text: &str) {
    let wide = to_wide_normalized(text);
    let _ = SetWindowTextW(hwnd_edit, PCWSTR(wide.as_ptr()));
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
        let cr = CHARRANGE { cpMin: 0, cpMax: -1 };
        SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&cr as *const _ as isize));
    }
}

pub unsafe fn apply_word_wrap_to_all_edits(hwnd: HWND, word_wrap: bool) {
    let edits = with_state(hwnd, |state| {
        state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>()
    }).unwrap_or_default();

    for hwnd_edit in edits {
        if hwnd_edit.0 == 0 { continue; }
        log_debug(&format!("Word wrap toggle for {:?}: {}", hwnd_edit, word_wrap));
        apply_text_limit(hwnd_edit);
    }
}

pub unsafe fn apply_text_appearance_to_all_edits(hwnd: HWND, text_color: u32, text_size: i32) {
    let edits = with_state(hwnd, |state| {
        state.docs.iter().map(|d| d.hwnd_edit).collect::<Vec<_>>()
    }).unwrap_or_default();

    for hwnd_edit in edits {
        if hwnd_edit.0 == 0 { continue; }
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
        SendMessageW(hwnd_edit, EM_SETCHARFORMAT, WPARAM(SCF_ALL as usize), LPARAM(&mut format as *mut _ as isize));
    }
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
        };
        state.docs.push(doc);
        insert_tab(state.hwnd_tab, &title, (state.docs.len() - 1) as i32);
        state.docs.len() - 1
    })
    .unwrap_or(0);
    select_tab(hwnd, new_index);
}

pub unsafe fn open_document(hwnd: HWND, path: &Path) {
    log_debug(&format!("Open document: {}", path.display()));

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if is_pdf_path(path) {
        crate::open_pdf_document_async(hwnd, path);
        return;
    }
    let (content, format) = if is_docx_path(path) {
        match read_docx_text(path, language) {
            Ok(text) => (text, FileFormat::Docx),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_epub_path(path) {
        match read_epub_text(path, language) {
            Ok(text) => (text, FileFormat::Epub),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_mp3_path(path) {
        (String::new(), FileFormat::Audiobook)
    } else if is_doc_path(path) {
        match read_doc_text(path, language) {
            Ok(text) => (text, FileFormat::Doc),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else if is_spreadsheet_path(path) {
        match read_spreadsheet_text(path, language) {
            Ok(text) => (text, FileFormat::Spreadsheet),
            Err(message) => {
                crate::show_error(hwnd, language, &message);
                return;
            }
        }
    } else {
        match std::fs::read(path) {
            Ok(bytes) => match decode_text(&bytes, language) {
                Ok((text, encoding)) => (text, FileFormat::Text(encoding)),
                Err(message) => {
                    crate::show_error(hwnd, language, &message);
                    return;
                }
            },
            Err(err) => {
                crate::show_error(hwnd, language, &crate::settings::error_open_file_message(language, err));
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
        unsafe { crate::audio_player::start_audiobook_playback(hwnd, path); }
    }
    crate::push_recent_file(hwnd, path);
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
        let is_audiobook = new_doc.map(|doc| matches!(doc.format, FileFormat::Audiobook)).unwrap_or(false);
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
    SendMessageW(hwnd_tab, TCM_INSERTITEMW, WPARAM(index as usize), LPARAM(&mut item as *mut _ as isize));
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
    SendMessageW(hwnd_tab, TCM_SETITEMW, WPARAM(index), LPARAM(&mut item as *mut _ as isize));
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
            let suffix = if doc.dirty { "*" } else { "" };
            let untitled = untitled_base(state.settings.language);
            let display_title = if doc.title.starts_with(untitled) {
                &doc.title
            } else {
                &doc.title
            };
            let full_title = format!("{} - Novapad{}", display_title, suffix);
            let wide = to_wide(&full_title);
            let _ = SetWindowTextW(hwnd, PCWSTR(wide.as_ptr()));
        }
    });
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

    let Some((hwnd_tab, edit_handles, voice_panel_visible, favorites_visible, tts_engine, label_engine, combo_engine, label_voice, combo_voice, checkbox_multilingual, label_favorites, combo_favorites)) = state_data else {
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
    SendMessageW(hwnd_tab, TCM_ADJUSTRECT, WPARAM(0), LPARAM(&mut tab_rc as *mut _ as isize));

    let mut panel_height = 0;
    let panel_visible = voice_panel_visible || favorites_visible;
    if panel_visible {
        let show_multilingual = voice_panel_visible && matches!(tts_engine, crate::settings::TtsEngine::Edge);
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
            let _ = MoveWindow(label_engine, label_x, row1_top, VOICE_PANEL_LABEL_WIDTH, VOICE_PANEL_ROW_HEIGHT, true);
            let _ = MoveWindow(combo_engine, combo_x, row1_top - 2, combo_width, VOICE_PANEL_COMBO_HEIGHT, true);
            let _ = MoveWindow(label_voice, label_x, row2_top, VOICE_PANEL_LABEL_WIDTH, VOICE_PANEL_ROW_HEIGHT, true);
            let _ = MoveWindow(combo_voice, combo_x, row2_top - 2, combo_width, VOICE_PANEL_COMBO_HEIGHT, true);
            if show_multilingual {
                let _ = MoveWindow(checkbox_multilingual, label_x, row3_top, combo_width + VOICE_PANEL_LABEL_WIDTH + VOICE_PANEL_PADDING, VOICE_PANEL_ROW_HEIGHT, true);
                if favorites_visible {
                    let _ = MoveWindow(label_favorites, label_x, row4_top, VOICE_PANEL_LABEL_WIDTH, VOICE_PANEL_ROW_HEIGHT, true);
                    let _ = MoveWindow(combo_favorites, combo_x, row4_top - 2, combo_width, VOICE_PANEL_COMBO_HEIGHT, true);
                }
            } else if favorites_visible {
                let _ = MoveWindow(label_favorites, label_x, row3_top, VOICE_PANEL_LABEL_WIDTH, VOICE_PANEL_ROW_HEIGHT, true);
                let _ = MoveWindow(combo_favorites, combo_x, row3_top - 2, combo_width, VOICE_PANEL_COMBO_HEIGHT, true);
            }
        } else if favorites_visible {
            let _ = MoveWindow(label_favorites, label_x, row1_top, VOICE_PANEL_LABEL_WIDTH, VOICE_PANEL_ROW_HEIGHT, true);
            let _ = MoveWindow(combo_favorites, combo_x, row1_top - 2, combo_width, VOICE_PANEL_COMBO_HEIGHT, true);
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

pub unsafe fn create_edit(parent: HWND, hfont: HFONT, word_wrap: bool, text_color: u32, text_size: i32) -> HWND {
    let mut style = WS_CHILD | WS_CLIPCHILDREN | WS_VSCROLL | WS_GROUP | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_MULTILINE as u32) | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_AUTOVSCROLL as u32) | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_WANTRETURN as u32);
    if !word_wrap {
        style |= WS_HSCROLL | windows::Win32::UI::WindowsAndMessaging::WINDOW_STYLE(ES_AUTOHSCROLL as u32);
    }

    let hwnd_edit = windows::Win32::UI::WindowsAndMessaging::CreateWindowExW(
        WS_EX_CLIENTEDGE,
        MSFTEDIT_CLASS,
        PCWSTR::null(),
        style,
        0, 0, 0, 0,
        parent,
        HMENU(0),
        HINSTANCE(0),
        None,
    );

    if hwnd_edit.0 != 0 {
        if hfont.0 != 0 {
            SendMessageW(hwnd_edit, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }
        SendMessageW(hwnd_edit, EM_SETEVENTMASK, WPARAM(0), LPARAM(ENM_CHANGE as isize));
        // Allow large pastes (default edit limit is ~32K).
        apply_text_limit(hwnd_edit);
        apply_text_appearance(hwnd_edit, text_color, text_size);
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
        let suggested_name = crate::suggested_filename_from_text(&text)
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| state.docs[index].title.clone());

        let path = if !force_dialog {
            state.docs[index].path.clone()
        } else {
            None
        };
        let path = match path {
            Some(path) => path,
            None => match crate::save_file_dialog(hwnd, Some(&suggested_name)) {
                Some(path) => path,
                None => return None,
            },
        };

        let is_docx = is_docx_path(&path);
        let is_pdf = is_pdf_path(&path);
        if is_docx {
            if let Err(message) = write_docx_text(&path, &text, language) {
                crate::show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Docx;
        } else if is_pdf {
            let pdf_title = path.file_stem().and_then(|s| s.to_str()).unwrap_or("Documento");
            if let Err(message) = write_pdf_text(&path, pdf_title, &text, language) {
                crate::show_error(hwnd, language, &message);
                return None;
            }
            state.docs[index].format = FileFormat::Pdf;
        } else {
            let encoding = match state.docs[index].format {
                FileFormat::Text(enc) => enc,
                FileFormat::Docx | FileFormat::Doc | FileFormat::Pdf | FileFormat::Spreadsheet | FileFormat::Epub | FileFormat::Audiobook => TextEncoding::Utf8,
            };
            let bytes = encode_text(&text, encoding);
            if let Err(err) = std::fs::write(&path, bytes) {
                crate::show_error(hwnd, language, &crate::settings::error_save_file_message(language, err));
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
        let _ = SendMessageW(hwnd_tab, windows::Win32::UI::Controls::TCM_DELETEITEM, WPARAM(index), LPARAM(0));

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
                state.docs.get(state.current).map(|d| matches!(d.format, FileFormat::Audiobook)).unwrap_or(false)
            }).unwrap_or(false);
            if is_audiobook {
                let hwnd_tab = with_state(hwnd, |state| state.hwnd_tab).unwrap_or(HWND(0));
                if hwnd_tab.0 != 0 { SetFocus(hwnd_tab); }
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
        state.docs.iter().enumerate()
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
        PCWSTR(to_wide(title_w).as_ptr()),
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
