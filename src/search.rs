use crate::accessibility::{EM_REPLACESEL, EM_SCROLLCARET, from_wide, to_wide};
use crate::editor_manager::get_edit_text;
use crate::i18n;
use crate::settings::{Language, find_title, text_not_found_message};
use crate::{get_active_edit, show_error, show_info, with_state};
use fancy_regex::Regex;
use windows::Win32::Foundation::HINSTANCE;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::Graphics::Gdi::{DEFAULT_GUI_FONT, GetStockObject};
use windows::Win32::UI::Controls::Dialogs::{
    FINDREPLACE_FLAGS, FINDREPLACEW, FR_DIALOGTERM, FR_DOWN, FR_ENABLEHOOK, FR_FINDNEXT,
    FR_MATCHCASE, FR_REPLACE, FR_REPLACEALL, FR_WHOLEWORD, FindTextW, ReplaceTextW,
};
use windows::Win32::UI::Controls::RichEdit::{
    CHARRANGE, EM_EXGETSEL, EM_EXSETSEL, EM_FINDTEXTEXW, FINDTEXTEXW,
};
use windows::Win32::UI::Controls::{BST_CHECKED, WC_BUTTON};
use windows::Win32::UI::Input::KeyboardAndMouse::SetFocus;
use windows::Win32::UI::WindowsAndMessaging::{
    BM_GETCHECK, BM_SETCHECK, BS_AUTOCHECKBOX, CreateWindowExW, GetClientRect, GetParent,
    GetWindowRect, HMENU, MB_ICONWARNING, MB_OK, MessageBoxW, SWP_NOMOVE, SWP_NOZORDER,
    SendMessageW, SetWindowPos, WINDOW_STYLE, WM_COMMAND, WM_INITDIALOG, WM_SETFONT,
};
use windows::core::{PCWSTR, PWSTR};

pub const FIND_DIALOG_ID: isize = 1;
pub const REPLACE_DIALOG_ID: isize = 2;
const FIND_ID_REGEX: isize = 5101;
const FIND_ID_DOT_MATCHES_NEWLINE: isize = 5102;
const FIND_ID_WRAP_AROUND: isize = 5103;
const REPLACE_ID_IN_SELECTION: isize = 5104;
const REPLACE_ID_IN_ALL_DOCS: isize = 5105;

#[derive(Copy, Clone)]
pub struct FindOptions {
    pub use_regex: bool,
    pub dot_matches_newline: bool,
    pub wrap_around: bool,
    pub replace_in_selection: bool,
    pub replace_in_all_docs: bool,
}

pub unsafe fn open_find_dialog(hwnd: HWND) {
    let has_dialog = with_state(hwnd, |state| state.find_dialog.0 != 0).unwrap_or(false);
    if has_dialog {
        let _ = with_state(hwnd, |state| {
            SetFocus(state.find_dialog);
        });
        return;
    }

    let _ = with_state(hwnd, |state| {
        state.find_replace = Some(FINDREPLACEW {
            lStructSize: std::mem::size_of::<FINDREPLACEW>() as u32,
            hwndOwner: hwnd,
            Flags: FR_DOWN | FR_ENABLEHOOK,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lCustData: LPARAM(FIND_DIALOG_ID),
            lpfnHook: Some(find_replace_hook_proc),
            ..Default::default()
        });

        if let Some(ref mut fr) = state.find_replace {
            let dialog = FindTextW(fr);
            state.find_dialog = dialog;
        }
    });
}

pub unsafe fn open_replace_dialog(hwnd: HWND) {
    let has_dialog = with_state(hwnd, |state| state.replace_dialog.0 != 0).unwrap_or(false);
    if has_dialog {
        let _ = with_state(hwnd, |state| {
            SetFocus(state.replace_dialog);
        });
        return;
    }

    let _ = with_state(hwnd, |state| {
        state.replace_replace = Some(FINDREPLACEW {
            lStructSize: std::mem::size_of::<FINDREPLACEW>() as u32,
            hwndOwner: hwnd,
            Flags: FR_DOWN | FR_ENABLEHOOK,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lpstrReplaceWith: PWSTR(state.replace_text.as_mut_ptr()),
            wReplaceWithLen: state.replace_text.len() as u16,
            lCustData: LPARAM(REPLACE_DIALOG_ID),
            lpfnHook: Some(find_replace_hook_proc),
            ..Default::default()
        });

        if let Some(ref mut fr) = state.replace_replace {
            let dialog = ReplaceTextW(fr);
            state.replace_dialog = dialog;
        }
    });
}

pub unsafe fn handle_find_message(hwnd: HWND, lparam: LPARAM) {
    let fr = &*(lparam.0 as *const FINDREPLACEW);
    if (fr.Flags & FR_DIALOGTERM) != FINDREPLACE_FLAGS(0) {
        let _ = with_state(hwnd, |state| {
            if fr.lCustData.0 == FIND_DIALOG_ID {
                state.find_dialog = HWND(0);
                state.find_replace = None;
            } else if fr.lCustData.0 == REPLACE_DIALOG_ID {
                state.replace_dialog = HWND(0);
                state.replace_replace = None;
            }
        });
        return;
    }

    if (fr.Flags & (FR_FINDNEXT | FR_REPLACE | FR_REPLACEALL)) == FINDREPLACE_FLAGS(0) {
        return;
    }

    let search = from_wide(fr.lpstrFindWhat.0);
    if search.is_empty() {
        return;
    }

    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();

    let find_flags = extract_find_flags(fr.Flags);
    let options = get_find_options(hwnd);
    let _ = with_state(hwnd, |state| {
        state.last_find_flags = find_flags;
    });

    if (fr.Flags & FR_REPLACEALL) != FINDREPLACE_FLAGS(0) {
        replace_all(
            hwnd,
            hwnd_edit,
            &search,
            &from_wide(fr.lpstrReplaceWith.0),
            find_flags,
            &options,
        );
        return;
    }

    if (fr.Flags & FR_REPLACE) != FINDREPLACE_FLAGS(0) {
        let replace = from_wide(fr.lpstrReplaceWith.0);
        let replaced =
            replace_selection_if_match(hwnd, hwnd_edit, &search, &replace, find_flags, &options);
        let found = find_next(
            hwnd,
            hwnd_edit,
            &search,
            find_flags,
            options.wrap_around,
            &options,
        );
        if !replaced && !found {
            let message = to_wide(&text_not_found_message(language));
            let title = to_wide(&find_title(language));
            MessageBoxW(
                hwnd,
                PCWSTR(message.as_ptr()),
                PCWSTR(title.as_ptr()),
                MB_OK | MB_ICONWARNING,
            );
        }
        return;
    }

    if find_next(
        hwnd,
        hwnd_edit,
        &search,
        find_flags,
        options.wrap_around,
        &options,
    ) {
        return;
    }
    let message = to_wide(&text_not_found_message(language));
    let title = to_wide(&find_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONWARNING,
    );
}

pub unsafe fn find_next_from_state(hwnd: HWND) {
    let (search, flags, language): (String, FINDREPLACE_FLAGS, Language) =
        with_state(hwnd, |state| {
            let search = from_wide(state.find_text.as_ptr());
            (search, state.last_find_flags, state.settings.language)
        })
        .unwrap_or((String::new(), FINDREPLACE_FLAGS(0), Language::default()));
    if search.is_empty() {
        open_find_dialog(hwnd);
        return;
    }
    let Some(hwnd_edit) = get_active_edit(hwnd) else {
        return;
    };
    let options = get_find_options(hwnd);
    if !find_next(
        hwnd,
        hwnd_edit,
        &search,
        flags,
        options.wrap_around,
        &options,
    ) {
        let message = to_wide(&text_not_found_message(language));
        let title = to_wide(&find_title(language));
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    }
}

fn extract_find_flags(flags: FINDREPLACE_FLAGS) -> FINDREPLACE_FLAGS {
    let mut out = FINDREPLACE_FLAGS(0);
    if (flags & FR_MATCHCASE) != FINDREPLACE_FLAGS(0) {
        out |= FR_MATCHCASE;
    }
    if (flags & FR_WHOLEWORD) != FINDREPLACE_FLAGS(0) {
        out |= FR_WHOLEWORD;
    }
    if (flags & FR_DOWN) != FINDREPLACE_FLAGS(0) {
        out |= FR_DOWN;
    }
    out
}

pub unsafe fn find_next(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    flags: FINDREPLACE_FLAGS,
    wrap: bool,
    options: &FindOptions,
) -> bool {
    if options.use_regex {
        return find_next_regex(hwnd, hwnd_edit, search, flags, wrap, options);
    }
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );

    let down = (flags & FR_DOWN) != FINDREPLACE_FLAGS(0);

    let mut ft = FINDTEXTEXW {
        chrg: CHARRANGE {
            cpMin: if down { cr.cpMax } else { cr.cpMin },
            cpMax: if down { -1 } else { 0 },
        },
        lpstrText: PCWSTR(to_wide(search).as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };

    let result = SendMessageW(
        hwnd_edit,
        EM_FINDTEXTEXW,
        WPARAM(flags.0 as usize),
        LPARAM(&mut ft as *mut _ as isize),
    );

    if result.0 != -1 {
        let mut sel = ft.chrgText;
        std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut sel as *mut _ as isize),
        );
        SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_edit);
        return true;
    }

    if wrap {
        ft.chrg.cpMin = if down { 0 } else { -1 };
        ft.chrg.cpMax = if down { -1 } else { 0 };
        let result = SendMessageW(
            hwnd_edit,
            EM_FINDTEXTEXW,
            WPARAM(flags.0 as usize),
            LPARAM(&mut ft as *mut _ as isize),
        );
        if result.0 != -1 {
            let mut sel = ft.chrgText;
            std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
            SendMessageW(
                hwnd_edit,
                EM_EXSETSEL,
                WPARAM(0),
                LPARAM(&mut sel as *mut _ as isize),
            );
            SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            SetFocus(hwnd_edit);
            return true;
        }
    }
    false
}

unsafe fn replace_selection_if_match(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
    options: &FindOptions,
) -> bool {
    if options.use_regex {
        return replace_selection_if_match_regex(hwnd, hwnd_edit, search, replace, flags, options);
    }
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );

    if cr.cpMin == cr.cpMax {
        return false;
    }

    let wide_search = to_wide(search);
    let mut ft = FINDTEXTEXW {
        chrg: cr,
        lpstrText: PCWSTR(wide_search.as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };

    let res = SendMessageW(
        hwnd_edit,
        EM_FINDTEXTEXW,
        WPARAM(flags.0 as usize),
        LPARAM(&mut ft as *mut _ as isize),
    );

    if res.0 == cr.cpMin as isize && ft.chrgText.cpMax == cr.cpMax {
        let replace_wide = to_wide(replace);
        SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(replace_wide.as_ptr() as isize),
        );
        true
    } else {
        false
    }
}

pub unsafe fn replace_all(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
    options: &FindOptions,
) {
    if options.use_regex {
        replace_all_regex(hwnd, hwnd_edit, search, replace, flags, options);
        return;
    }
    if options.replace_in_all_docs {
        replace_all_in_all_docs(hwnd, search, replace, flags);
        return;
    }
    if options.replace_in_selection {
        replace_all_in_selection(hwnd, hwnd_edit, search, replace, flags);
        return;
    }
    if search.is_empty() {
        return;
    }
    let mut start = 0i32;
    let mut count = 0usize;
    let replace_wide = to_wide(replace);

    loop {
        let mut ft = FINDTEXTEXW {
            chrg: CHARRANGE {
                cpMin: start,
                cpMax: -1,
            },
            lpstrText: PCWSTR(to_wide(search).as_ptr()),
            chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
        };

        let res = SendMessageW(
            hwnd_edit,
            EM_FINDTEXTEXW,
            WPARAM(flags.0 as usize),
            LPARAM(&mut ft as *mut _ as isize),
        );

        if res.0 != -1 {
            SendMessageW(
                hwnd_edit,
                EM_EXSETSEL,
                WPARAM(0),
                LPARAM(&mut ft.chrgText as *mut _ as isize),
            );
            SendMessageW(
                hwnd_edit,
                EM_REPLACESEL,
                WPARAM(1),
                LPARAM(replace_wide.as_ptr() as isize),
            );
            count += 1;

            let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
            SendMessageW(
                hwnd_edit,
                EM_EXGETSEL,
                WPARAM(0),
                LPARAM(&mut cr as *mut _ as isize),
            );
            start = cr.cpMax;
        } else {
            break;
        }
    }

    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if count == 0 {
        let message = to_wide(&text_not_found_message(language));
        let title = to_wide(&find_title(language));
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    } else {
        let message = i18n::tr_f(
            language,
            "find.replaced_count",
            &[("count", &count.to_string())],
        );
        show_info(hwnd, language, &message);
    }
}

unsafe fn get_find_options(hwnd: HWND) -> FindOptions {
    with_state(hwnd, |state| FindOptions {
        use_regex: state.find_use_regex,
        dot_matches_newline: state.find_dot_matches_newline,
        wrap_around: state.find_wrap_around,
        replace_in_selection: state.find_replace_in_selection,
        replace_in_all_docs: state.find_replace_in_all_docs,
    })
    .unwrap_or(FindOptions {
        use_regex: false,
        dot_matches_newline: false,
        wrap_around: true,
        replace_in_selection: false,
        replace_in_all_docs: false,
    })
}

unsafe extern "system" fn find_replace_hook_proc(
    hdlg: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> usize {
    match msg {
        WM_INITDIALOG => {
            let fr = &*(lparam.0 as *const FINDREPLACEW);
            let parent = fr.hwndOwner;
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let is_replace = fr.lCustData.0 == REPLACE_DIALOG_ID;
            let (width, height, client_bottom) = dialog_metrics(hdlg);

            let (regex_text, dot_text, wrap_text, selection_text, all_docs_text) = (
                i18n::tr(language, "find.regex"),
                i18n::tr(language, "find.dot_matches_newline"),
                i18n::tr(language, "find.wrap_around"),
                i18n::tr(language, "find.replace_in_selection"),
                i18n::tr(language, "find.replace_in_all_docs"),
            );
            let options = get_find_options(parent);

            let mut y = client_bottom + 8;
            let line_h = 18;
            let gap = 4;
            let mut added = 3;
            if is_replace {
                added += 2;
            }
            let extra = (added * (line_h + gap)) + 10;
            let _ = SetWindowPos(
                hdlg,
                HWND(0),
                0,
                0,
                width,
                height + extra,
                SWP_NOMOVE | SWP_NOZORDER,
            );

            let x = 12;
            let checkbox_width = width.saturating_sub(x + 16);
            let font = GetStockObject(DEFAULT_GUI_FONT);

            let regex = create_checkbox(hdlg, FIND_ID_REGEX, &regex_text, x, y, checkbox_width);
            set_checkbox_checked(regex, options.use_regex);
            y += line_h + gap;

            let dot = create_checkbox(
                hdlg,
                FIND_ID_DOT_MATCHES_NEWLINE,
                &dot_text,
                x,
                y,
                checkbox_width,
            );
            set_checkbox_checked(dot, options.dot_matches_newline);
            y += line_h + gap;

            let wrap = create_checkbox(hdlg, FIND_ID_WRAP_AROUND, &wrap_text, x, y, checkbox_width);
            set_checkbox_checked(wrap, options.wrap_around);
            y += line_h + gap;

            if is_replace {
                let sel = create_checkbox(
                    hdlg,
                    REPLACE_ID_IN_SELECTION,
                    &selection_text,
                    x,
                    y,
                    checkbox_width,
                );
                set_checkbox_checked(sel, options.replace_in_selection);
                y += line_h + gap;

                let all_docs = create_checkbox(
                    hdlg,
                    REPLACE_ID_IN_ALL_DOCS,
                    &all_docs_text,
                    x,
                    y,
                    checkbox_width,
                );
                set_checkbox_checked(all_docs, options.replace_in_all_docs);
            }

            if font.0 != 0 {
                for id in [
                    FIND_ID_REGEX,
                    FIND_ID_DOT_MATCHES_NEWLINE,
                    FIND_ID_WRAP_AROUND,
                    REPLACE_ID_IN_SELECTION,
                    REPLACE_ID_IN_ALL_DOCS,
                ] {
                    let hwnd_child =
                        windows::Win32::UI::WindowsAndMessaging::GetDlgItem(hdlg, id as i32);
                    if hwnd_child.0 != 0 {
                        let _ = SendMessageW(
                            hwnd_child,
                            WM_SETFONT,
                            WPARAM(font.0 as usize),
                            LPARAM(1),
                        );
                    }
                }
            }
            1
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as isize;
            let parent = GetParent(hdlg);
            if parent.0 == 0 {
                return 0;
            }
            match cmd_id {
                FIND_ID_REGEX => {
                    let checked = is_checkbox_checked(hdlg, FIND_ID_REGEX);
                    let _ = with_state(parent, |state| {
                        state.find_use_regex = checked;
                    });
                }
                FIND_ID_DOT_MATCHES_NEWLINE => {
                    let checked = is_checkbox_checked(hdlg, FIND_ID_DOT_MATCHES_NEWLINE);
                    let _ = with_state(parent, |state| {
                        state.find_dot_matches_newline = checked;
                    });
                }
                FIND_ID_WRAP_AROUND => {
                    let checked = is_checkbox_checked(hdlg, FIND_ID_WRAP_AROUND);
                    let _ = with_state(parent, |state| {
                        state.find_wrap_around = checked;
                    });
                }
                REPLACE_ID_IN_SELECTION => {
                    let checked = is_checkbox_checked(hdlg, REPLACE_ID_IN_SELECTION);
                    let _ = with_state(parent, |state| {
                        state.find_replace_in_selection = checked;
                        if checked {
                            state.find_replace_in_all_docs = false;
                        }
                    });
                    if checked {
                        set_checkbox_checked_by_id(hdlg, REPLACE_ID_IN_ALL_DOCS, false);
                    }
                }
                REPLACE_ID_IN_ALL_DOCS => {
                    let checked = is_checkbox_checked(hdlg, REPLACE_ID_IN_ALL_DOCS);
                    let _ = with_state(parent, |state| {
                        state.find_replace_in_all_docs = checked;
                        if checked {
                            state.find_replace_in_selection = false;
                        }
                    });
                    if checked {
                        set_checkbox_checked_by_id(hdlg, REPLACE_ID_IN_SELECTION, false);
                    }
                }
                _ => {}
            }
            0
        }
        _ => 0,
    }
}

fn dialog_metrics(hwnd: HWND) -> (i32, i32, i32) {
    unsafe {
        let mut rc_client = windows::Win32::Foundation::RECT::default();
        let mut rc_window = windows::Win32::Foundation::RECT::default();
        let _ = GetClientRect(hwnd, &mut rc_client);
        let _ = GetWindowRect(hwnd, &mut rc_window);
        let width = rc_window.right - rc_window.left;
        let height = rc_window.bottom - rc_window.top;
        (width, height, rc_client.bottom)
    }
}

unsafe fn create_checkbox(parent: HWND, id: isize, text: &str, x: i32, y: i32, width: i32) -> HWND {
    let wide = to_wide(text);
    CreateWindowExW(
        Default::default(),
        WC_BUTTON,
        PCWSTR(wide.as_ptr()),
        windows::Win32::UI::WindowsAndMessaging::WS_CHILD
            | windows::Win32::UI::WindowsAndMessaging::WS_VISIBLE
            | windows::Win32::UI::WindowsAndMessaging::WS_TABSTOP
            | WINDOW_STYLE(BS_AUTOCHECKBOX as u32),
        x,
        y,
        width,
        18,
        parent,
        HMENU(id),
        HINSTANCE(0),
        None,
    )
}

unsafe fn is_checkbox_checked(hwnd: HWND, id: isize) -> bool {
    let hwnd_child = windows::Win32::UI::WindowsAndMessaging::GetDlgItem(hwnd, id as i32);
    if hwnd_child.0 == 0 {
        return false;
    }
    let res = SendMessageW(hwnd_child, BM_GETCHECK, WPARAM(0), LPARAM(0));
    res.0 as u32 == BST_CHECKED.0
}

unsafe fn set_checkbox_checked(hwnd: HWND, checked: bool) {
    if hwnd.0 == 0 {
        return;
    }
    let value = if checked {
        BST_CHECKED
    } else {
        Default::default()
    };
    let _ = SendMessageW(hwnd, BM_SETCHECK, WPARAM(value.0 as usize), LPARAM(0));
}

unsafe fn set_checkbox_checked_by_id(hwnd: HWND, id: isize, checked: bool) {
    let hwnd_child = windows::Win32::UI::WindowsAndMessaging::GetDlgItem(hwnd, id as i32);
    set_checkbox_checked(hwnd_child, checked);
}

fn normalize_regex_replacement(replace: &str) -> String {
    let mut out = String::with_capacity(replace.len());
    let mut chars = replace.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            if let Some(next) = chars.peek() {
                if next.is_ascii_digit() {
                    out.push('$');
                    out.push(*next);
                    chars.next();
                    continue;
                }
            }
        }
        out.push(ch);
    }
    out
}

fn build_regex(
    search: &str,
    flags: FINDREPLACE_FLAGS,
    options: &FindOptions,
) -> Result<Regex, String> {
    let mut pattern = search.to_string();
    if (flags & FR_WHOLEWORD) != FINDREPLACE_FLAGS(0) {
        pattern = format!(r"\b(?:{})\b", pattern);
    }
    let mut prefix = String::new();
    if (flags & FR_MATCHCASE) == FINDREPLACE_FLAGS(0) {
        prefix.push_str("(?i)");
    }
    if options.dot_matches_newline {
        prefix.push_str("(?s)");
    }
    let final_pattern = format!("{prefix}{pattern}");
    Regex::new(&final_pattern).map_err(|err| err.to_string())
}

unsafe fn find_next_regex(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    flags: FINDREPLACE_FLAGS,
    wrap: bool,
    options: &FindOptions,
) -> bool {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let regex = match build_regex(search, flags, options) {
        Ok(regex) => regex,
        Err(err) => {
            let message = i18n::tr_f(language, "find.regex_error", &[("err", &err)]);
            show_error(hwnd, language, &message);
            return false;
        }
    };
    let text = get_edit_text(hwnd_edit);
    if text.is_empty() {
        return false;
    }
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );
    let down = (flags & FR_DOWN) != FINDREPLACE_FLAGS(0);
    let start_utf16 = if down { cr.cpMax } else { cr.cpMin };
    let start_byte = utf16_index_to_byte(&text, start_utf16);
    let found = if down {
        regex_find_forward(&regex, &text, start_byte).or_else(|| {
            if wrap {
                regex_find_forward(&regex, &text, 0)
            } else {
                None
            }
        })
    } else {
        regex_find_backward(&regex, &text, start_byte).or_else(|| {
            if wrap {
                regex_find_backward(&regex, &text, text.len())
            } else {
                None
            }
        })
    };
    let Some((start, end)) = found else {
        return false;
    };
    let mut sel = CHARRANGE {
        cpMin: byte_index_to_utf16(&text, start),
        cpMax: byte_index_to_utf16(&text, end),
    };
    std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut sel as *mut _ as isize),
    );
    SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
    SetFocus(hwnd_edit);
    true
}

fn regex_find_forward(regex: &Regex, text: &str, start: usize) -> Option<(usize, usize)> {
    let slice = &text[start..];
    let mut iter = regex.find_iter(slice);
    while let Some(res) = iter.next() {
        if let Ok(found) = res {
            return Some((start + found.start(), start + found.end()));
        }
    }
    None
}

fn regex_find_backward(regex: &Regex, text: &str, end: usize) -> Option<(usize, usize)> {
    let slice = &text[..end];
    let mut last = None;
    let mut iter = regex.find_iter(slice);
    while let Some(res) = iter.next() {
        if let Ok(found) = res {
            last = Some((found.start(), found.end()));
        }
    }
    last
}

unsafe fn replace_selection_if_match_regex(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
    options: &FindOptions,
) -> bool {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let regex = match build_regex(search, flags, options) {
        Ok(regex) => regex,
        Err(err) => {
            let message = i18n::tr_f(language, "find.regex_error", &[("err", &err)]);
            show_error(hwnd, language, &message);
            return false;
        }
    };
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );
    if cr.cpMin == cr.cpMax {
        return false;
    }
    let text = get_edit_text(hwnd_edit);
    let start = utf16_index_to_byte(&text, cr.cpMin);
    let end = utf16_index_to_byte(&text, cr.cpMax);
    let selected = &text[start..end];
    let mut iter = regex.find_iter(selected);
    let Some(Ok(found)) = iter.next() else {
        return false;
    };
    if found.start() != 0 || found.end() != selected.len() {
        return false;
    }
    let normalized = normalize_regex_replacement(replace);
    let replaced = regex.replace(selected, normalized.as_str()).to_string();
    let replace_wide = to_wide(&replaced);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(replace_wide.as_ptr() as isize),
    );
    true
}

unsafe fn replace_all_regex(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
    options: &FindOptions,
) {
    if search.is_empty() {
        return;
    }
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let regex = match build_regex(search, flags, options) {
        Ok(regex) => regex,
        Err(err) => {
            let message = i18n::tr_f(language, "find.regex_error", &[("err", &err)]);
            show_error(hwnd, language, &message);
            return;
        }
    };
    let normalized = normalize_regex_replacement(replace);
    let mut total_count = 0usize;

    if options.replace_in_all_docs {
        let edits = with_state(hwnd, |state| {
            state
                .docs
                .iter()
                .map(|doc| doc.hwnd_edit)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
        for hwnd_doc in edits {
            let text = get_edit_text(hwnd_doc);
            if text.is_empty() {
                continue;
            }
            let count = regex.find_iter(&text).count();
            if count == 0 {
                continue;
            }
            total_count += count;
            let replaced = regex.replace_all(&text, normalized.as_str()).to_string();
            replace_range_text(hwnd_doc, 0, -1, &replaced);
        }
    } else if options.replace_in_selection {
        let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );
        if cr.cpMin == cr.cpMax {
            return;
        }
        let text = get_edit_text(hwnd_edit);
        let start = utf16_index_to_byte(&text, cr.cpMin);
        let end = utf16_index_to_byte(&text, cr.cpMax);
        let selected = &text[start..end];
        let count = regex.find_iter(selected).count();
        if count > 0 {
            total_count = count;
            let replaced = regex.replace_all(selected, normalized.as_str()).to_string();
            replace_range_text(hwnd_edit, cr.cpMin, cr.cpMax, &replaced);
        }
    } else {
        let text = get_edit_text(hwnd_edit);
        let count = regex.find_iter(&text).count();
        if count > 0 {
            total_count = count;
            let replaced = regex.replace_all(&text, normalized.as_str()).to_string();
            replace_range_text(hwnd_edit, 0, -1, &replaced);
        }
    }

    if total_count == 0 {
        let message = to_wide(&text_not_found_message(language));
        let title = to_wide(&find_title(language));
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    } else {
        let message = i18n::tr_f(
            language,
            "find.replaced_count",
            &[("count", &total_count.to_string())],
        );
        show_info(hwnd, language, &message);
    }
}

unsafe fn replace_all_in_all_docs(
    hwnd: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
) {
    let edits = with_state(hwnd, |state| {
        state
            .docs
            .iter()
            .map(|doc| doc.hwnd_edit)
            .collect::<Vec<_>>()
    })
    .unwrap_or_default();
    let mut total_count = 0usize;
    for hwnd_doc in edits {
        total_count += replace_all_in_range(hwnd_doc, search, replace, flags, 0, -1);
    }
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if total_count == 0 {
        let message = to_wide(&text_not_found_message(language));
        let title = to_wide(&find_title(language));
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    } else {
        let message = i18n::tr_f(
            language,
            "find.replaced_count",
            &[("count", &total_count.to_string())],
        );
        show_info(hwnd, language, &message);
    }
}

unsafe fn replace_all_in_selection(
    hwnd: HWND,
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
) {
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(
        hwnd_edit,
        EM_EXGETSEL,
        WPARAM(0),
        LPARAM(&mut cr as *mut _ as isize),
    );
    if cr.cpMin == cr.cpMax {
        return;
    }
    let count = replace_all_in_range(hwnd_edit, search, replace, flags, cr.cpMin, cr.cpMax);
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    if count == 0 {
        let message = to_wide(&text_not_found_message(language));
        let title = to_wide(&find_title(language));
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    } else {
        let message = i18n::tr_f(
            language,
            "find.replaced_count",
            &[("count", &count.to_string())],
        );
        show_info(hwnd, language, &message);
    }
}

unsafe fn replace_all_in_range(
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
    start_utf16: i32,
    end_utf16: i32,
) -> usize {
    if search.is_empty() {
        return 0;
    }
    let mut start = start_utf16;
    let mut end = end_utf16;
    let mut count = 0usize;
    let replace_wide = to_wide(replace);
    let replace_len = replace.chars().map(|c| c.len_utf16() as i32).sum::<i32>();

    loop {
        let mut ft = FINDTEXTEXW {
            chrg: CHARRANGE {
                cpMin: start,
                cpMax: end,
            },
            lpstrText: PCWSTR(to_wide(search).as_ptr()),
            chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
        };
        let res = SendMessageW(
            hwnd_edit,
            EM_FINDTEXTEXW,
            WPARAM(flags.0 as usize),
            LPARAM(&mut ft as *mut _ as isize),
        );
        if res.0 == -1 {
            break;
        }
        SendMessageW(
            hwnd_edit,
            EM_EXSETSEL,
            WPARAM(0),
            LPARAM(&mut ft.chrgText as *mut _ as isize),
        );
        SendMessageW(
            hwnd_edit,
            EM_REPLACESEL,
            WPARAM(1),
            LPARAM(replace_wide.as_ptr() as isize),
        );
        count += 1;
        let match_len = ft.chrgText.cpMax - ft.chrgText.cpMin;
        let delta = replace_len - match_len;
        let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
        SendMessageW(
            hwnd_edit,
            EM_EXGETSEL,
            WPARAM(0),
            LPARAM(&mut cr as *mut _ as isize),
        );
        start = cr.cpMax;
        if end != -1 {
            end += delta;
        }
    }
    count
}

unsafe fn replace_range_text(hwnd_edit: HWND, start_utf16: i32, end_utf16: i32, text: &str) {
    let mut range = CHARRANGE {
        cpMin: start_utf16,
        cpMax: end_utf16,
    };
    SendMessageW(
        hwnd_edit,
        EM_EXSETSEL,
        WPARAM(0),
        LPARAM(&mut range as *mut _ as isize),
    );
    let wide = to_wide(text);
    SendMessageW(
        hwnd_edit,
        EM_REPLACESEL,
        WPARAM(1),
        LPARAM(wide.as_ptr() as isize),
    );
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
