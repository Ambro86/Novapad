use windows::core::{PCWSTR, PWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    SendMessageW, MessageBoxW, MB_OK, MB_ICONWARNING
};
use windows::Win32::UI::Controls::RichEdit::{
    CHARRANGE, FINDTEXTEXW, EM_FINDTEXTEXW, EM_EXGETSEL, EM_EXSETSEL
};
use windows::Win32::UI::Controls::Dialogs::{
    FindTextW, ReplaceTextW, FINDREPLACEW, FINDREPLACE_FLAGS,
    FR_DOWN, FR_DIALOGTERM, FR_FINDNEXT, FR_REPLACE, FR_REPLACEALL,
    FR_MATCHCASE, FR_WHOLEWORD
};
use windows::Win32::UI::Input::KeyboardAndMouse::{SetFocus};
use crate::settings::{Language, find_title, text_not_found_message};
use crate::accessibility::{to_wide, from_wide, EM_SCROLLCARET, EM_REPLACESEL};
use crate::{with_state, get_active_edit};

pub const FIND_DIALOG_ID: isize = 1;
pub const REPLACE_DIALOG_ID: isize = 2;

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
            Flags: FR_DOWN,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lCustData: LPARAM(FIND_DIALOG_ID),
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
            Flags: FR_DOWN,
            lpstrFindWhat: PWSTR(state.find_text.as_mut_ptr()),
            wFindWhatLen: state.find_text.len() as u16,
            lpstrReplaceWith: PWSTR(state.replace_text.as_mut_ptr()),
            wReplaceWithLen: state.replace_text.len() as u16,
            lCustData: LPARAM(REPLACE_DIALOG_ID),
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
    let _ = with_state(hwnd, |state| {
        state.last_find_flags = find_flags;
    });

    if (fr.Flags & FR_REPLACEALL) != FINDREPLACE_FLAGS(0) {
        replace_all(hwnd, hwnd_edit, &search, &from_wide(fr.lpstrReplaceWith.0), find_flags);
        return;
    }

    if (fr.Flags & FR_REPLACE) != FINDREPLACE_FLAGS(0) {
        let replace = from_wide(fr.lpstrReplaceWith.0);
        let replaced = replace_selection_if_match(hwnd_edit, &search, &replace, find_flags);
        let found = find_next(hwnd_edit, &search, find_flags, true);
        if !replaced && !found {
            let message = to_wide(text_not_found_message(language));
            let title = to_wide(find_title(language));
            MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
        }
        return;
    }

    if find_next(hwnd_edit, &search, find_flags, true) {
        return;
    }
    let message = to_wide(text_not_found_message(language));
    let title = to_wide(find_title(language));
    MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
}

pub unsafe fn find_next_from_state(hwnd: HWND) {
    let (search, flags, language): (String, FINDREPLACE_FLAGS, Language) = with_state(hwnd, |state| {
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
    if !find_next(hwnd_edit, &search, flags, true) {
        let message = to_wide(text_not_found_message(language));
        let title = to_wide(find_title(language));
        MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
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
    hwnd_edit: HWND,
    search: &str,
    flags: FINDREPLACE_FLAGS,
    wrap: bool,
) -> bool {
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
    
    let down = (flags & FR_DOWN) != FINDREPLACE_FLAGS(0);
    
    let mut ft = FINDTEXTEXW {
        chrg: CHARRANGE {
            cpMin: if down { cr.cpMax } else { cr.cpMin },
            cpMax: if down { -1 } else { 0 },
        },
        lpstrText: PCWSTR(to_wide(search).as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };

    let result = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
    
    if result.0 != -1 {
        let mut sel = ft.chrgText;
        std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
        SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut sel as *mut _ as isize));
        SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_edit);
        return true;
    }

    if wrap {
        ft.chrg.cpMin = if down { 0 } else { -1 };
        ft.chrg.cpMax = if down { -1 } else { 0 };
        let result = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
        if result.0 != -1 {
            let mut sel = ft.chrgText;
            std::mem::swap(&mut sel.cpMin, &mut sel.cpMax);
            SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut sel as *mut _ as isize));
            SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
            SetFocus(hwnd_edit);
            return true;
        }
    }
    false
}

unsafe fn replace_selection_if_match(
    hwnd_edit: HWND,
    search: &str,
    replace: &str,
    flags: FINDREPLACE_FLAGS,
) -> bool {
    let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
    SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
    
    if cr.cpMin == cr.cpMax {
        return false;
    }

    let wide_search = to_wide(search);
    let mut ft = FINDTEXTEXW {
        chrg: cr,
        lpstrText: PCWSTR(wide_search.as_ptr()),
        chrgText: CHARRANGE { cpMin: 0, cpMax: 0 },
    };
    
    let res = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
    
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
) {
    if search.is_empty() {
        return;
    }
    let mut start = 0i32;
    let mut replaced_any = false;
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

        let res = SendMessageW(hwnd_edit, EM_FINDTEXTEXW, WPARAM(flags.0 as usize), LPARAM(&mut ft as *mut _ as isize));
        
        if res.0 != -1 {
            SendMessageW(hwnd_edit, EM_EXSETSEL, WPARAM(0), LPARAM(&mut ft.chrgText as *mut _ as isize));
            SendMessageW(
                hwnd_edit,
                EM_REPLACESEL,
                WPARAM(1),
                LPARAM(replace_wide.as_ptr() as isize),
            );
            replaced_any = true;
            
            let mut cr = CHARRANGE { cpMin: 0, cpMax: 0 };
            SendMessageW(hwnd_edit, EM_EXGETSEL, WPARAM(0), LPARAM(&mut cr as *mut _ as isize));
            start = cr.cpMax;
        } else {
            break;
        }
    }
    
    if !replaced_any {
        let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
        let message = to_wide(text_not_found_message(language));
        let title = to_wide(find_title(language));
        MessageBoxW(hwnd, PCWSTR(message.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONWARNING);
    }
}
