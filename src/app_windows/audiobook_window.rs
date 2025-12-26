use windows::core::{PCWSTR, w};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM, LRESULT, HINSTANCE, RECT};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, GetWindowLongPtrW, RegisterClassW,
    SendMessageW, SetWindowLongPtrW, SetForegroundWindow, GetParent, MessageBoxW,
    GWLP_USERDATA, WM_CREATE, WM_DESTROY, WM_NCDESTROY, WM_CLOSE, WM_COMMAND, WM_SETFOCUS,
    WM_KEYDOWN, WM_APP, MoveWindow, SetWindowTextW, MSG,
    WS_POPUP, WS_CAPTION, WS_VISIBLE, WS_CHILD, WS_TABSTOP, WS_EX_DLGMODALFRAME,
    CW_USEDEFAULT, HMENU, WNDCLASSW,
    BS_DEFPUSHBUTTON, IDYES, MB_YESNO, MB_ICONWARNING,
    CREATESTRUCTW, LoadCursorW, IDC_ARROW, WINDOW_STYLE
};
use windows::Win32::UI::Controls::{
    WC_BUTTON, PBM_SETRANGE, PBM_SETPOS
};
use windows::Win32::Graphics::Gdi::{HBRUSH, COLOR_WINDOW};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus, EnableWindow, VK_RETURN};
use crate::{with_state};
use crate::settings::{Language};
use crate::accessibility::{to_wide, handle_accessibility, ES_CENTER, ES_READONLY};
use std::sync::atomic::{Ordering};

const PROGRESS_CLASS_NAME: &str = "NovapadProgress";
const PROGRESS_ID_CANCEL: usize = 8001;
const WM_UPDATE_PROGRESS: u32 = WM_APP + 6;

struct ProgressDialogState {
    hwnd_pb: HWND,
    hwnd_text: HWND,
    hwnd_cancel: HWND,
    total: usize,
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        let cancel_btn = with_progress_state(hwnd, |s| s.hwnd_cancel).unwrap_or(HWND(0));
        if focus == cancel_btn {
            request_cancel(hwnd);
            return true;
        }
    }
    handle_accessibility(hwnd, msg)
}

pub unsafe fn open(parent: HWND, total: usize) -> HWND {
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PROGRESS_CLASS_NAME);
    
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(LoadCursorW(None, IDC_ARROW).unwrap_or_default().0),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(progress_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc); 

    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        w!("Creazione Audiolibro"),
        WS_POPUP | WS_CAPTION | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT,
        300, 150,
        parent,
        HMENU(0),
        hinstance,
        Some(parent.0 as *const _),
    );
    
    if hwnd.0 != 0 {
         EnableWindow(parent, false);
         let _ = with_progress_state(hwnd, |state| {
             let _ = SendMessageW(state.hwnd_pb, PBM_SETRANGE, WPARAM(0), LPARAM((total as isize) << 16));
             state.total = total;
         });
         
         // Center window relative to parent
         let mut rc_parent = RECT::default();
         let mut rc_dlg = RECT::default();
         let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(parent, &mut rc_parent);
         let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowRect(hwnd, &mut rc_dlg);
         
         let dlg_w = rc_dlg.right - rc_dlg.left;
         let dlg_h = rc_dlg.bottom - rc_dlg.top;
         let parent_w = rc_parent.right - rc_parent.left;
         let parent_h = rc_parent.bottom - rc_parent.top;
         
         let x = rc_parent.left + (parent_w - dlg_w) / 2;
         let y = rc_parent.top + (parent_h - dlg_h) / 2;
         
         let _ = MoveWindow(hwnd, x, y, dlg_w, dlg_h, true);
    }
    hwnd
}

unsafe extern "system" fn progress_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
             let create_struct = lparam.0 as *const CREATESTRUCTW;
             let _parent = HWND((*create_struct).lpCreateParams as isize);
             
             let label = CreateWindowExW(
                 Default::default(),
                 w!("EDIT"),
                 w!("Creazione audiolibro in corso. Avanzamento: 0%"),
                 WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE((ES_CENTER | ES_READONLY) as u32),
                 20, 20, 240, 20,
                 hwnd, HMENU(0), HINSTANCE(0), None
             );
             
             let pb = CreateWindowExW(
                 Default::default(),
                 w!("msctls_progress32"),
                 PCWSTR::null(),
                 WS_CHILD | WS_VISIBLE,
                 20, 50, 240, 20,
                 hwnd, HMENU(0), HINSTANCE(0), None
             );

             let hwnd_cancel = CreateWindowExW(
                 Default::default(),
                 WC_BUTTON,
                 w!("Annulla"),
                 WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                 95, 80, 90, 28,
                 hwnd, HMENU(PROGRESS_ID_CANCEL as isize), HINSTANCE(0), None
             );
             
             let state = Box::new(ProgressDialogState { hwnd_pb: pb, hwnd_text: label, hwnd_cancel, total: 0 }); 
             SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
             
             if label.0 != 0 {
                 SetFocus(label);
             }
             LRESULT(0)
        }
        WM_SETFOCUS => {
            let _ = with_progress_state(hwnd, |state| {
                SetFocus(state.hwnd_text);
            });
            LRESULT(0)
        }
        WM_COMMAND => {
             let cmd_id = (wparam.0 & 0xffff) as usize;
             if cmd_id == PROGRESS_ID_CANCEL || cmd_id == 2 { // 2 is IDCANCEL
                 request_cancel(hwnd);
                 LRESULT(0)
             } else {
                 DefWindowProcW(hwnd, msg, wparam, lparam)
             }
        }
        WM_UPDATE_PROGRESS => {
             let current = wparam.0;
             let _ = with_progress_state(hwnd, |state| {
                 let _ = SendMessageW(state.hwnd_pb, PBM_SETPOS, WPARAM(current), LPARAM(0));
                 if state.total > 0 {
                     let pct = (current * 100) / state.total;
                     let text = format!("Creazione audiolibro in corso. Avanzamento: {}%", pct);
                     let wide = to_wide(&text);
                     let _ = SetWindowTextW(state.hwnd_text, PCWSTR(wide.as_ptr()));
                 }
             });
             LRESULT(0)
        }
        WM_CLOSE => {
            request_cancel(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

pub unsafe fn request_cancel(hwnd: HWND) {
    let parent = GetParent(hwnd);
    if parent.0 == 0 { return; }
    
    let already_cancelled = with_state(parent, |state| {
        state.audiobook_cancel.as_ref().map(|c| c.load(Ordering::Relaxed)).unwrap_or(false)
    }).unwrap_or(false);

    if already_cancelled { return; }
    
    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let (msg, title) = match language {
        Language::Italian => ("Sei sicuro di voler annullare la creazione dell'audiolibro?", "Conferma"),
        Language::English => ("Are you sure you want to cancel the audiobook creation?", "Confirm"),
    };
    
    let msg_w = to_wide(msg);
    let title_w = to_wide(title);
    
    if MessageBoxW(hwnd, PCWSTR(msg_w.as_ptr()), PCWSTR(title_w.as_ptr()), MB_YESNO | MB_ICONWARNING) == IDYES {
        let _ = with_state(parent, |state| {
            if let Some(cancel) = &state.audiobook_cancel {
                cancel.store(true, Ordering::Relaxed);
            }
        });
    }
}

unsafe fn with_progress_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut ProgressDialogState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}