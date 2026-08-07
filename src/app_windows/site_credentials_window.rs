use crate::accessibility::to_wide;
use crate::i18n;
use crate::settings::{
    Language, clear_all_ytdlp_site_credentials, clear_ytdlp_site_credentials, confirm_title,
    list_ytdlp_site_credentials, save_settings,
};
use crate::with_state;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_LISTBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, SetFocus};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CREATESTRUCTW, CW_USEDEFAULT, CreateWindowExW, DefWindowProcW,
    DispatchMessageW, GWLP_USERDATA, GetMessageW, HMENU, IDC_ARROW, IDCANCEL, IDYES,
    IsDialogMessageW, IsWindow, LB_ADDSTRING, LB_GETCURSEL, LB_GETTEXT, LB_GETTEXTLEN,
    LB_RESETCONTENT, LB_SETCURSEL, LBN_DBLCLK, LBN_SELCHANGE, LBS_HASSTRINGS, LBS_NOTIFY,
    LoadCursorW, MB_ICONQUESTION, MB_YESNO, MSG, MessageBoxW, RegisterClassW, SW_HIDE, SW_SHOW,
    SetForegroundWindow, ShowWindow, TranslateMessage, WINDOW_STYLE, WM_CLOSE, WM_COMMAND,
    WM_CREATE, WM_DESTROY, WM_NCDESTROY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::PCWSTR;

const SITE_CREDENTIALS_CLASS_NAME: &str = "SonarpadSiteCredentials";
const SITE_CREDENTIALS_ID_LIST: usize = 9101;
const SITE_CREDENTIALS_ID_REMOVE: usize = 9102;
const SITE_CREDENTIALS_ID_REMOVE_ALL: usize = 9103;
const SITE_CREDENTIALS_ID_CLOSE: usize = 9104;

struct SiteCredentialsCreateParams {
    options_parent: HWND,
    app_parent: HWND,
}

struct SiteCredentialsWindowState {
    options_parent: HWND,
    app_parent: HWND,
    language: Language,
    hwnd_list: HWND,
    hwnd_empty: HWND,
    hwnd_remove: HWND,
    hwnd_remove_all: HWND,
}

pub fn open(options_parent: HWND, app_parent: HWND) {
    unsafe {
        let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
        let class_name = to_wide(SITE_CREDENTIALS_CLASS_NAME);

        static ONCE: std::sync::Once = std::sync::Once::new();
        ONCE.call_once(|| {
            let wc = WNDCLASSW {
                hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
                    LoadCursorW(None, IDC_ARROW).unwrap_or_default().0,
                ),
                hInstance: hinstance,
                lpszClassName: PCWSTR(class_name.as_ptr()),
                lpfnWndProc: Some(site_credentials_wndproc),
                hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
                ..Default::default()
            };
            RegisterClassW(&wc);
        });

        let language = with_state(app_parent, |state| state.settings.language).unwrap_or_default();
        let title = to_wide(&i18n::tr(language, "site_credentials.title"));
        let mut create_params = SiteCredentialsCreateParams {
            options_parent,
            app_parent,
        };

        let hwnd = CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            430,
            360,
            options_parent,
            HMENU(0),
            hinstance,
            Some(&mut create_params as *mut _ as *const std::ffi::c_void),
        );

        if hwnd.0 == 0 {
            return;
        }

        EnableWindow(options_parent, false);
        SetForegroundWindow(hwnd);

        let mut msg = MSG::default();
        while IsWindow(hwnd).as_bool() && GetMessageW(&mut msg, HWND(0), 0, 0).into() {
            if crate::app_windows::calendar_window::handle_reminder_alert_message(&msg) {
                continue;
            }
            if !IsDialogMessageW(hwnd, &msg).as_bool() {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }

        EnableWindow(options_parent, true);
        SetForegroundWindow(options_parent);
    }
}

unsafe extern "system" fn site_credentials_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "site_credentials_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || site_credentials_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn site_credentials_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let params_ptr =
                unsafe { (*create_struct).lpCreateParams as *mut SiteCredentialsCreateParams };
            let Some((options_parent, app_parent)) =
                crate::with_raw_mut_ptr_safe(params_ptr, |params| {
                    (params.options_parent, params.app_parent)
                })
            else {
                crate::log_debug("Site credentials dialog create params pointer unavailable");
                return LRESULT(0);
            };

            let language =
                with_state(app_parent, |state| state.settings.language).unwrap_or_default();
            let hfont = with_state(app_parent, |state| state.hfont).unwrap_or_else(|| {
                HFONT(
                    crate::get_stock_object_safe(windows::Win32::Graphics::Gdi::DEFAULT_GUI_FONT).0,
                )
            });

            let sites_label_text = i18n::tr(language, "site_credentials.sites_label");
            let empty_text = i18n::tr(language, "site_credentials.empty");
            let remove_text = i18n::tr(language, "site_credentials.remove");
            let remove_all_text = i18n::tr(language, "site_credentials.remove_all");
            let close_text = i18n::tr(language, "site_credentials.close");

            let label_sites = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&sites_label_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    16,
                    380,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };
            let hwnd_list = unsafe {
                CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    WC_LISTBOXW,
                    PCWSTR::null(),
                    WS_CHILD
                        | WS_VISIBLE
                        | WS_VSCROLL
                        | WS_TABSTOP
                        | WINDOW_STYLE((LBS_NOTIFY | LBS_HASSTRINGS) as u32),
                    16,
                    40,
                    380,
                    190,
                    hwnd,
                    HMENU(SITE_CREDENTIALS_ID_LIST as isize),
                    HINSTANCE(0),
                    None,
                )
            };
            let hwnd_empty = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_STATIC,
                    PCWSTR(to_wide(&empty_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    16,
                    238,
                    380,
                    20,
                    hwnd,
                    HMENU(0),
                    HINSTANCE(0),
                    None,
                )
            };
            let hwnd_remove = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&remove_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    16,
                    272,
                    110,
                    28,
                    hwnd,
                    HMENU(SITE_CREDENTIALS_ID_REMOVE as isize),
                    HINSTANCE(0),
                    None,
                )
            };
            let hwnd_remove_all = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&remove_all_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    142,
                    272,
                    110,
                    28,
                    hwnd,
                    HMENU(SITE_CREDENTIALS_ID_REMOVE_ALL as isize),
                    HINSTANCE(0),
                    None,
                )
            };
            let hwnd_close = unsafe {
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&close_text).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    286,
                    272,
                    110,
                    28,
                    hwnd,
                    HMENU(SITE_CREDENTIALS_ID_CLOSE as isize),
                    HINSTANCE(0),
                    None,
                )
            };

            for control in [
                label_sites,
                hwnd_list,
                hwnd_empty,
                hwnd_remove,
                hwnd_remove_all,
                hwnd_close,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    crate::send_message_w_safe(
                        control,
                        WM_SETFONT,
                        WPARAM(hfont.0 as usize),
                        LPARAM(1),
                    );
                }
            }

            let state = Box::new(SiteCredentialsWindowState {
                options_parent,
                app_parent,
                language,
                hwnd_list,
                hwnd_empty,
                hwnd_remove,
                hwnd_remove_all,
            });
            crate::set_window_long_ptr_w_safe(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            refresh_site_credentials_list(hwnd);

            unsafe {
                SetFocus(hwnd_list);
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = wparam.0 & 0xffff;
            let notify = (wparam.0 >> 16) as u16;
            match cmd_id {
                SITE_CREDENTIALS_ID_REMOVE => {
                    remove_selected_site_credentials(hwnd);
                    LRESULT(0)
                }
                SITE_CREDENTIALS_ID_REMOVE_ALL => {
                    remove_all_site_credentials(hwnd);
                    LRESULT(0)
                }
                SITE_CREDENTIALS_ID_CLOSE => {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    LRESULT(0)
                }
                SITE_CREDENTIALS_ID_LIST if notify == LBN_SELCHANGE as u16 => {
                    update_site_credentials_buttons(hwnd);
                    LRESULT(0)
                }
                SITE_CREDENTIALS_ID_LIST if notify == LBN_DBLCLK as u16 => {
                    remove_selected_site_credentials(hwnd);
                    LRESULT(0)
                }
                cmd if cmd == IDCANCEL.0 as usize || cmd == 2 => {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    LRESULT(0)
                }
                _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            crate::log_if_err!(crate::destroy_window_safe(hwnd));
            LRESULT(0)
        }
        WM_DESTROY => {
            if let Some(options_parent) =
                with_site_credentials_state(hwnd, |state| state.options_parent)
            {
                crate::enable_window_safe(options_parent, true);
                crate::set_foreground_window_safe(options_parent);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA)
                as *mut SiteCredentialsWindowState;
            if !ptr.is_null() {
                let _unused_box = crate::box_from_raw_safe(ptr);
            }
            LRESULT(0)
        }
        _ => crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
    }
}

fn with_site_credentials_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut SiteCredentialsWindowState) -> R,
{
    let ptr =
        crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut SiteCredentialsWindowState;
    crate::with_raw_mut_ptr_safe(ptr, f)
}

fn selected_site(hwnd_list: HWND) -> Option<String> {
    let sel = crate::send_message_w_safe(hwnd_list, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel < 0 {
        return None;
    }
    let len =
        crate::send_message_w_safe(hwnd_list, LB_GETTEXTLEN, WPARAM(sel as usize), LPARAM(0)).0;
    if len < 0 {
        return None;
    }
    let mut buf = vec![0u16; (len + 1) as usize];
    crate::send_message_w_safe(
        hwnd_list,
        LB_GETTEXT,
        WPARAM(sel as usize),
        LPARAM(buf.as_mut_ptr() as isize),
    );
    Some(String::from_utf16_lossy(&buf[..len as usize]))
}

fn refresh_site_credentials_list(hwnd: HWND) {
    let Some((app_parent, hwnd_list, hwnd_empty, hwnd_remove_all)) =
        with_site_credentials_state(hwnd, |state| {
            (
                state.app_parent,
                state.hwnd_list,
                state.hwnd_empty,
                state.hwnd_remove_all,
            )
        })
    else {
        crate::log_debug("Failed to access site credentials window state");
        return;
    };

    let sites = with_state(app_parent, |state| {
        list_ytdlp_site_credentials(&state.settings)
    })
    .unwrap_or_default();
    crate::send_message_w_safe(hwnd_list, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
    for site in &sites {
        let wide = to_wide(site);
        unsafe {
            windows::Win32::UI::WindowsAndMessaging::SendMessageW(
                hwnd_list,
                LB_ADDSTRING,
                WPARAM(0),
                LPARAM(wide.as_ptr() as isize),
            );
        }
    }

    if !sites.is_empty() {
        crate::send_message_w_safe(hwnd_list, LB_SETCURSEL, WPARAM(0), LPARAM(0));
        unsafe {
            ShowWindow(hwnd_empty, SW_HIDE);
        }
        crate::enable_window_safe(hwnd_remove_all, true);
    } else {
        unsafe {
            ShowWindow(hwnd_empty, SW_SHOW);
        }
        crate::enable_window_safe(hwnd_remove_all, false);
    }
    update_site_credentials_buttons(hwnd);
}

fn update_site_credentials_buttons(hwnd: HWND) {
    let Some((hwnd_list, hwnd_remove)) =
        with_site_credentials_state(hwnd, |state| (state.hwnd_list, state.hwnd_remove))
    else {
        crate::log_debug("Failed to access site credentials window state");
        return;
    };
    crate::enable_window_safe(hwnd_remove, selected_site(hwnd_list).is_some());
}

fn remove_selected_site_credentials(hwnd: HWND) {
    let Some((app_parent, language, hwnd_list)) = with_site_credentials_state(hwnd, |state| {
        (state.app_parent, state.language, state.hwnd_list)
    }) else {
        crate::log_debug("Failed to access site credentials window state");
        return;
    };
    let Some(site) = selected_site(hwnd_list) else {
        return;
    };

    let title = to_wide(&confirm_title(language));
    let message = to_wide(&i18n::tr_f(
        language,
        "site_credentials.remove_confirm",
        &[("site", site.as_str())],
    ));
    let result = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if result != IDYES {
        return;
    }

    let settings_snapshot = with_state(app_parent, |state| {
        if clear_ytdlp_site_credentials(&mut state.settings, &site) {
            Some(state.settings.clone())
        } else {
            None
        }
    })
    .flatten();
    if let Some(settings_snapshot) = settings_snapshot {
        save_settings(settings_snapshot);
        refresh_site_credentials_list(hwnd);
    }
}

fn remove_all_site_credentials(hwnd: HWND) {
    let Some((app_parent, language)) =
        with_site_credentials_state(hwnd, |state| (state.app_parent, state.language))
    else {
        crate::log_debug("Failed to access site credentials window state");
        return;
    };

    let title = to_wide(&confirm_title(language));
    let message = to_wide(&i18n::tr(language, "site_credentials.remove_all_confirm"));
    let result = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(message.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    if result != IDYES {
        return;
    }

    let settings_snapshot = with_state(app_parent, |state| {
        if clear_all_ytdlp_site_credentials(&mut state.settings) {
            Some(state.settings.clone())
        } else {
            None
        }
    })
    .flatten();
    if let Some(settings_snapshot) = settings_snapshot {
        save_settings(settings_snapshot);
        refresh_site_credentials_list(hwnd);
    }
}
