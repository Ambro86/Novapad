use crate::accessibility::{from_wide, to_wide};

use crate::i18n;
use crate::tools::rss::{self, RssItem, RssSource, RssSourceType};
use crate::with_state;
use std::collections::{HashMap, HashSet};
use std::mem;
use std::path::PathBuf;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::DataExchange::COPYDATASTRUCT;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::{
    NM_RCLICK, NMHDR, NMTREEVIEWW, NMTVKEYDOWN, TVE_EXPAND, TVGN_CARET, TVGN_CHILD, TVHITTESTINFO,
    TVI_LAST, TVI_ROOT, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0,
    TVITEMEXW_CHILDREN, TVITEMW, TVM_DELETEITEM, TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_GETNEXTITEM,
    TVM_HITTEST, TVM_INSERTITEMW, TVM_SELECTITEM, TVM_SETITEMW, TVN_ITEMEXPANDINGW, TVN_KEYDOWN,
    TVN_SELCHANGEDW, WC_BUTTON,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetActiveWindow, SetFocus, VK_APPS, VK_ESCAPE, VK_F10, VK_SHIFT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BS_DEFPUSHBUTTON, CHILDID_SELF, CREATESTRUCTW, CW_USEDEFAULT, CallWindowProcW,
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DestroyWindow,
    EVENT_OBJECT_FOCUS, GWLP_USERDATA, GWLP_WNDPROC, GetCursorPos, GetDlgItem, GetParent,
    GetWindowLongPtrW, GetWindowRect, HMENU, IDYES, KillTimer, MB_ICONQUESTION, MB_YESNO,
    MF_STRING, MessageBoxW, OBJID_CLIENT, PostMessageW, RegisterClassW, SendMessageW,
    SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW, TrackPopupMenu, WINDOW_STYLE, WM_CLOSE,
    WM_COMMAND, WM_CONTEXTMENU, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL,
    WM_NOTIFY, WM_NULL, WM_SETFOCUS, WM_SETFONT, WM_SETREDRAW, WM_SYSKEYDOWN, WM_TIMER, WM_USER,
    WNDCLASSW, WNDPROC, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_DLGMODALFRAME, WS_POPUP,
    WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const RSS_WINDOW_CLASS: &str = "NovapadRssWindow";
const ID_TREE: usize = 1001;
const ID_BTN_ADD: usize = 1002;
const ID_BTN_CLOSE: usize = 1003;
const ID_CTX_EDIT: usize = 1101;
const ID_CTX_DELETE: usize = 1102;

const WM_RSS_FETCH_COMPLETE: u32 = WM_USER + 200;
const WM_RSS_IMPORT_COMPLETE: u32 = WM_USER + 201;
const WM_SHOW_ADD_DIALOG: u32 = WM_USER + 202;
const WM_CLEAR_ENTER_GUARD: u32 = WM_USER + 203;
const WM_CLEAR_ADD_GUARD: u32 = WM_USER + 204;
pub(crate) const WM_RSS_SHOW_CONTEXT: u32 = WM_USER + 205;
const ADD_GUARD_TIMER_ID: usize = 1;
const EM_REPLACESEL: u32 = 0x00C2;

const FEED_EN_DATA: &str = include_str!("../../i18n/feed_en.txt");
const FEED_IT_DATA: &str = include_str!("../../i18n/feed_it.txt");
const FEED_ES_DATA: &str = include_str!("../../i18n/feed_es.txt");
const FEED_PT_DATA: &str = include_str!("../../i18n/feed_pt.txt");
const EM_SETSEL: u32 = 0x00B1;
const EM_SCROLLCARET: u32 = 0x00B7;
const EM_LIMITTEXT: u32 = 0x00C5;
const INITIAL_LOAD_COUNT: usize = 5;
const LOAD_MORE_COUNT: usize = 5;

// Normalize article text before sending it to the editor:
// - collapse multiple blank lines to a single blank line
// - replace embedded NULs (which would truncate Win32 edit text)
fn normalize_article_text(s: &str) -> String {
    let no_nul: String = s.chars().map(|c| if c == '\0' { ' ' } else { c }).collect();
    collapse_blank_lines(&no_nul)
}

fn normalize_rss_url_key(url: &str) -> String {
    let mut s = url.trim().to_string();
    if s.is_empty() {
        return s;
    }
    if let Some(rest) = s.strip_prefix("https://") {
        s = rest.to_string();
    } else if let Some(rest) = s.strip_prefix("http://") {
        s = rest.to_string();
    }
    if let Some((left, _)) = s.split_once('#') {
        s = left.to_string();
    }
    if let Some((left, _)) = s.split_once('?') {
        s = left.to_string();
    }
    while s.ends_with('/') && s.len() > 1 {
        s.pop();
    }
    s.to_ascii_lowercase()
}

unsafe fn rss_page_sizes(parent: HWND) -> (usize, usize) {
    with_state(parent, |s| {
        (
            s.settings.rss_initial_page_size,
            s.settings.rss_next_page_size,
        )
    })
    .unwrap_or((INITIAL_LOAD_COUNT, LOAD_MORE_COUNT))
}

unsafe fn ensure_rss_http(parent: HWND) {
    let config = with_state(parent, |s| rss::config_from_settings(&s.settings))
        .unwrap_or_else(|| rss::RssHttpConfig::default());
    if let Err(err) = rss::init_http(config) {
        eprintln!("rss_http_init_error: {}", err);
    }
}

unsafe fn rss_fetch_config(parent: HWND) -> rss::RssFetchConfig {
    with_state(parent, |s| rss::fetch_config_from_settings(&s.settings))
        .unwrap_or_else(|| rss::RssFetchConfig::default())
}

fn default_feed_path(language: crate::settings::Language) -> Option<PathBuf> {
    let file_name = match language {
        crate::settings::Language::English => "feed_en.txt",
        crate::settings::Language::Italian => "feed_it.txt",
        crate::settings::Language::Spanish => "feed_es.txt",
        crate::settings::Language::Portuguese => "feed_pt.txt",
    };
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()));
    let mut candidates = Vec::new();
    if let Some(dir) = exe_dir {
        candidates.push(dir.join("i18n").join(file_name));
    }
    if let Ok(dir) = std::env::current_dir() {
        candidates.push(dir.join("i18n").join(file_name));
    }
    for path in candidates {
        if path.exists() {
            return Some(path);
        }
    }
    None
}

fn embedded_default_feeds(language: crate::settings::Language) -> &'static str {
    match language {
        crate::settings::Language::English => FEED_EN_DATA,
        crate::settings::Language::Italian => FEED_IT_DATA,
        crate::settings::Language::Spanish => FEED_ES_DATA,
        crate::settings::Language::Portuguese => FEED_PT_DATA,
    }
}

fn load_default_feeds(language: crate::settings::Language) -> Vec<(String, String)> {
    let data = default_feed_path(language).and_then(|path| std::fs::read_to_string(path).ok());
    let data = data
        .as_deref()
        .unwrap_or_else(|| embedded_default_feeds(language));
    data.lines()
        .map(|line| line.trim())
        .filter(|line| !line.is_empty())
        .map(|line| {
            if let Some((left, right)) = line.split_once('|') {
                let title = left.trim();
                let url = right.trim();
                let title = if title.is_empty() { url } else { title };
                (title.to_string(), url.to_string())
            } else {
                (line.to_string(), line.to_string())
            }
        })
        .filter(|(_, url)| !url.is_empty())
        .collect()
}

fn is_default_key(
    language: crate::settings::Language,
    settings: &crate::settings::AppSettings,
    key: &str,
) -> bool {
    match language {
        crate::settings::Language::English => settings
            .rss_default_en_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Italian => settings
            .rss_default_it_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Spanish => settings
            .rss_default_es_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Portuguese => settings
            .rss_default_pt_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
    }
}

fn apply_default_sources(
    rss_sources: &mut Vec<RssSource>,
    removed_list: &mut Vec<String>,
    keys_list: &mut Vec<String>,
    defaults: &[(String, String)],
) -> bool {
    let mut default_items: Vec<(String, String, String)> = Vec::new();
    let mut default_by_key: HashMap<String, (String, String)> = HashMap::new();
    for (title, url) in defaults {
        let key = normalize_rss_url_key(url);
        if key.is_empty() {
            continue;
        }
        if default_by_key.contains_key(&key) {
            continue;
        }
        default_by_key.insert(key.clone(), (title.clone(), url.clone()));
        default_items.push((key, title.clone(), url.clone()));
    }
    if default_items.is_empty() {
        return false;
    }
    let current_default_keys: HashSet<String> =
        default_items.iter().map(|(k, _, _)| k.clone()).collect();

    let mut removed = HashSet::new();
    for url in removed_list.iter() {
        let key = normalize_rss_url_key(url);
        if !key.is_empty() {
            removed.insert(key);
        }
    }
    let mut existing = HashSet::new();
    for src in rss_sources.iter() {
        let key = normalize_rss_url_key(&src.url);
        if !key.is_empty() {
            existing.insert(key);
        }
    }
    let mut changed = false;
    let stored_keys: HashSet<String> = keys_list
        .iter()
        .map(|k| k.trim().to_string())
        .filter(|k| !k.is_empty())
        .collect();
    if !stored_keys.is_empty() {
        let before_len = rss_sources.len();
        rss_sources.retain(|src| {
            let key = normalize_rss_url_key(&src.url);
            if key.is_empty() {
                return true;
            }
            if stored_keys.contains(&key) && !current_default_keys.contains(&key) {
                return false;
            }
            true
        });
        if rss_sources.len() != before_len {
            changed = true;
        }
    }

    let mut seen_keys = HashSet::new();
    let before_len = rss_sources.len();
    rss_sources.retain(|src| {
        let key = normalize_rss_url_key(&src.url);
        if key.is_empty() {
            return true;
        }
        if current_default_keys.contains(&key) || stored_keys.contains(&key) {
            if !seen_keys.insert(key) {
                return false;
            }
        }
        true
    });
    if rss_sources.len() != before_len {
        changed = true;
    }

    for src in rss_sources.iter_mut() {
        let key = normalize_rss_url_key(&src.url);
        let Some((title, _url)) = default_by_key.get(&key) else {
            continue;
        };
        if removed.contains(&key) {
            continue;
        }
        if !title.trim().is_empty() && src.title != *title {
            src.title = title.clone();
            changed = true;
        }
        if src.user_title != (title.trim() != src.url.trim()) {
            src.user_title = title.trim() != src.url.trim();
            changed = true;
        }
        if !matches!(src.kind, RssSourceType::Feed) {
            src.kind = RssSourceType::Feed;
            changed = true;
        }
    }

    for (key, title, url) in &default_items {
        if removed.contains(key) || existing.contains(key) {
            continue;
        }
        rss_sources.push(RssSource {
            title: title.clone(),
            url: url.clone(),
            kind: RssSourceType::Feed,
            user_title: title.trim() != url.trim(),
            cache: rss::RssFeedCache::default(),
        });
        existing.insert(key.clone());
        changed = true;
    }

    let mut new_keys: Vec<String> = current_default_keys.into_iter().collect();
    new_keys.sort();
    let mut old_keys: Vec<String> = keys_list.clone();
    old_keys.sort();
    if new_keys != old_keys {
        *keys_list = new_keys;
        changed = true;
    }
    changed
}

unsafe fn ensure_default_sources(parent: HWND) {
    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    let defaults = load_default_feeds(language);
    if defaults.is_empty() {
        return;
    }
    with_state(parent, |s| {
        let changed = match language {
            crate::settings::Language::English => apply_default_sources(
                &mut s.settings.rss_sources,
                &mut s.settings.rss_removed_default_en,
                &mut s.settings.rss_default_en_keys,
                &defaults,
            ),
            crate::settings::Language::Italian => apply_default_sources(
                &mut s.settings.rss_sources,
                &mut s.settings.rss_removed_default_it,
                &mut s.settings.rss_default_it_keys,
                &defaults,
            ),
            crate::settings::Language::Spanish => apply_default_sources(
                &mut s.settings.rss_sources,
                &mut s.settings.rss_removed_default_es,
                &mut s.settings.rss_default_es_keys,
                &defaults,
            ),
            crate::settings::Language::Portuguese => apply_default_sources(
                &mut s.settings.rss_sources,
                &mut s.settings.rss_removed_default_pt,
                &mut s.settings.rss_default_pt_keys,
                &defaults,
            ),
        };
        if changed {
            crate::settings::save_settings(s.settings.clone());
        }
    });
}

struct RssWindowState {
    parent: HWND,
    hwnd_tree: HWND,
    node_data: HashMap<isize, NodeData>,
    pending_fetches: HashMap<String, isize>, // URL -> hItem
    source_items: HashMap<isize, SourceItemsState>,
    enter_guard: bool,
    add_guard: bool,
    pending_edit: Option<usize>,
    tree_proc: WNDPROC,
    last_selected: isize,
}

enum NodeData {
    Source(usize), // Index in settings
    Item(RssItem),
}

struct SourceItemsState {
    items: Vec<RssItem>,
    loaded: usize,
}

struct AddDialogInit {
    parent: HWND,
    prefill_title: String,
    prefill_url: String,
}

pub unsafe fn open(parent: HWND) {
    let exists = with_state(parent, |s| s.rss_window).unwrap_or(HWND(0));
    if exists.0 != 0 {
        SetForegroundWindow(exists);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(RSS_WINDOW_CLASS);

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            windows::Win32::UI::WindowsAndMessaging::LoadCursorW(
                None,
                windows::Win32::UI::WindowsAndMessaging::IDC_ARROW,
            )
            .unwrap_or_default()
            .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(rss_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    let title = to_wide(&i18n::tr(language, "rss.window.title"));

    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        500,
        600,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const _),
    );

    if hwnd.0 != 0 {
        let _ = with_state(parent, |s| s.rss_window = hwnd);
        // IMPORTANT: do NOT disable the parent window.
        // If the parent is disabled, Windows (and NVDA) treat the editor as unavailable,
        // and SetFocus/SetForegroundWindow will not behave reliably.
    }
}

pub unsafe fn show_context_menu_from_keyboard(hwnd: HWND) {
    let mut pt = POINT::default();
    let _ = GetCursorPos(&mut pt);
    show_rss_context_menu(hwnd, pt.x, pt.y, false);
}

unsafe fn show_rss_context_menu(hwnd: HWND, x: i32, y: i32, use_hit_test: bool) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }

    let mut rect = RECT::default();
    if use_hit_test {
        if GetWindowRect(hwnd_tree, &mut rect).is_ok() {
            if x < rect.left || x > rect.right || y < rect.top || y > rect.bottom {
                return;
            }
        }
    }

    let hitem = if use_hit_test {
        let mut pt = POINT { x, y };
        if rect.right != 0 || rect.bottom != 0 {
            pt.x -= rect.left;
            pt.y -= rect.top;
        }
        let mut hit = TVHITTESTINFO {
            pt,
            ..Default::default()
        };
        let hitem = windows::Win32::UI::Controls::HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_HITTEST,
                WPARAM(0),
                LPARAM(&mut hit as *mut _ as isize),
            )
            .0,
        );
        if hitem.0 != 0 {
            let _ = SendMessageW(
                hwnd_tree,
                TVM_SELECTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(hitem.0),
            );
            let _ = SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));
        }
        hitem
    } else {
        windows::Win32::UI::Controls::HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
            .0,
        )
    };

    if hitem.0 == 0 {
        return;
    }

    let is_source = with_rss_state(hwnd, |s| {
        matches!(s.node_data.get(&hitem.0), Some(NodeData::Source(_)))
    })
    .unwrap_or(false);
    if !is_source {
        return;
    }

    let language = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
    })
    .unwrap_or_default();
    let edit_label = i18n::tr(language, "rss.context.edit");
    let delete_label = i18n::tr(language, "rss.context.delete");

    if let Ok(menu) = CreatePopupMenu() {
        if menu.0 != 0 {
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_EDIT,
                PCWSTR(to_wide(&edit_label).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_DELETE,
                PCWSTR(to_wide(&delete_label).as_ptr()),
            );
            SetForegroundWindow(hwnd);
            let _ = TrackPopupMenu(
                menu,
                windows::Win32::UI::WindowsAndMessaging::TPM_RIGHTBUTTON,
                x,
                y,
                0,
                hwnd,
                None,
            );
            let _ = PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0));
            let _ = DestroyMenu(menu);
        }
    }
}

unsafe extern "system" fn rss_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*cs).lpCreateParams as isize);

            let state = Box::new(RssWindowState {
                parent,
                hwnd_tree: HWND(0),
                node_data: HashMap::new(),
                pending_fetches: HashMap::new(),
                source_items: HashMap::new(),
                enter_guard: false,
                add_guard: false,
                pending_edit: None,
                tree_proc: None,
                last_selected: 0,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            create_controls(hwnd);
            ensure_default_sources(parent);
            reload_tree(hwnd);

            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                let _ = with_state(parent, |s| s.rss_window = HWND(0));
                // Parent was never disabled; just bring it to front as a convenience.
                force_focus_editor_on_parent(parent);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut RssWindowState;
            if !ptr.is_null() {
                let parent = (*ptr).parent;
                let hwnd_tree = (*ptr).hwnd_tree;
                if hwnd_tree.0 != 0 {
                    if let Some(proc) = (*ptr).tree_proc {
                        SetWindowLongPtrW(hwnd_tree, GWLP_WNDPROC, proc as isize);
                    }
                }
                if parent.0 != 0 {
                    force_focus_editor_on_parent(parent);
                }
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            match id {
                ID_BTN_CLOSE | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                ID_BTN_ADD => {
                    // Direct activation (mouse click, or Space/Enter if the button sends BN_CLICKED).
                    // Guard against the common case where Enter also generates IDOK (1),
                    // which would otherwise open the dialog twice and crash.
                    let already = with_rss_state(hwnd, |s| s.add_guard).unwrap_or(false);
                    if !already {
                        with_rss_state(hwnd, |s| s.add_guard = true);
                        let _ = PostMessageW(hwnd, WM_SHOW_ADD_DIALOG, WPARAM(0), LPARAM(0));
                    }
                    LRESULT(0)
                }
                1 => {
                    // IDOK (Enter key often triggers this generic command)
                    let focus = GetFocus();
                    let btn_add = GetDlgItem(hwnd, ID_BTN_ADD as i32);
                    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));

                    if focus == btn_add {
                        let already = with_rss_state(hwnd, |s| s.add_guard).unwrap_or(false);
                        if !already {
                            with_rss_state(hwnd, |s| s.add_guard = true);
                            let _ = PostMessageW(hwnd, WM_SHOW_ADD_DIALOG, WPARAM(0), LPARAM(0));
                        }
                        return LRESULT(0);
                    }

                    if focus == hwnd_tree {
                        let already = with_rss_state(hwnd, |s| s.enter_guard).unwrap_or(false);
                        if !already {
                            with_rss_state(hwnd, |s| s.enter_guard = true);
                            let _ = PostMessageW(hwnd, WM_CLEAR_ENTER_GUARD, WPARAM(0), LPARAM(0));
                            handle_enter(hwnd);
                        }
                        return LRESULT(0);
                    }

                    LRESULT(0)
                }
                ID_CTX_EDIT => {
                    handle_edit_source(hwnd);
                    LRESULT(0)
                }
                ID_CTX_DELETE => {
                    handle_delete(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        windows::Win32::UI::WindowsAndMessaging::WM_COPYDATA => {
            let cds = lparam.0 as *const COPYDATASTRUCT;
            if (*cds).dwData == 0x52535331 {
                let len = ((*cds).cbData / 2) as usize;
                let slice = std::slice::from_raw_parts((*cds).lpData as *const u16, len);
                // Need 0-term
                let s = String::from_utf16_lossy(slice);
                let payload = s.trim_matches(char::from(0)).to_string();
                let mut lines = payload.lines();
                let first = lines.next().unwrap_or("");
                let second = lines.next();
                let (mut title, url) = if let Some(url_line) = second {
                    (first.trim().to_string(), url_line.trim().to_string())
                } else {
                    (String::new(), first.trim().to_string())
                };
                if url.is_empty() {
                    return LRESULT(0);
                }
                if title.trim().is_empty() {
                    title = url.clone();
                }

                let edit_idx = with_rss_state(hwnd, |s| s.pending_edit.take()).unwrap_or(None);
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if let Some(idx) = edit_idx {
                    with_state(parent, |state| {
                        if let Some(src) = state.settings.rss_sources.get_mut(idx) {
                            src.title = title.clone();
                            src.url = url.clone();
                            src.user_title = title.trim() != url.trim();
                            src.cache = rss::RssFeedCache::default();
                        }
                        crate::settings::save_settings(state.settings.clone());
                    });
                    reload_tree(hwnd);
                } else {
                    with_state(parent, |state| {
                        state.settings.rss_sources.push(RssSource {
                            title: title.clone(),
                            url: url.clone(),
                            kind: RssSourceType::Site,
                            user_title: title.trim() != url.trim(),
                            cache: rss::RssFeedCache::default(),
                        });
                        crate::settings::save_settings(state.settings.clone());
                    });
                    reload_tree(hwnd);

                    // Auto-expand the new item to trigger fetch (and title update)
                    let idx = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| ps.settings.rss_sources.len()).unwrap_or(0)
                    })
                    .unwrap_or(0);
                    if idx > 0 {
                        let last_idx = idx - 1;
                        let hitem = with_rss_state(hwnd, |s| {
                            s.node_data.iter().find_map(|(k, v)| {
                                if let NodeData::Source(i) = v {
                                    if *i == last_idx { Some(*k) } else { None }
                                } else {
                                    None
                                }
                            })
                        })
                        .flatten();

                        if let Some(h) = hitem {
                            let hwnd_tree =
                                with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                            SendMessageW(
                                hwnd_tree,
                                TVM_EXPAND,
                                WPARAM(TVE_EXPAND.0 as usize),
                                LPARAM(h),
                            );
                        }
                    }
                }
            }
            LRESULT(0)
        }
        WM_NOTIFY => {
            let nmhdr = lparam.0 as *const NMHDR;
            if (*nmhdr).idFrom == ID_TREE {
                match (*nmhdr).code {
                    TVN_ITEMEXPANDINGW => {
                        let pnmtv = lparam.0 as *const NMTREEVIEWW;
                        // Handle expansion
                        // TVE_EXPAND is u32 or constant?
                        // action is NM_TREEVIEW_ACTION.
                        // We need to check if it equals TVE_EXPAND.
                        // But wait, TVE_EXPAND is action flag?
                        // Assuming action == TVE_EXPAND
                        if (*pnmtv).action == TVE_EXPAND {
                            let hitem = (*pnmtv).itemNew.hItem;
                            handle_expand(hwnd, hitem);
                        }
                        LRESULT(0)
                    }
                    TVN_SELCHANGEDW => {
                        let pnmtv = lparam.0 as *const NMTREEVIEWW;
                        let hitem = (*pnmtv).itemNew.hItem;
                        handle_selection_changed(hwnd, hitem);
                        LRESULT(0)
                    }
                    TVN_KEYDOWN => {
                        let ptvkd = lparam.0 as *const NMTVKEYDOWN;
                        if (*ptvkd).wVKey
                            == windows::Win32::UI::Input::KeyboardAndMouse::VK_RETURN.0
                        {
                            with_rss_state(hwnd, |s| s.enter_guard = true);
                            let _ = PostMessageW(hwnd, WM_CLEAR_ENTER_GUARD, WPARAM(0), LPARAM(0));
                            handle_enter(hwnd);
                            LRESULT(1)
                        } else if (*ptvkd).wVKey
                            == windows::Win32::UI::Input::KeyboardAndMouse::VK_DELETE.0
                        {
                            handle_delete(hwnd);
                            LRESULT(1)
                        } else if (*ptvkd).wVKey == VK_F10.0 && GetKeyState(VK_SHIFT.0 as i32) < 0 {
                            let hwnd_tree =
                                with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                            if hwnd_tree.0 != 0 {
                                let _ = PostMessageW(
                                    hwnd,
                                    WM_CONTEXTMENU,
                                    WPARAM(hwnd_tree.0 as usize),
                                    LPARAM(-1),
                                );
                            }
                            LRESULT(1)
                        } else if (*ptvkd).wVKey == VK_APPS.0 {
                            let hwnd_tree =
                                with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                            if hwnd_tree.0 != 0 {
                                let _ = PostMessageW(
                                    hwnd,
                                    WM_CONTEXTMENU,
                                    WPARAM(hwnd_tree.0 as usize),
                                    LPARAM(-1),
                                );
                            }
                            LRESULT(1)
                        } else if (*ptvkd).wVKey == VK_ESCAPE.0 {
                            let _ = DestroyWindow(hwnd);
                            LRESULT(1)
                        } else {
                            LRESULT(0)
                        }
                    }
                    NM_RCLICK => {
                        let mut pt = POINT::default();
                        let _ = GetCursorPos(&mut pt);
                        show_rss_context_menu(hwnd, pt.x, pt.y, true);
                        LRESULT(1)
                    }
                    _ => LRESULT(0),
                }
            } else {
                LRESULT(0)
            }
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            let key = wparam.0 as u32;
            if key == u32::from(VK_ESCAPE.0) {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
            if hwnd_tree.0 == 0 {
                return LRESULT(0);
            }
            if GetFocus() != hwnd_tree {
                return LRESULT(0);
            }
            if key == u32::from(VK_APPS.0)
                || (key == u32::from(VK_F10.0) && GetKeyState(VK_SHIFT.0 as i32) < 0)
            {
                let _ = PostMessageW(
                    hwnd,
                    WM_CONTEXTMENU,
                    WPARAM(hwnd_tree.0 as usize),
                    LPARAM(-1),
                );
                return LRESULT(0);
            }
            if key == u32::from(VK_ESCAPE.0) {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CONTEXTMENU => {
            let mut x = (lparam.0 & 0xffff) as i32;
            let mut y = ((lparam.0 >> 16) & 0xffff) as i32;
            let use_hit_test = !(x == -1 && y == -1);
            if x == -1 && y == -1 {
                let mut pt = POINT::default();
                let _ = GetCursorPos(&mut pt);
                x = pt.x;
                y = pt.y;
            }
            show_rss_context_menu(hwnd, x, y, use_hit_test);
            LRESULT(0)
        }
        WM_RSS_FETCH_COMPLETE => {
            let ptr = lparam.0 as *mut FetchResult;
            let res = *Box::from_raw(ptr);
            process_fetch_result(hwnd, res);
            LRESULT(0)
        }
        WM_CLEAR_ENTER_GUARD => {
            with_rss_state(hwnd, |s| s.enter_guard = false);
            LRESULT(0)
        }
        WM_CLEAR_ADD_GUARD => {
            with_rss_state(hwnd, |s| s.add_guard = false);
            LRESULT(0)
        }

        WM_TIMER => {
            // Clear short-lived guards (prevents double-open on Enter on the Add button)
            if wparam.0 as usize == ADD_GUARD_TIMER_ID {
                with_rss_state(hwnd, |s| s.add_guard = false);
                let _ = KillTimer(hwnd, ADD_GUARD_TIMER_ID);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }

        WM_RSS_IMPORT_COMPLETE => {
            let ptr = lparam.0 as *mut ImportResult;
            let res = Box::from_raw(ptr);

            let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            let mut hwnd_edit = crate::get_active_edit(parent);

            if hwnd_edit.is_none() {
                SendMessageW(
                    parent,
                    WM_COMMAND,
                    WPARAM(crate::menu::IDM_FILE_NEW),
                    LPARAM(0),
                );
                hwnd_edit = crate::get_active_edit(parent);
            }

            if let Some(h_edit) = hwnd_edit {
                // Bring the main window to the front *before* moving focus.
                // Doing it the other way around is frequently ignored by Windows,
                // especially when focus changes originate from posted messages.
                let main_window = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if main_window.0 != 0 {
                    SetForegroundWindow(main_window);
                }
                // Ensure large articles are not truncated by the default edit limit.
                SendMessageW(h_edit, EM_LIMITTEXT, WPARAM(0x7FFF_FFFEusize), LPARAM(0));

                // Replace the entire editor contents, then move caret to the start.
                // We normalize text to avoid embedded NULs (which truncate Win32 edit text)
                // and to reduce multiple blank lines.
                let cleaned = normalize_article_text(&res.text);
                let wide = to_wide(&cleaned);
                SendMessageW(h_edit, EM_SETSEL, WPARAM(0), LPARAM(-1));
                SendMessageW(
                    h_edit,
                    EM_REPLACESEL,
                    WPARAM(1),
                    LPARAM(wide.as_ptr() as isize),
                );
                SendMessageW(h_edit, EM_SETSEL, WPARAM(0), LPARAM(0));

                SetFocus(h_edit);
            }
            LRESULT(0)
        }
        WM_SHOW_ADD_DIALOG => {
            // If the add dialog is already open, just bring it to the front.
            let main_hwnd = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            let existing = with_state(main_hwnd, |s| s.rss_add_dialog).unwrap_or(HWND(0));
            if existing.0 != 0 {
                SetForegroundWindow(existing);
            } else {
                show_add_dialog(hwnd);
            }
            let _ = PostMessageW(hwnd, WM_CLEAR_ADD_GUARD, WPARAM(0), LPARAM(0));
            LRESULT(0)
        }
        WM_RSS_SHOW_CONTEXT => {
            show_context_menu_from_keyboard(hwnd);
            LRESULT(0)
        }
        WM_SETFOCUS => {
            let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
            if hwnd_tree.0 != 0 {
                SetFocus(hwnd_tree);
                let last = with_rss_state(hwnd, |s| s.last_selected).unwrap_or(0);
                if last != 0 {
                    let _ = SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(last),
                    );
                    let _ = SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(last));
                }
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_rss_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut RssWindowState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut RssWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe extern "system" fn rss_tree_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN {
        let key = wparam.0 as u32;
        if key == u32::from(VK_APPS.0)
            || (key == u32::from(VK_F10.0) && GetKeyState(VK_SHIFT.0 as i32) < 0)
        {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                let _ = PostMessageW(parent, WM_CONTEXTMENU, WPARAM(hwnd.0 as usize), LPARAM(-1));
                return LRESULT(0);
            }
        }
    }

    let parent = GetParent(hwnd);
    let prev_proc = if parent.0 != 0 {
        with_rss_state(parent, |s| s.tree_proc).unwrap_or(None)
    } else {
        None
    };
    if let Some(proc) = prev_proc {
        CallWindowProcW(Some(proc), hwnd, msg, wparam, lparam)
    } else {
        DefWindowProcW(hwnd, msg, wparam, lparam)
    }
}

unsafe fn create_controls(hwnd: HWND) {
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);

    let hwnd_tree = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("SysTreeView32"),
        PCWSTR::null(),
        WS_CHILD
            | WS_VISIBLE
            | WS_TABSTOP
            | WS_VSCROLL
            | WINDOW_STYLE(
                (windows::Win32::UI::Controls::TVS_HASLINES
                    | windows::Win32::UI::Controls::TVS_HASBUTTONS
                    | windows::Win32::UI::Controls::TVS_LINESATROOT
                    | windows::Win32::UI::Controls::TVS_SHOWSELALWAYS) as u32,
            ),
        10,
        10,
        460,
        500,
        hwnd,
        HMENU(ID_TREE as isize),
        hinstance,
        None,
    );
    if hwnd_tree.0 != 0 {
        let old = SetWindowLongPtrW(hwnd_tree, GWLP_WNDPROC, rss_tree_wndproc as isize);
        with_rss_state(hwnd, |s| {
            s.tree_proc = mem::transmute::<isize, WNDPROC>(old)
        });
    }

    let language = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
    })
    .unwrap_or_default();

    let hwnd_add = CreateWindowExW(
        Default::default(),
        WC_BUTTON,
        PCWSTR(to_wide(&i18n::tr(language, "rss.tree.add_source")).as_ptr()),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        520,
        100,
        30,
        hwnd,
        HMENU(ID_BTN_ADD as isize),
        hinstance,
        None,
    );

    let hwnd_close = CreateWindowExW(
        Default::default(),
        WC_BUTTON,
        PCWSTR(to_wide(&i18n::tr(language, "rss.tree.close")).as_ptr()),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        370,
        520,
        100,
        30,
        hwnd,
        HMENU(ID_BTN_CLOSE as isize),
        hinstance,
        None,
    );

    with_rss_state(hwnd, |s| s.hwnd_tree = hwnd_tree);

    let hfont = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| ps.hfont).unwrap_or(HFONT(0))
    })
    .unwrap_or(HFONT(0));
    if hfont.0 != 0 {
        SendMessageW(hwnd_tree, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        SendMessageW(hwnd_add, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        SendMessageW(hwnd_close, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
    }

    SetFocus(hwnd_tree);
}

unsafe fn reload_tree(hwnd: HWND) {
    let (hwnd_tree, sources) = match with_rss_state(hwnd, |s| {
        (
            s.hwnd_tree,
            with_state(s.parent, |ps| ps.settings.rss_sources.clone()),
        )
    }) {
        Some((t, Some(src))) => (t, src),
        _ => return,
    };

    SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(TVI_ROOT.0));

    with_rss_state(hwnd, |s| {
        s.node_data.clear();
        s.source_items.clear();
    });

    for (i, source) in sources.into_iter().enumerate() {
        let title = to_wide(&source.title);
        let mut tvis = TVINSERTSTRUCTW {
            hParent: TVI_ROOT,
            hInsertAfter: TVI_LAST,
            Anonymous: TVINSERTSTRUCTW_0 {
                item: TVITEMW {
                    mask: TVIF_TEXT | TVIF_PARAM | windows::Win32::UI::Controls::TVIF_CHILDREN,
                    pszText: windows::core::PWSTR(title.as_ptr() as *mut _),
                    cChildren: TVITEMEXW_CHILDREN(1),
                    lParam: LPARAM(i as isize),
                    ..Default::default()
                },
            },
        };
        let hitem = SendMessageW(
            hwnd_tree,
            TVM_INSERTITEMW,
            WPARAM(0),
            LPARAM(&mut tvis as *mut _ as isize),
        );

        with_rss_state(hwnd, |s| {
            s.node_data.insert(hitem.0, NodeData::Source(i));
        });
    }
}

unsafe fn handle_expand(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM) {
    let item_info_opt = with_rss_state(hwnd, |s| {
        if let Some(NodeData::Source(idx)) = s.node_data.get(&(hitem.0)) {
            with_state(s.parent, |ps| {
                ps.settings
                    .rss_sources
                    .get(*idx)
                    .map(|src| (src.url.clone(), src.kind.clone(), src.cache.clone(), true))
            })
            .flatten()
        } else if let Some(NodeData::Item(item)) = s.node_data.get(&(hitem.0)) {
            if item.is_folder {
                Some((
                    item.link.clone(),
                    RssSourceType::Site,
                    rss::RssFeedCache::default(),
                    false,
                ))
            } else {
                None
            }
        } else {
            None
        }
    });

    let (url, source_kind, cache, _is_source) = if let Some(info) = item_info_opt.flatten() {
        info
    } else {
        return;
    };

    with_rss_state(hwnd, |s| {
        s.pending_fetches.insert(url.clone(), hitem.0);
    });

    let url_clone = url.clone();

    // Ensure the node expands immediately for keyboard users (Right Arrow),
    // even when children are populated asynchronously.
    // If there are no children yet, insert a temporary "Loading…" child.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 != 0 {
        let first_child = SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CHILD as usize),
            LPARAM(hitem.0),
        );
        if first_child.0 == 0 {
            let mut loading_label = "Loading...".to_string();
            if loading_label.trim().is_empty() {
                loading_label = "Loading...".to_string();
            }
            let loading_txt = to_wide(&loading_label);
            let mut tvis_loading = TVINSERTSTRUCTW {
                hParent: hitem,
                hInsertAfter: TVI_LAST,
                Anonymous: windows::Win32::UI::Controls::TVINSERTSTRUCTW_0 {
                    item: TVITEMW {
                        mask: TVIF_TEXT,
                        pszText: windows::core::PWSTR(loading_txt.as_ptr() as *mut _),
                        cchTextMax: loading_txt.len() as i32,
                        ..Default::default()
                    },
                },
            };
            let _ = SendMessageW(
                hwnd_tree,
                TVM_INSERTITEMW,
                WPARAM(0),
                LPARAM(&mut tvis_loading as *mut _ as isize),
            );
        }
        // Force visual expansion now.
        let _ = SendMessageW(
            hwnd_tree,
            TVM_EXPAND,
            WPARAM(TVE_EXPAND.0 as usize),
            LPARAM(hitem.0),
        );
        let _ = SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));
    }

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let fetch_config = if parent.0 != 0 {
        rss_fetch_config(parent)
    } else {
        rss::RssFetchConfig::default()
    };
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    // UI: "Refresh feeds" should trigger this fetch for the selected source.
    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let res = rt.block_on(rss::fetch_and_parse(
            &url_clone,
            source_kind,
            cache,
            fetch_config,
        ));
        let msg = Box::new(FetchResult {
            hitem: hitem.0,
            result: res,
        });
        let _ = PostMessageW(
            hwnd,
            WM_RSS_FETCH_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        );
    });
}

struct FetchResult {
    hitem: isize,
    result: Result<rss::RssFetchOutcome, String>,
}

unsafe fn process_fetch_result(hwnd: HWND, res: FetchResult) {
    let hitem = windows::Win32::UI::Controls::HTREEITEM(res.hitem);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));

    // Preserve selection if it is a child of this item
    let selected = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );
    let mut selected_idx = None;
    if selected.0 != 0 {
        // Check if selected is child of hitem
        let parent_of_selected = windows::Win32::UI::Controls::HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(windows::Win32::UI::Controls::TVGN_PARENT as usize),
                LPARAM(selected.0),
            )
            .0,
        );
        if parent_of_selected == hitem {
            // It is a child. Find its index.
            let mut idx = 0;
            let mut child = windows::Win32::UI::Controls::HTREEITEM(
                SendMessageW(
                    hwnd_tree,
                    TVM_GETNEXTITEM,
                    WPARAM(TVGN_CHILD as usize),
                    LPARAM(hitem.0),
                )
                .0,
            );
            while child.0 != 0 {
                if child == selected {
                    selected_idx = Some(idx);
                    break;
                }
                child = windows::Win32::UI::Controls::HTREEITEM(
                    SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(windows::Win32::UI::Controls::TVGN_NEXT as usize),
                        LPARAM(child.0),
                    )
                    .0,
                );
                idx += 1;
            }
        }
    }

    match res.result {
        Ok(outcome) => {
            // Update source title if applicable
            let is_source_node = with_rss_state(hwnd, |s| {
                s.node_data.contains_key(&hitem.0)
                    && matches!(s.node_data[&hitem.0], NodeData::Source(_))
            })
            .unwrap_or(false);
            if is_source_node {
                let idx = with_rss_state(hwnd, |s| {
                    if let NodeData::Source(i) = s.node_data[&hitem.0] {
                        Some(i)
                    } else {
                        None
                    }
                })
                .flatten();
                if let Some(i) = idx {
                    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    let mut final_title = outcome.title.clone();
                    with_state(parent, |ps| {
                        let lang = ps.settings.language;
                        let (_key, keep_default_title) = ps
                            .settings
                            .rss_sources
                            .get(i)
                            .map(|src| {
                                let key = normalize_rss_url_key(&src.url);
                                let keep = is_default_key(lang, &ps.settings, &key);
                                (key, keep)
                            })
                            .unwrap_or_default();
                        if let Some(src) = ps.settings.rss_sources.get_mut(i) {
                            let looks_auto = src.title.trim().is_empty() || src.title == src.url;
                            if !src.user_title
                                && !keep_default_title
                                && looks_auto
                                && !outcome.title.is_empty()
                            {
                                src.title = outcome.title.clone();
                            }
                            final_title = src.title.clone();
                            if src.kind != outcome.kind {
                                src.kind = outcome.kind;
                            }
                            src.cache = outcome.cache.clone();
                        }
                        crate::settings::save_settings(ps.settings.clone());
                    });

                    let title_wide = to_wide(&final_title);
                    let mut tvi = TVITEMW {
                        mask: TVIF_TEXT,
                        hItem: hitem,
                        pszText: windows::core::PWSTR(title_wide.as_ptr() as *mut _),
                        ..Default::default()
                    };
                    SendMessageW(
                        hwnd_tree,
                        TVM_SETITEMW,
                        WPARAM(0),
                        LPARAM(&mut tvi as *mut _ as isize),
                    );
                }
            }

            if outcome.not_modified {
                return;
            }

            loop {
                let child = SendMessageW(
                    hwnd_tree,
                    TVM_GETNEXTITEM,
                    WPARAM(TVGN_CHILD as usize),
                    LPARAM(hitem.0),
                );
                if child.0 == 0 {
                    break;
                }
                SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(child.0));
            }
            with_rss_state(hwnd, |s| {
                s.source_items.remove(&hitem.0);
            });

            with_rss_state(hwnd, |s| {
                s.source_items.insert(
                    hitem.0,
                    SourceItemsState {
                        items: outcome.items,
                        loaded: 0,
                    },
                );
            });
            let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            let (initial_count, _next_count) = rss_page_sizes(parent);
            let inserted = load_more_items(hwnd, hitem, initial_count);

            // Force expansion/visibility now that children are populated.
            // This prevents cases where the node appears to expand only after moving selection.
            SendMessageW(
                hwnd_tree,
                TVM_EXPAND,
                WPARAM(TVE_EXPAND.0 as usize),
                LPARAM(hitem.0),
            );
            SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));

            // Restore selection only if focus stays on this tree item.
            if let Some(idx) = selected_idx {
                let focused = GetFocus();
                let caret = windows::Win32::UI::Controls::HTREEITEM(
                    SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(0),
                    )
                    .0,
                );
                if focused == hwnd_tree && caret == hitem && idx < inserted {
                    // specific restoration might be hard because handlers changed.
                    // We can try to find the Nth child.
                    let mut child = SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(TVGN_CHILD as usize),
                        LPARAM(hitem.0),
                    );
                    for _ in 0..idx {
                        if child.0 == 0 {
                            break;
                        }
                        child = SendMessageW(
                            hwnd_tree,
                            TVM_GETNEXTITEM,
                            WPARAM(windows::Win32::UI::Controls::TVGN_NEXT as usize),
                            LPARAM(child.0),
                        );
                    }
                    if child.0 != 0 {
                        SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(child.0),
                        );
                    }
                }
            }
        }
        Err(e) => {
            loop {
                let child = SendMessageW(
                    hwnd_tree,
                    TVM_GETNEXTITEM,
                    WPARAM(TVGN_CHILD as usize),
                    LPARAM(hitem.0),
                );
                if child.0 == 0 {
                    break;
                }
                SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(child.0));
            }
            with_rss_state(hwnd, |s| {
                s.source_items.remove(&hitem.0);
            });
            if e.to_ascii_lowercase().contains("resource limit") {
                return;
            }
            let text = to_wide(&format!("Error: {}", e));
            let mut tvis = TVINSERTSTRUCTW {
                hParent: hitem,
                hInsertAfter: TVI_LAST,
                Anonymous: TVINSERTSTRUCTW_0 {
                    item: TVITEMW {
                        mask: TVIF_TEXT,
                        pszText: windows::core::PWSTR(text.as_ptr() as *mut _),
                        ..Default::default()
                    },
                },
            };
            SendMessageW(
                hwnd_tree,
                TVM_INSERTITEMW,
                WPARAM(0),
                LPARAM(&mut tvis as *mut _ as isize),
            );
        }
    }
}

unsafe fn handle_selection_changed(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM) {
    if hitem.0 == 0 {
        return;
    }
    with_rss_state(hwnd, |s| s.last_selected = hitem.0);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let parent = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(windows::Win32::UI::Controls::TVGN_PARENT as usize),
            LPARAM(hitem.0),
        )
        .0,
    );
    if parent.0 == 0 {
        return;
    }
    let has_more = with_rss_state(hwnd, |s| {
        s.source_items
            .get(&parent.0)
            .map(|state| state.loaded < state.items.len())
            .unwrap_or(false)
    })
    .unwrap_or(false);
    if !has_more {
        return;
    }
    let child = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CHILD as usize),
            LPARAM(parent.0),
        )
        .0,
    );
    if child.0 == 0 {
        return;
    }
    let mut last = child;
    loop {
        let next = windows::Win32::UI::Controls::HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(windows::Win32::UI::Controls::TVGN_NEXT as usize),
                LPARAM(last.0),
            )
            .0,
        );
        if next.0 == 0 {
            break;
        }
        last = next;
    }
    if hitem == last {
        let parent_hwnd = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
        let (_initial_count, next_count) = rss_page_sizes(parent_hwnd);
        let _ = load_more_items(hwnd, parent, next_count);
    }
}

unsafe fn load_more_items(
    hwnd: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
    batch: usize,
) -> usize {
    // UI: "Load more titles" can call this to append the next page locally.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return 0;
    }
    let _ = SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(0), LPARAM(0));
    let inserted = with_rss_state(hwnd, |s| {
        let Some(state) = s.source_items.get_mut(&hitem.0) else {
            return 0;
        };
        if state.loaded >= state.items.len() {
            return 0;
        }
        let mut inserted = 0usize;
        let mut idx = state.loaded;
        while idx < state.items.len() && inserted < batch {
            let item = &state.items[idx];
            idx += 1;
            if item.title.trim().is_empty() {
                continue;
            }
            let text = to_wide(&item.title);
            let c_children = if item.is_folder { 1 } else { 0 };
            let mut tvis = TVINSERTSTRUCTW {
                hParent: hitem,
                hInsertAfter: TVI_LAST,
                Anonymous: TVINSERTSTRUCTW_0 {
                    item: TVITEMW {
                        mask: TVIF_TEXT | TVIF_PARAM | windows::Win32::UI::Controls::TVIF_CHILDREN,
                        pszText: windows::core::PWSTR(text.as_ptr() as *mut _),
                        cChildren: TVITEMEXW_CHILDREN(c_children),
                        lParam: LPARAM(0),
                        ..Default::default()
                    },
                },
            };
            let hchild = SendMessageW(
                hwnd_tree,
                TVM_INSERTITEMW,
                WPARAM(0),
                LPARAM(&mut tvis as *mut _ as isize),
            );
            s.node_data.insert(hchild.0, NodeData::Item(item.clone()));
            inserted += 1;
        }
        state.loaded = idx;
        inserted
    })
    .unwrap_or(0);
    let _ = SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
    inserted
}

unsafe fn handle_enter(hwnd: HWND) {
    // UI: On Enter, fetch article content and import into the editor.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );
    if hitem.0 == 0 {
        return;
    }

    let item_opt = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Item(item)) if !item.is_folder => Some(item.clone()),
        _ => None,
    })
    .flatten();

    if let Some(item) = item_opt {
        import_item(hwnd, item);
    } else {
        SendMessageW(
            hwnd_tree,
            TVM_EXPAND,
            WPARAM(TVE_EXPAND.0 as usize),
            LPARAM(hitem.0),
        );
    }
}

unsafe fn handle_delete(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );
    if hitem.0 == 0 {
        return;
    }

    let source_idx = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten();

    if let Some(idx) = source_idx {
        let source_info = with_rss_state(hwnd, |s| {
            with_state(s.parent, |ps| {
                ps.settings
                    .rss_sources
                    .get(idx)
                    .map(|src| (src.title.clone(), src.url.clone()))
            })
            .flatten()
        })
        .flatten();
        let (title, url) = source_info.unwrap_or_default();

        // Localize message and title
        let language = with_rss_state(hwnd, |s| {
            with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
        })
        .unwrap_or_default();
        let msg_template = i18n::tr(language, "rss.delete_confirm");
        let msg_text = msg_template.replace("{title}", &title);
        let caption = i18n::tr(language, "rss.delete_title");

        let ret = MessageBoxW(
            hwnd,
            PCWSTR(to_wide(&msg_text).as_ptr()),
            PCWSTR(to_wide(&caption).as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        );
        if ret == IDYES {
            let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            with_state(parent, |ps| {
                if matches!(
                    language,
                    crate::settings::Language::English
                        | crate::settings::Language::Italian
                        | crate::settings::Language::Spanish
                        | crate::settings::Language::Portuguese
                ) {
                    let defaults = load_default_feeds(language);
                    if !defaults.is_empty() {
                        let mut default_keys = HashSet::new();
                        for (_title, url) in defaults {
                            let key = normalize_rss_url_key(&url);
                            if !key.is_empty() {
                                default_keys.insert(key);
                            }
                        }
                        let key = normalize_rss_url_key(&url);
                        if !key.is_empty() && default_keys.contains(&key) {
                            let removed_list = match language {
                                crate::settings::Language::English => {
                                    &mut ps.settings.rss_removed_default_en
                                }
                                crate::settings::Language::Italian => {
                                    &mut ps.settings.rss_removed_default_it
                                }
                                crate::settings::Language::Spanish => {
                                    &mut ps.settings.rss_removed_default_es
                                }
                                crate::settings::Language::Portuguese => {
                                    &mut ps.settings.rss_removed_default_pt
                                }
                            };
                            let already =
                                removed_list.iter().any(|u| normalize_rss_url_key(u) == key);
                            if !already {
                                removed_list.push(key);
                            }
                        }
                    }
                }
                ps.settings.rss_sources.remove(idx);
                crate::settings::save_settings(ps.settings.clone());
            });
            reload_tree(hwnd);
        }
    }
    if hwnd_tree.0 != 0 {
        SetForegroundWindow(hwnd);
        SetFocus(hwnd_tree);
        let _ = SendMessageW(hwnd_tree, WM_SETFOCUS, WPARAM(0), LPARAM(0));
    }
}

unsafe fn handle_edit_source(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );
    if hitem.0 == 0 {
        return;
    }

    let source_idx = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten();

    let Some(idx) = source_idx else {
        return;
    };

    let source_info = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| {
            ps.settings
                .rss_sources
                .get(idx)
                .map(|src| (src.title.clone(), src.url.clone()))
        })
        .flatten()
    })
    .flatten();
    let (title, url) = source_info.unwrap_or_default();
    if url.trim().is_empty() {
        return;
    }

    let main_hwnd = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let existing = with_state(main_hwnd, |s| s.rss_add_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    with_rss_state(hwnd, |s| s.pending_edit = Some(idx));
    show_add_dialog_with_prefill(hwnd, title, url);
}

unsafe fn force_focus_editor_on_parent(parent: HWND) {
    if parent.0 == 0 {
        return;
    }
    SetForegroundWindow(parent);
    SetActiveWindow(parent);
    let _ = SendMessageW(parent, WM_SETFOCUS, WPARAM(0), LPARAM(0));
    if crate::get_active_edit(parent).is_none() {
        let _ = SendMessageW(
            parent,
            WM_COMMAND,
            WPARAM(crate::menu::IDM_FILE_NEW),
            LPARAM(0),
        );
    }
    if let Some(hwnd_edit) = crate::get_active_edit(parent) {
        SetFocus(hwnd_edit);
        let _ = SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(
            parent,
            WM_NEXTDLGCTL,
            WPARAM(hwnd_edit.0 as usize),
            LPARAM(1),
        );
        // Re-assert focus after dialog navigation to help NVDA settle on the edit control.
        SetFocus(hwnd_edit);
        let _ = SendMessageW(hwnd_edit, EM_SETSEL, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(hwnd_edit, EM_SCROLLCARET, WPARAM(0), LPARAM(0));
        let _ = SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
        let _ = NotifyWinEvent(
            EVENT_OBJECT_FOCUS,
            hwnd_edit,
            OBJID_CLIENT.0,
            CHILDID_SELF as i32,
        );
    }
    let _ = SendMessageW(parent, WM_SETFOCUS, WPARAM(0), LPARAM(0));
    let _ = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
}

unsafe fn import_item(hwnd: HWND, item: RssItem) {
    let url = item.link.clone();

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let content_res = rt.block_on(rss::fetch_article_text(
            &url,
            &item.title,
            &item.description,
        ));

        if let Ok(text) = content_res {
            let msg = Box::new(ImportResult { text });
            let _ = PostMessageW(
                hwnd,
                WM_RSS_IMPORT_COMPLETE,
                WPARAM(0),
                LPARAM(Box::into_raw(msg) as isize),
            );
        }
    });
}

struct ImportResult {
    text: String,
}

/// Collapse multiple consecutive blank (or whitespace-only) lines into a single blank line.
/// This improves readability for screen-reader users and keeps the editor content compact.
fn collapse_blank_lines(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let mut prev_blank = false;

    for line in input.lines() {
        let is_blank = line.trim().is_empty();

        if is_blank {
            if prev_blank {
                continue;
            }
            prev_blank = true;
            out.push_str("\n");
        } else {
            prev_blank = false;
            out.push_str(line);
            out.push('\n');
        }
    }

    // If the input ended without a newline, `lines()` won't tell us that.
    // Keep behavior stable by not forcing an extra newline when the output is empty.
    if out.is_empty() { String::new() } else { out }
}

unsafe fn show_add_dialog(parent_hwnd: HWND) {
    show_add_dialog_with_prefill(parent_hwnd, String::new(), String::new());
}

unsafe fn show_add_dialog_with_prefill(parent_hwnd: HWND, title: String, url: String) {
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide("NovapadInput");

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            windows::Win32::UI::WindowsAndMessaging::LoadCursorW(
                None,
                windows::Win32::UI::WindowsAndMessaging::IDC_ARROW,
            )
            .unwrap_or_default()
            .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(input_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let main_hwnd = with_rss_state(parent_hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = with_state(main_hwnd, |s| s.settings.language).unwrap_or_default();
    let init_ptr = Box::into_raw(Box::new(AddDialogInit {
        parent: parent_hwnd,
        prefill_title: title,
        prefill_url: url,
    }));
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide(&i18n::tr(language, "rss.add_dialog.title")).as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_POPUP,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        400,
        190,
        parent_hwnd,
        None,
        hinstance,
        Some(init_ptr as *const _),
    );
    if hwnd.0 == 0 {
        let _ = Box::from_raw(init_ptr);
    }

    let main_window = with_rss_state(parent_hwnd, |s| s.parent).unwrap_or(HWND(0));
    with_state(main_window, |s| s.rss_add_dialog = hwnd);
}

unsafe extern "system" fn input_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*cs).lpCreateParams as *mut AddDialogInit;
            let (parent, prefill_title, prefill_url) = if init_ptr.is_null() {
                (HWND(0), String::new(), String::new())
            } else {
                let init = Box::from_raw(init_ptr);
                (init.parent, init.prefill_title, init.prefill_url)
            };
            // We need language. But we can't easily pass it.
            // We can get it from parent (rss_window) -> parent (main)
            let main_hwnd = with_rss_state(parent, |s| s.parent).unwrap_or(HWND(0));
            let language = with_state(main_hwnd, |s| s.settings.language).unwrap_or_default();

            let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
            // URL label
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.url_label")).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                10,
                360,
                16,
                hwnd,
                HMENU(105),
                hinstance,
                None,
            );
            // URL edit
            CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WINDOW_STYLE(windows::Win32::UI::Controls::PGS_AUTOSCROLL as u32),
                10,
                28,
                360,
                24,
                hwnd,
                HMENU(101),
                hinstance,
                None,
            );
            // Title label
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.title_label")).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                58,
                360,
                16,
                hwnd,
                HMENU(106),
                hinstance,
                None,
            );
            // Title edit
            CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WINDOW_STYLE(windows::Win32::UI::Controls::PGS_AUTOSCROLL as u32),
                10,
                76,
                360,
                24,
                hwnd,
                HMENU(104),
                hinstance,
                None,
            );
            // OK
            CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.ok")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                180,
                120,
                90,
                24,
                hwnd,
                HMENU(102),
                hinstance,
                None,
            );
            // Cancel
            CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.cancel")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                280,
                120,
                90,
                24,
                hwnd,
                HMENU(103),
                hinstance,
                None,
            );
            if !prefill_url.trim().is_empty() {
                let _ = SetWindowTextW(
                    GetDlgItem(hwnd, 101),
                    PCWSTR(to_wide(&prefill_url).as_ptr()),
                );
            }
            if !prefill_title.trim().is_empty() {
                let _ = SetWindowTextW(
                    GetDlgItem(hwnd, 104),
                    PCWSTR(to_wide(&prefill_title).as_ptr()),
                );
            }
            SetFocus(GetDlgItem(hwnd, 101));
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            match id {
                1 => {
                    // IDOK (Enter). This window is not a real DialogBox, so we map Enter to our OK button.
                    // Re-dispatch as if OK (102) was pressed.
                    SendMessageW(hwnd, WM_COMMAND, WPARAM(102), LPARAM(0));
                    LRESULT(0)
                }
                102 => {
                    // OK
                    let h_edit_url = GetDlgItem(hwnd, 101);
                    let len = SendMessageW(
                        h_edit_url,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXTLENGTH,
                        WPARAM(0),
                        LPARAM(0),
                    )
                    .0;
                    let mut buf = vec![0u16; len as usize + 1];
                    SendMessageW(
                        h_edit_url,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXT,
                        WPARAM(buf.len()),
                        LPARAM(buf.as_mut_ptr() as isize),
                    );
                    let url = from_wide(buf.as_ptr());

                    let h_edit_title = GetDlgItem(hwnd, 104);
                    let tlen = SendMessageW(
                        h_edit_title,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXTLENGTH,
                        WPARAM(0),
                        LPARAM(0),
                    )
                    .0;
                    let mut tbuf = vec![0u16; tlen as usize + 1];
                    SendMessageW(
                        h_edit_title,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXT,
                        WPARAM(tbuf.len()),
                        LPARAM(tbuf.as_mut_ptr() as isize),
                    );
                    let title = from_wide(tbuf.as_ptr());

                    if !url.trim().is_empty() {
                        let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);

                        let payload = format!("{}\n{}", title.trim(), url.trim());
                        let url_wide = to_wide(&payload);
                        let cds = COPYDATASTRUCT {
                            dwData: 0x52535331,
                            cbData: (url_wide.len() * 2) as u32,
                            lpData: url_wide.as_ptr() as *mut _,
                        };
                        SendMessageW(
                            parent,
                            windows::Win32::UI::WindowsAndMessaging::WM_COPYDATA,
                            WPARAM(hwnd.0 as usize),
                            LPARAM(&cds as *const _ as isize),
                        );
                    }
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                103 | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_DESTROY => {
            // We can't access AppState directly from here easily without exact parent link.
            // But we know parent is rss_window.
            // However, rss_window state has parent.
            // Let's rely on GWL_USERDATA of parent if possible?
            // Actually, when show_add_dialog creates it, it passes parent_hwnd as parent in CreateWindowEx
            let desktop = windows::Win32::UI::WindowsAndMessaging::GetDesktopWindow();
            let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);

            if parent != desktop && parent.0 != 0 {
                // This parent is rss_window
                // We need to reach main window to clear rss_add_dialog
                // rss_window stores its state in GWLP_USERDATA
                let main_hwnd = with_rss_state(parent, |s| s.parent).unwrap_or(HWND(0));
                if main_hwnd.0 != 0 {
                    with_state(main_hwnd, |s| s.rss_add_dialog = HWND(0));
                }
                with_rss_state(parent, |s| s.pending_edit = None);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}
