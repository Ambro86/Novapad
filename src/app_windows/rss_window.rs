use crate::accessibility::{from_wide, nvda_speak, to_wide};

use crate::editor_manager;
use crate::i18n;
use crate::log_debug;
use crate::tools::rss::{self, RssItem, RssSource, RssSourceType};
use crate::with_state;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::{HashMap, HashSet};
use std::mem;
use std::path::{Path, PathBuf};
use url::Url;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::DataExchange::COPYDATASTRUCT;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_PATHMUSTEXIST,
    OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    NM_RCLICK, NMHDR, NMTREEVIEWW, NMTVKEYDOWN, TVE_EXPAND, TVGN_CARET, TVGN_CHILD, TVGN_NEXT,
    TVGN_ROOT, TVHITTESTINFO, TVI_LAST, TVI_ROOT, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW,
    TVINSERTSTRUCTW_0, TVITEMEXW_CHILDREN, TVITEMW, TVM_DELETEITEM, TVM_ENSUREVISIBLE, TVM_EXPAND,
    TVM_GETITEMW, TVM_GETNEXTITEM, TVM_HITTEST, TVM_INSERTITEMW, TVM_SELECTITEM, TVM_SETITEMW,
    TVM_SORTCHILDRENCB, TVN_ITEMEXPANDINGW, TVN_KEYDOWN, TVN_SELCHANGEDW, TVSORTCB, WC_BUTTON,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetActiveWindow, SetFocus, VK_APPS, VK_ESCAPE, VK_F10, VK_RETURN,
    VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BS_DEFPUSHBUTTON, CHILDID_SELF, CREATESTRUCTW, CW_USEDEFAULT, CallWindowProcW,
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DestroyWindow,
    EVENT_OBJECT_FOCUS, GWLP_USERDATA, GWLP_WNDPROC, GetCursorPos, GetDlgCtrlID, GetDlgItem,
    GetParent, GetWindowLongPtrW, GetWindowRect, HMENU, IDYES, KillTimer, MB_ICONQUESTION,
    MB_YESNO, MF_GRAYED, MF_POPUP, MF_STRING, MessageBoxW, OBJID_CLIENT, PostMessageW,
    RegisterClassW, SW_SHOW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW,
    TrackPopupMenu, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU, WM_CREATE, WM_DESTROY,
    WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL, WM_NOTIFY, WM_NULL, WM_SETFOCUS, WM_SETFONT,
    WM_SETREDRAW, WM_SYSKEYDOWN, WM_TIMER, WM_USER, WNDCLASSW, WNDPROC, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_DLGMODALFRAME, WS_POPUP, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR, w};

const RSS_WINDOW_CLASS: &str = "NovapadRssWindow";
const ID_TREE: usize = 1001;
const ID_BTN_ADD: usize = 1002;
const ID_BTN_CLOSE: usize = 1003;
const ID_BTN_IMPORT: usize = 1004;
const ID_CTX_EDIT: usize = 1101;
const ID_CTX_DELETE: usize = 1102;
const ID_CTX_RETRY: usize = 1103;
const ID_CTX_REORDER_UP: usize = 1301;
const ID_CTX_REORDER_DOWN: usize = 1302;
const ID_CTX_REORDER_TOP: usize = 1303;
const ID_CTX_REORDER_BOTTOM: usize = 1304;
const ID_CTX_REORDER_POSITION: usize = 1305;
const ID_CTX_OPEN_BROWSER: usize = 1201;
const ID_CTX_SHARE_FACEBOOK: usize = 1202;
const ID_CTX_SHARE_TWITTER: usize = 1203;
const ID_CTX_SHARE_WHATSAPP: usize = 1204;

const WM_RSS_FETCH_COMPLETE: u32 = WM_USER + 200;
const WM_RSS_IMPORT_COMPLETE: u32 = WM_USER + 201;
const WM_SHOW_ADD_DIALOG: u32 = WM_USER + 202;
const WM_CLEAR_ENTER_GUARD: u32 = WM_USER + 203;
const WM_CLEAR_ADD_GUARD: u32 = WM_USER + 204;
pub(crate) const WM_RSS_SHOW_CONTEXT: u32 = WM_USER + 205;
const WM_RSS_BACKGROUND_CHECK_COMPLETE: u32 = WM_USER + 206;
const ADD_GUARD_TIMER_ID: usize = 1;
const EM_REPLACESEL: u32 = 0x00C2;
const REORDER_EDIT_ID: usize = 1401;
const REORDER_OK_ID: usize = 1402;
const REORDER_CANCEL_ID: usize = 1403;

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

fn rss_source_display_title(source: &RssSource, language: crate::settings::Language) -> String {
    let base_title = if source.title.trim().is_empty() {
        source.url.clone()
    } else {
        source.title.clone()
    };
    if source.unread {
        format!("{}{}", i18n::tr(language, "rss.unread_prefix"), base_title)
    } else {
        base_title
    }
}

fn percent_encode(input: &str) -> String {
    url::form_urlencoded::byte_serialize(input.as_bytes()).collect()
}

fn parse_single_path(buffer: &[u16]) -> Option<PathBuf> {
    let end = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
    if end == 0 {
        return None;
    }
    Some(PathBuf::from(String::from_utf16_lossy(&buffer[..end])))
}

fn parse_opml_sources(text: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(text);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if !e.name().as_ref().eq_ignore_ascii_case(b"outline") {
                    buf.clear();
                    continue;
                }
                let mut url = String::new();
                let mut title = String::new();
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let value = attr
                        .decode_and_unescape_value(&reader)
                        .unwrap_or_default()
                        .to_string();
                    if key.eq_ignore_ascii_case(b"xmlUrl") {
                        url = value;
                    } else if key.eq_ignore_ascii_case(b"title")
                        || key.eq_ignore_ascii_case(b"text")
                    {
                        if title.is_empty() {
                            title = value;
                        }
                    }
                }
                if !url.trim().is_empty() {
                    out.push((title, url));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn open_import_txt_dialog(hwnd: HWND, language: crate::settings::Language) -> Option<PathBuf> {
    let filter_raw = i18n::tr(language, "rss.import_filter");
    let filter = to_wide(&filter_raw.replace("\\0", "\0"));
    let mut buffer = vec![0u16; 4096];
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        ..Default::default()
    };
    if !unsafe { GetOpenFileNameW(&mut ofn).as_bool() } {
        return None;
    }
    parse_single_path(&buffer)
}

unsafe fn import_sources_from_file(hwnd: HWND, path: &Path) {
    let bytes = match std::fs::read(path) {
        Ok(bytes) => bytes,
        Err(err) => {
            log_debug(&format!(
                "rss_import_file_error path=\"{}\" error=\"{}\"",
                path.to_string_lossy(),
                err
            ));
            return;
        }
    };
    let text = String::from_utf8_lossy(&bytes);
    let is_opml = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.eq_ignore_ascii_case("opml"))
        .unwrap_or(false)
        || text.to_ascii_lowercase().contains("<opml");
    let opml_sources = if is_opml {
        parse_opml_sources(&text)
    } else {
        Vec::new()
    };
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }
    let mut added = 0usize;
    with_state(parent, |state| {
        let mut existing: HashSet<String> = state
            .settings
            .rss_sources
            .iter()
            .map(|src| normalize_rss_url_key(&src.url))
            .filter(|k| !k.is_empty())
            .collect();
        if !opml_sources.is_empty() {
            for (mut title, url_raw) in opml_sources {
                let url = rss::normalize_url(&url_raw);
                if url.is_empty() {
                    continue;
                }
                let key = normalize_rss_url_key(&url);
                if key.is_empty() || existing.contains(&key) {
                    continue;
                }
                if title.trim().is_empty() {
                    title = url.clone();
                }
                state.settings.rss_sources.push(RssSource {
                    title: title.clone(),
                    url: url.clone(),
                    kind: RssSourceType::Feed,
                    user_title: title.trim() != url.trim(),
                    unread: false,
                    cache: rss::RssFeedCache::default(),
                    last_seen_guid: None,
                });
                existing.insert(key);
                added += 1;
            }
        } else {
            for line in text.lines() {
                let url_raw = line.trim();
                if url_raw.is_empty() {
                    continue;
                }
                let url = rss::normalize_url(url_raw);
                if url.is_empty() {
                    continue;
                }
                let key = normalize_rss_url_key(&url);
                if key.is_empty() || existing.contains(&key) {
                    continue;
                }
                state.settings.rss_sources.push(RssSource {
                    title: url.clone(),
                    url: url.clone(),
                    kind: RssSourceType::Feed,
                    user_title: false,
                    unread: false,
                    cache: rss::RssFeedCache::default(),
                    last_seen_guid: None,
                });
                existing.insert(key);
                added += 1;
            }
        }
        if added > 0 {
            crate::settings::save_settings(state.settings.clone());
        }
    });
    if added > 0 {
        log_debug(&format!(
            "rss_import_file_added path=\"{}\" count={}",
            path.to_string_lossy(),
            added
        ));
        reload_tree(hwnd);
    }
}

fn is_valid_article_url(url: &str) -> bool {
    let Ok(parsed) = Url::parse(url) else {
        return false;
    };
    matches!(parsed.scheme(), "http" | "https")
}

fn open_url_in_browser(url: &str) -> Result<(), String> {
    let url_wide = to_wide(url);
    let verb = to_wide("open");
    unsafe {
        let result = ShellExecuteW(
            HWND(0),
            PCWSTR(verb.as_ptr()),
            PCWSTR(url_wide.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOW,
        );
        if result.0 as isize <= 32 {
            return Err(format!("ShellExecute failed ({})", result.0 as isize));
        }
    }
    Ok(())
}

fn move_vec_to_index<T>(items: &mut Vec<T>, from: usize, to: usize) -> bool {
    if from >= items.len() {
        return false;
    }
    let target = to.min(items.len().saturating_sub(1));
    if from == target {
        return false;
    }
    let item = items.remove(from);
    items.insert(target, item);
    true
}

fn rss_item_key(item: &RssItem) -> String {
    if !item.guid.trim().is_empty() {
        return item.guid.trim().to_string();
    }
    if !item.link.trim().is_empty() {
        return item.link.trim().to_string();
    }
    item.title.trim().to_string()
}

unsafe extern "system" fn rss_tree_compare(
    lparam1: LPARAM,
    lparam2: LPARAM,
    _lparam_sort: LPARAM,
) -> i32 {
    let a = lparam1.0;
    let b = lparam2.0;
    a.cmp(&b) as i32
}

fn collect_root_items(hwnd_tree: HWND) -> Vec<windows::Win32::UI::Controls::HTREEITEM> {
    let mut items = Vec::new();
    let mut current = windows::Win32::UI::Controls::HTREEITEM(
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_ROOT as usize),
                LPARAM(0),
            )
        }
        .0,
    );
    while current.0 != 0 {
        items.push(current);
        current = windows::Win32::UI::Controls::HTREEITEM(
            unsafe {
                SendMessageW(
                    hwnd_tree,
                    TVM_GETNEXTITEM,
                    WPARAM(TVGN_NEXT as usize),
                    LPARAM(current.0),
                )
            }
            .0,
        );
    }
    items
}

fn apply_root_order(
    hwnd: HWND,
    hwnd_tree: HWND,
    ordered_items: &[windows::Win32::UI::Controls::HTREEITEM],
) {
    for (i, hitem) in ordered_items.iter().enumerate() {
        let mut item = TVITEMW {
            mask: TVIF_PARAM,
            lParam: LPARAM(i as isize),
            ..Default::default()
        };
        item.hItem = *hitem;
        let _ = unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_SETITEMW,
                WPARAM(0),
                LPARAM(&mut item as *mut _ as isize),
            )
        };
    }
    unsafe {
        with_rss_state(hwnd, |s| {
            for (i, hitem) in ordered_items.iter().enumerate() {
                s.node_data.insert(hitem.0, NodeData::Source(i));
            }
        });
    }
    let mut sort_cb = TVSORTCB {
        hParent: TVI_ROOT,
        lpfnCompare: Some(rss_tree_compare),
        lParam: LPARAM(0),
    };
    let _ = unsafe {
        SendMessageW(
            hwnd_tree,
            TVM_SORTCHILDRENCB,
            WPARAM(0),
            LPARAM(&mut sort_cb as *mut _ as isize),
        )
    };
}

fn announce_rss_status(message: &str) {
    log_debug(&format!("rss_status {}", message));
    let _ = nvda_speak(message);
}

fn rss_page_sizes(parent: HWND) -> (usize, usize) {
    unsafe {
        with_state(parent, |s| {
            (
                s.settings.rss_initial_page_size,
                s.settings.rss_next_page_size,
            )
        })
        .unwrap_or((INITIAL_LOAD_COUNT, LOAD_MORE_COUNT))
    }
}

fn ensure_rss_http(parent: HWND) {
    let config = unsafe {
        with_state(parent, |s| rss::config_from_settings(&s.settings))
            .unwrap_or_else(rss::RssHttpConfig::default)
    };
    if let Err(err) = rss::init_http(config) {
        log_debug(&format!("rss_http_init_error: {}", err));
    }
}

fn rss_fetch_config(parent: HWND) -> rss::RssFetchConfig {
    unsafe {
        with_state(parent, |s| rss::fetch_config_from_settings(&s.settings))
            .unwrap_or_else(rss::RssFetchConfig::default)
    }
}

fn default_feed_path(language: crate::settings::Language) -> Option<PathBuf> {
    let file_name = match language {
        crate::settings::Language::English => "feed_en.txt",
        crate::settings::Language::Italian => "feed_it.txt",
        crate::settings::Language::Spanish => "feed_es.txt",
        crate::settings::Language::Portuguese => "feed_pt.txt",
        crate::settings::Language::Vietnamese => "feed_en.txt",
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
    candidates.into_iter().find(|path| path.exists())
}

fn embedded_default_feeds(language: crate::settings::Language) -> &'static str {
    match language {
        crate::settings::Language::English => FEED_EN_DATA,
        crate::settings::Language::Italian => FEED_IT_DATA,
        crate::settings::Language::Spanish => FEED_ES_DATA,
        crate::settings::Language::Portuguese => FEED_PT_DATA,
        crate::settings::Language::Vietnamese => FEED_EN_DATA,
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
        crate::settings::Language::Vietnamese => settings
            .rss_default_en_keys
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
            unread: false,
            cache: rss::RssFeedCache::default(),
            last_seen_guid: None,
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
            crate::settings::Language::Vietnamese => apply_default_sources(
                &mut s.settings.rss_sources,
                &mut s.settings.rss_removed_default_en,
                &mut s.settings.rss_default_en_keys,
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
    hwnd_import: HWND,
    node_data: HashMap<isize, NodeData>,
    pending_fetches: HashMap<String, isize>, // URL -> hItem
    source_items: HashMap<isize, SourceItemsState>,
    enter_guard: bool,
    add_guard: bool,
    reorder_dialog: HWND,
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

pub unsafe fn focus_library(hwnd: HWND) {
    if hwnd.0 == 0 {
        return;
    }
    SetForegroundWindow(hwnd);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 != 0 {
        SetFocus(hwnd_tree);
    }
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

    let (is_source, source_index, article_item) =
        with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
            Some(NodeData::Source(idx)) => (true, Some(*idx), None),
            Some(NodeData::Item(item)) => (false, None, Some(item.clone())),
            None => (false, None, None),
        })
        .unwrap_or((false, None, None));
    if !is_source && article_item.is_none() {
        return;
    }

    let language = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
    })
    .unwrap_or_default();
    let edit_label = i18n::tr(language, "rss.context.edit");
    let delete_label = i18n::tr(language, "rss.context.delete");
    let retry_label = i18n::tr(language, "rss.context.retry_now");
    let reorder_label = i18n::tr(language, "rss.context.reorder");
    let reorder_up = i18n::tr(language, "rss.reorder.move_up");
    let reorder_down = i18n::tr(language, "rss.reorder.move_down");
    let reorder_top = i18n::tr(language, "rss.reorder.move_top");
    let reorder_bottom = i18n::tr(language, "rss.reorder.move_bottom");
    let reorder_position = i18n::tr(language, "rss.reorder.move_to_position");
    let open_label = i18n::tr(language, "rss.context.open_browser");
    let facebook_label = i18n::tr(language, "rss.context.share_facebook");
    let twitter_label = i18n::tr(language, "rss.context.share_twitter");
    let whatsapp_label = i18n::tr(language, "rss.context.share_whatsapp");

    if let Ok(menu) = CreatePopupMenu() {
        if menu.0 != 0 {
            if is_source {
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
                let _ = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_RETRY,
                    PCWSTR(to_wide(&retry_label).as_ptr()),
                );
                if let Some(idx) = source_index {
                    let total = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| ps.settings.rss_sources.len())
                    })
                    .flatten()
                    .unwrap_or(0);
                    let at_top = idx == 0;
                    let at_bottom = total == 0 || idx + 1 >= total;
                    if let Ok(submenu) = CreatePopupMenu() {
                        if submenu.0 != 0 {
                            let up_flags = if at_top {
                                MF_STRING | MF_GRAYED
                            } else {
                                MF_STRING
                            };
                            let down_flags = if at_bottom {
                                MF_STRING | MF_GRAYED
                            } else {
                                MF_STRING
                            };
                            let _ = AppendMenuW(
                                submenu,
                                up_flags,
                                ID_CTX_REORDER_UP,
                                PCWSTR(to_wide(&reorder_up).as_ptr()),
                            );
                            let _ = AppendMenuW(
                                submenu,
                                down_flags,
                                ID_CTX_REORDER_DOWN,
                                PCWSTR(to_wide(&reorder_down).as_ptr()),
                            );
                            let _ = AppendMenuW(
                                submenu,
                                up_flags,
                                ID_CTX_REORDER_TOP,
                                PCWSTR(to_wide(&reorder_top).as_ptr()),
                            );
                            let _ = AppendMenuW(
                                submenu,
                                down_flags,
                                ID_CTX_REORDER_BOTTOM,
                                PCWSTR(to_wide(&reorder_bottom).as_ptr()),
                            );
                            let _ = AppendMenuW(
                                submenu,
                                MF_STRING,
                                ID_CTX_REORDER_POSITION,
                                PCWSTR(to_wide(&reorder_position).as_ptr()),
                            );
                            let _ = AppendMenuW(
                                menu,
                                MF_POPUP,
                                submenu.0 as usize,
                                PCWSTR(to_wide(&reorder_label).as_ptr()),
                            );
                        }
                    }
                }
            } else if let Some(item) = article_item {
                let url = item.link.trim();
                let valid_url = !url.is_empty() && is_valid_article_url(url);
                let flags = if valid_url {
                    MF_STRING
                } else {
                    MF_STRING | MF_GRAYED
                };
                let _ = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_OPEN_BROWSER,
                    PCWSTR(to_wide(&open_label).as_ptr()),
                );
                let _ = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_FACEBOOK,
                    PCWSTR(to_wide(&facebook_label).as_ptr()),
                );
                let _ = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_TWITTER,
                    PCWSTR(to_wide(&twitter_label).as_ptr()),
                );
                let _ = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_WHATSAPP,
                    PCWSTR(to_wide(&whatsapp_label).as_ptr()),
                );
            }
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
                hwnd_import: HWND(0),
                node_data: HashMap::new(),
                pending_fetches: HashMap::new(),
                source_items: HashMap::new(),
                enter_guard: false,
                add_guard: false,
                reorder_dialog: HWND(0),
                pending_edit: None,
                tree_proc: None,
                last_selected: 0,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

            create_controls(hwnd);
            ensure_default_sources(parent);
            reload_tree(hwnd);

            // Start background check for new articles on all feeds
            start_background_unread_check(hwnd);

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
                ID_BTN_IMPORT => {
                    let language = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
                    })
                    .unwrap_or_default();
                    if let Some(path) = open_import_txt_dialog(hwnd, language) {
                        import_sources_from_file(hwnd, &path);
                    }
                    LRESULT(0)
                }
                1 => {
                    // IDOK (Enter key often triggers this generic command)
                    let focus = GetFocus();
                    let btn_add = GetDlgItem(hwnd, ID_BTN_ADD as i32);
                    let btn_import = GetDlgItem(hwnd, ID_BTN_IMPORT as i32);
                    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));

                    if focus == btn_add {
                        let already = with_rss_state(hwnd, |s| s.add_guard).unwrap_or(false);
                        if !already {
                            with_rss_state(hwnd, |s| s.add_guard = true);
                            let _ = PostMessageW(hwnd, WM_SHOW_ADD_DIALOG, WPARAM(0), LPARAM(0));
                        }
                        return LRESULT(0);
                    }

                    if focus == btn_import {
                        let language = with_rss_state(hwnd, |s| {
                            with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
                        })
                        .unwrap_or_default();
                        if let Some(path) = open_import_txt_dialog(hwnd, language) {
                            import_sources_from_file(hwnd, &path);
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
                ID_CTX_RETRY => {
                    handle_retry_now(hwnd);
                    LRESULT(0)
                }
                ID_CTX_REORDER_UP => {
                    handle_reorder_action(hwnd, ReorderAction::Up);
                    LRESULT(0)
                }
                ID_CTX_REORDER_DOWN => {
                    handle_reorder_action(hwnd, ReorderAction::Down);
                    LRESULT(0)
                }
                ID_CTX_REORDER_TOP => {
                    handle_reorder_action(hwnd, ReorderAction::Top);
                    LRESULT(0)
                }
                ID_CTX_REORDER_BOTTOM => {
                    handle_reorder_action(hwnd, ReorderAction::Bottom);
                    LRESULT(0)
                }
                ID_CTX_REORDER_POSITION => {
                    handle_reorder_action(hwnd, ReorderAction::Position);
                    LRESULT(0)
                }
                ID_CTX_OPEN_BROWSER => {
                    handle_article_action(hwnd, ArticleAction::OpenInBrowser);
                    LRESULT(0)
                }
                ID_CTX_SHARE_FACEBOOK => {
                    handle_article_action(hwnd, ArticleAction::ShareFacebook);
                    LRESULT(0)
                }
                ID_CTX_SHARE_TWITTER => {
                    handle_article_action(hwnd, ArticleAction::ShareTwitter);
                    LRESULT(0)
                }
                ID_CTX_SHARE_WHATSAPP => {
                    handle_article_action(hwnd, ArticleAction::ShareWhatsApp);
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
                            unread: false,
                            cache: rss::RssFeedCache::default(),
                            last_seen_guid: None,
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
                        // action contains TVE_EXPAND (2) when expanding
                        // Use bitwise AND to check for the expand flag
                        let action_val = (*pnmtv).action.0 as u32;
                        if (action_val & TVE_EXPAND.0) != 0 {
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
        WM_RSS_BACKGROUND_CHECK_COMPLETE => {
            let ptr = lparam.0 as *mut BackgroundCheckResult;
            let res = *Box::from_raw(ptr);
            process_background_check_result(hwnd, res);
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
                if parent.0 != 0 {
                    editor_manager::mark_current_document_from_rss(parent, true);
                }
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

unsafe extern "system" fn reorder_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const CREATESTRUCTW;
            let init_ptr = (*cs).lpCreateParams as *mut ReorderDialogInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, init_ptr as isize);
            let init = &*init_ptr;

            let language = with_rss_state(init.parent, |s| {
                with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
            })
            .unwrap_or_default();
            let position_template = i18n::tr(language, "rss.reorder.position_of");
            let position_text = position_template
                .replace("{x}", &(init.source_index + 1).to_string())
                .replace("{n}", &init.total.to_string());
            let move_label = i18n::tr(language, "rss.reorder.move_to_position");
            let ok_label = i18n::tr(language, "rss.dialog.ok");
            let cancel_label = i18n::tr(language, "rss.dialog.cancel");

            let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&position_text).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                10,
                330,
                16,
                hwnd,
                HMENU(1),
                hinstance,
                None,
            );
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&move_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                36,
                330,
                16,
                hwnd,
                HMENU(2),
                hinstance,
                None,
            );
            let edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                10,
                54,
                120,
                24,
                hwnd,
                HMENU(REORDER_EDIT_ID as isize),
                hinstance,
                None,
            );
            let _ = SendMessageW(edit, EM_LIMITTEXT, WPARAM(6), LPARAM(0));
            let ok = CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&ok_label).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                170,
                96,
                80,
                24,
                hwnd,
                HMENU(REORDER_OK_ID as isize),
                hinstance,
                None,
            );
            let cancel = CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&cancel_label).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                260,
                96,
                80,
                24,
                hwnd,
                HMENU(REORDER_CANCEL_ID as isize),
                hinstance,
                None,
            );
            for control in [edit, ok, cancel] {
                let prev = SetWindowLongPtrW(
                    control,
                    GWLP_WNDPROC,
                    reorder_control_subclass_proc as isize,
                );
                let _ = SetWindowLongPtrW(control, GWLP_USERDATA, prev);
            }
            SetFocus(edit);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            match id {
                1 => {
                    SendMessageW(hwnd, WM_COMMAND, WPARAM(REORDER_OK_ID), LPARAM(0));
                    LRESULT(0)
                }
                REORDER_OK_ID => {
                    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ReorderDialogInit;
                    if ptr.is_null() {
                        return LRESULT(0);
                    }
                    let init = &*ptr;
                    let edit = GetDlgItem(hwnd, REORDER_EDIT_ID as i32);
                    let len = SendMessageW(
                        edit,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXTLENGTH,
                        WPARAM(0),
                        LPARAM(0),
                    )
                    .0;
                    let mut buf = vec![0u16; len as usize + 1];
                    SendMessageW(
                        edit,
                        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXT,
                        WPARAM(buf.len()),
                        LPARAM(buf.as_mut_ptr() as isize),
                    );
                    let text = from_wide(buf.as_ptr());
                    let language = with_rss_state(init.parent, |s| {
                        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
                    })
                    .unwrap_or_default();
                    let pos = match text.trim().parse::<usize>() {
                        Ok(v) if v > 0 => v,
                        _ => {
                            let message = i18n::tr(language, "rss.reorder.invalid_position");
                            announce_rss_status(&message);
                            SetFocus(edit);
                            return LRESULT(0);
                        }
                    };
                    let target = pos.clamp(1, init.total) - 1;
                    if let Some(new_index) = apply_reorder_action(
                        init.parent,
                        init.source_index,
                        ReorderAction::Position,
                        target,
                    ) {
                        if new_index != init.source_index {
                            let template = i18n::tr(language, "rss.reorder.moved_position");
                            let message = template.replace("{x}", &(new_index + 1).to_string());
                            announce_rss_status(&message);
                        }
                    }
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                REORDER_CANCEL_ID | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u16 == VK_ESCAPE.0 {
                SendMessageW(hwnd, WM_COMMAND, WPARAM(REORDER_CANCEL_ID), LPARAM(0));
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ReorderDialogInit;
            if !ptr.is_null() {
                let init = Box::from_raw(ptr);
                with_rss_state(init.parent, |s| s.reorder_dialog = HWND(0));
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
        120,
        30,
        hwnd,
        HMENU(ID_BTN_ADD as isize),
        hinstance,
        None,
    );

    let hwnd_import = CreateWindowExW(
        Default::default(),
        WC_BUTTON,
        PCWSTR(to_wide(&i18n::tr(language, "rss.tree.import_txt")).as_ptr()),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        140,
        520,
        160,
        30,
        hwnd,
        HMENU(ID_BTN_IMPORT as isize),
        hinstance,
        None,
    );

    let hwnd_close = CreateWindowExW(
        Default::default(),
        WC_BUTTON,
        PCWSTR(to_wide(&i18n::tr(language, "rss.tree.close")).as_ptr()),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        320,
        520,
        120,
        30,
        hwnd,
        HMENU(ID_BTN_CLOSE as isize),
        hinstance,
        None,
    );

    with_rss_state(hwnd, |s| {
        s.hwnd_tree = hwnd_tree;
        s.hwnd_import = hwnd_import;
    });

    let hfont = with_rss_state(hwnd, |s| {
        with_state(s.parent, |ps| ps.hfont).unwrap_or(HFONT(0))
    })
    .unwrap_or(HFONT(0));
    if hfont.0 != 0 {
        SendMessageW(hwnd_tree, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        SendMessageW(hwnd_add, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        SendMessageW(hwnd_import, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        SendMessageW(hwnd_close, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
    }

    SetFocus(hwnd_tree);
}

unsafe fn reload_tree(hwnd: HWND) {
    let (hwnd_tree, sources, language) = match with_rss_state(hwnd, |s| {
        (
            s.hwnd_tree,
            with_state(s.parent, |ps| ps.settings.rss_sources.clone()),
            with_state(s.parent, |ps| ps.settings.language),
        )
    }) {
        Some((t, Some(src), Some(language))) => (t, src, language),
        _ => return,
    };

    SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(TVI_ROOT.0));

    with_rss_state(hwnd, |s| {
        s.node_data.clear();
        s.source_items.clear();
    });

    for (i, source) in sources.into_iter().enumerate() {
        let title = to_wide(&rss_source_display_title(&source, language));
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

unsafe fn update_source_tree_title(
    hwnd_tree: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
    title: &str,
) {
    if hwnd_tree.0 == 0 || hitem.0 == 0 {
        return;
    }
    let title_wide = to_wide(title);
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

unsafe fn set_source_unread(
    hwnd: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
    unread: bool,
) {
    // When marking as read (!unread), also update last_seen_guid from the first item
    let first_item_key: Option<String> = if !unread {
        with_rss_state(hwnd, |s| {
            s.source_items
                .get(&hitem.0)
                .and_then(|state| state.items.first().map(rss_item_key))
        })
        .flatten()
    } else {
        None
    };

    let (hwnd_tree, title_opt) = with_rss_state(hwnd, |s| {
        let hwnd_tree = s.hwnd_tree;
        let parent = s.parent;
        let source_idx = s.node_data.get(&hitem.0).and_then(|node| match node {
            NodeData::Source(idx) => Some(*idx),
            _ => None,
        });
        let Some(idx) = source_idx else {
            return (hwnd_tree, None);
        };
        let language = with_state(parent, |ps| ps.settings.language).unwrap_or_default();
        let title_opt = with_state(parent, |ps| {
            if let Some(src) = ps.settings.rss_sources.get_mut(idx) {
                let mut changed = false;
                if src.unread != unread {
                    src.unread = unread;
                    changed = true;
                }
                // Update last_seen_guid when marking as read
                if let Some(ref key) = first_item_key {
                    if src.last_seen_guid.as_ref() != Some(key) {
                        src.last_seen_guid = Some(key.clone());
                        changed = true;
                    }
                }
                if changed {
                    let title = rss_source_display_title(src, language);
                    crate::settings::save_settings(ps.settings.clone());
                    return Some(title);
                }
            }
            None
        })
        .flatten();
        (hwnd_tree, title_opt)
    })
    .unwrap_or((HWND(0), None));
    if let Some(title) = title_opt {
        update_source_tree_title(hwnd_tree, hitem, &title);
    }
}

unsafe fn handle_expand(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM) {
    // Check if items are already loaded - if so, mark as read immediately
    let has_loaded_items = with_rss_state(hwnd, |s| {
        if !matches!(s.node_data.get(&(hitem.0)), Some(NodeData::Source(_))) {
            return false;
        }
        s.source_items
            .get(&hitem.0)
            .map(|state| !state.items.is_empty())
            .unwrap_or(false)
    })
    .unwrap_or(false);

    if has_loaded_items {
        // Items already loaded, mark as read now
        set_source_unread(hwnd, hitem, false);
    }
    // If items not loaded yet, they will be fetched and mark_as_read will happen
    // in process_fetch_result after items are loaded
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

    let (url, source_kind, mut cache, _is_source) = if let Some(info) = item_info_opt.flatten() {
        info
    } else {
        return;
    };
    let empty_items = with_rss_state(hwnd, |s| {
        s.source_items
            .get(&hitem.0)
            .map(|state| state.items.is_empty())
            .unwrap_or(true)
    })
    .unwrap_or(true);
    if empty_items {
        cache.etag = None;
        cache.last_modified = None;
    }

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
            false,
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
    result: Result<rss::RssFetchOutcome, rss::FeedFetchError>,
}

/// Result of a background check for new articles (lightweight, no UI update needed)
struct BackgroundCheckResult {
    source_idx: usize,
    newest_item_key: Option<String>,
}

/// Launch background check for all feeds to detect new articles without blocking UI
unsafe fn start_background_unread_check(hwnd: HWND) {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let sources: Vec<(usize, String, rss::RssSourceType, rss::RssFeedCache)> =
        with_state(parent, |ps| {
            ps.settings
                .rss_sources
                .iter()
                .enumerate()
                .map(|(i, src)| (i, src.url.clone(), src.kind.clone(), src.cache.clone()))
                .collect()
        })
        .unwrap_or_default();

    if sources.is_empty() {
        return;
    }

    let fetch_config = rss_fetch_config(parent);
    ensure_rss_http(parent);

    let hwnd_raw = hwnd.0;
    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();

        rt.block_on(async {
            // Process feeds concurrently but not all at once (limit concurrency)
            let semaphore = std::sync::Arc::new(tokio::sync::Semaphore::new(4));
            let mut handles = Vec::new();

            for (idx, url, kind, cache) in sources {
                let sem = semaphore.clone();
                let cfg = fetch_config;
                let hwnd_val = hwnd_raw;

                let handle = tokio::spawn(async move {
                    let _permit = sem.acquire().await.ok()?;
                    let result = rss::fetch_and_parse(&url, kind, cache, cfg, false).await;
                    if let Ok(outcome) = result {
                        let newest_key = outcome.items.first().map(|item| {
                            if !item.guid.trim().is_empty() {
                                item.guid.trim().to_string()
                            } else if !item.link.trim().is_empty() {
                                item.link.trim().to_string()
                            } else {
                                item.title.trim().to_string()
                            }
                        });
                        let msg = Box::new(BackgroundCheckResult {
                            source_idx: idx,
                            newest_item_key: newest_key,
                        });
                        let _ = PostMessageW(
                            HWND(hwnd_val),
                            WM_RSS_BACKGROUND_CHECK_COMPLETE,
                            WPARAM(0),
                            LPARAM(Box::into_raw(msg) as isize),
                        );
                    }
                    Some(())
                });
                handles.push(handle);
            }

            for h in handles {
                let _ = h.await;
            }
        });
    });
}

/// Process background check result - update unread state if new articles detected
unsafe fn process_background_check_result(hwnd: HWND, res: BackgroundCheckResult) {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let Some(newest_key) = res.newest_item_key else {
        return;
    };

    // Check if this is a new article compared to last_seen_guid
    let should_mark_unread = with_state(parent, |ps| {
        ps.settings
            .rss_sources
            .get(res.source_idx)
            .map(|src| match &src.last_seen_guid {
                Some(last_seen) => last_seen != &newest_key,
                None => true, // Never seen before
            })
            .unwrap_or(false)
    })
    .unwrap_or(false);

    if should_mark_unread {
        // Find the tree item for this source and mark it unread
        let hitem_opt = with_rss_state(hwnd, |s| {
            for (&h, node) in &s.node_data {
                if let NodeData::Source(idx) = node {
                    if *idx == res.source_idx {
                        return Some(windows::Win32::UI::Controls::HTREEITEM(h));
                    }
                }
            }
            None
        })
        .flatten();

        if let Some(hitem) = hitem_opt {
            set_source_unread(hwnd, hitem, true);
        }
    }
}

unsafe fn process_fetch_result(hwnd: HWND, res: FetchResult) {
    let hitem = windows::Win32::UI::Controls::HTREEITEM(res.hitem);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));

    let caret = windows::Win32::UI::Controls::HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );

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
                    let allow_title_update = caret != hitem;
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
                            final_title = rss_source_display_title(src, lang);
                            if src.kind != outcome.kind {
                                src.kind = outcome.kind;
                            }
                            src.cache = outcome.cache.clone();
                        }
                        crate::settings::save_settings(ps.settings.clone());
                    });

                    if allow_title_update {
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
            }

            if outcome.not_modified {
                let has_items = with_rss_state(hwnd, |s| {
                    s.source_items
                        .get(&hitem.0)
                        .map(|state| !state.items.is_empty())
                        .unwrap_or(false)
                })
                .unwrap_or(false);
                if !has_items {
                    let child = SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(TVGN_CHILD as usize),
                        LPARAM(hitem.0),
                    );
                    if child.0 != 0 {
                        let mut item = TVITEMW {
                            mask: TVIF_TEXT,
                            hItem: windows::Win32::UI::Controls::HTREEITEM(child.0),
                            pszText: windows::core::PWSTR::null(),
                            cchTextMax: 0,
                            ..Default::default()
                        };
                        let _ = SendMessageW(
                            hwnd_tree,
                            TVM_GETITEMW,
                            WPARAM(0),
                            LPARAM(&mut item as *mut _ as isize),
                        );
                        let mut buf = vec![0u16; 64];
                        item.pszText = windows::core::PWSTR(buf.as_mut_ptr());
                        item.cchTextMax = buf.len() as i32;
                        if SendMessageW(
                            hwnd_tree,
                            TVM_GETITEMW,
                            WPARAM(0),
                            LPARAM(&mut item as *mut _ as isize),
                        )
                        .0 != 0
                        {
                            let text = from_wide(buf.as_ptr());
                            if text.trim() == "Loading..." {
                                let _ = SendMessageW(
                                    hwnd_tree,
                                    TVM_DELETEITEM,
                                    WPARAM(0),
                                    LPARAM(child.0),
                                );
                            }
                        }
                    }
                }
                return;
            }

            let mut appended = 0usize;
            let mut loaded_before = 0usize;
            let mut total_before = 0usize;
            let existing =
                with_rss_state(hwnd, |s| s.source_items.contains_key(&hitem.0)).unwrap_or(false);

            if existing {
                with_rss_state(hwnd, |s| {
                    let Some(state) = s.source_items.get_mut(&hitem.0) else {
                        return;
                    };
                    loaded_before = state.loaded;
                    total_before = state.items.len();
                    let mut seen: HashSet<String> = state.items.iter().map(rss_item_key).collect();
                    for item in outcome.items {
                        let key = rss_item_key(&item);
                        if seen.insert(key) {
                            state.items.push(item);
                            appended += 1;
                        }
                    }
                });
                if appended > 0 {
                    log_debug(&format!(
                        "rss_ui_batch start source={} append_count={} loaded_before={} total_before={}",
                        hitem.0, appended, loaded_before, total_before
                    ));
                    if loaded_before >= total_before {
                        let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                        let (_initial_count, next_count) = rss_page_sizes(parent);
                        let _ = load_more_items(hwnd, hitem, next_count);
                    }
                    log_debug(&format!(
                        "rss_ui_batch end source={} appended={}",
                        hitem.0, appended
                    ));
                    set_source_unread(hwnd, hitem, true);
                }
            } else {
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
                log_debug(&format!(
                    "rss_ui_batch start source={} append_count={} loaded_before=0 total_before=0",
                    hitem.0, initial_count
                ));
                let inserted = load_more_items(hwnd, hitem, initial_count);
                log_debug(&format!(
                    "rss_ui_batch end source={} appended={}",
                    hitem.0, inserted
                ));

                // User expanded the feed and items are now loaded - mark as read
                // This updates last_seen_guid to the newest item
                set_source_unread(hwnd, hitem, false);
            }
        }
        Err(e) => {
            let (message, cache) = match e {
                rss::FeedFetchError::InCooldown { kind, cache, .. } => (
                    format!("Feed temporarily blocked/cooldown ({kind}). Try again later."),
                    Some(cache),
                ),
                rss::FeedFetchError::HttpStatus {
                    status,
                    kind,
                    cache,
                } => (format!("Feed error {status} ({kind})."), Some(cache)),
                rss::FeedFetchError::Network { message, cache } => {
                    (format!("Error: {message}"), Some(cache))
                }
            };

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
                if let (Some(i), Some(cache)) = (idx, cache) {
                    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    with_state(parent, |ps| {
                        if let Some(src) = ps.settings.rss_sources.get_mut(i) {
                            src.cache = cache;
                        }
                        crate::settings::save_settings(ps.settings.clone());
                    });
                }
            }

            let has_items =
                with_rss_state(hwnd, |s| s.source_items.contains_key(&hitem.0)).unwrap_or(false);
            if has_items {
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
            let text = to_wide(&message);
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
    let (inserted, loaded_after, total_after) = with_rss_state(hwnd, |s| {
        let Some(state) = s.source_items.get_mut(&hitem.0) else {
            return (0usize, 0usize, 0usize);
        };
        if state.loaded >= state.items.len() {
            return (0usize, state.loaded, state.items.len());
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
        (inserted, state.loaded, state.items.len())
    })
    .unwrap_or((0usize, 0usize, 0usize));
    let _ = SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
    if inserted > 0 {
        log_debug(&format!(
            "rss_ui_batch append source={} inserted={} loaded={} total={}",
            hitem.0, inserted, loaded_after, total_after
        ));
    }
    inserted
}

unsafe fn handle_enter(hwnd: HWND) {
    // UI: On Enter, fetch article content and import into the editor.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
        }
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
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
        }
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
                        | crate::settings::Language::Vietnamese
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
                                crate::settings::Language::Vietnamese => {
                                    &mut ps.settings.rss_removed_default_en
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
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
        }
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

unsafe fn handle_retry_now(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
        }
        .0,
    );
    if hitem.0 == 0 {
        return;
    }

    let source_info = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => with_state(s.parent, |ps| {
            ps.settings
                .rss_sources
                .get(*idx)
                .map(|src| (src.url.clone(), src.kind.clone(), src.cache.clone()))
        })
        .flatten(),
        _ => None,
    })
    .flatten();

    let Some((url, source_kind, cache)) = source_info else {
        return;
    };
    if url.trim().is_empty() {
        return;
    }

    let host = Url::parse(&url)
        .ok()
        .and_then(|u| u.host_str().map(|h| h.to_string()))
        .unwrap_or_default();
    log_debug(&format!(
        "rss_action kind=feed action=retry_now override=true host=\"{}\"",
        host
    ));

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let fetch_config = if parent.0 != 0 {
        rss_fetch_config(parent)
    } else {
        rss::RssFetchConfig::default()
    };
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    let url_clone = url.clone();
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
            true,
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

#[derive(Clone, Copy)]
enum ReorderAction {
    Up,
    Down,
    Top,
    Bottom,
    Position,
}

#[derive(Clone, Copy)]
struct ReorderDialogInit {
    parent: HWND,
    source_index: usize,
    total: usize,
}

fn selected_source_index(hwnd: HWND) -> Option<usize> {
    let hwnd_tree = unsafe { with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0)) };
    if hwnd_tree.0 == 0 {
        return None;
    }
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(0),
            )
        }
        .0,
    );
    if hitem.0 == 0 {
        return None;
    }
    unsafe {
        with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
            Some(NodeData::Source(idx)) => Some(*idx),
            _ => None,
        })
        .flatten()
    }
}

fn apply_reorder_action(
    hwnd: HWND,
    source_index: usize,
    action: ReorderAction,
    target_index: usize,
) -> Option<usize> {
    let hwnd_tree = unsafe { with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0)) };
    let parent = unsafe { with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0)) };
    if hwnd_tree.0 == 0 || parent.0 == 0 {
        return None;
    }
    let mut root_items = collect_root_items(hwnd_tree);
    if source_index >= root_items.len() {
        return None;
    }
    let new_index = unsafe {
        with_state(parent, |ps| {
            let moved = match action {
                ReorderAction::Up => {
                    crate::settings::move_rss_feed_up(&mut ps.settings, source_index)
                }
                ReorderAction::Down => {
                    crate::settings::move_rss_feed_down(&mut ps.settings, source_index)
                }
                ReorderAction::Top => {
                    crate::settings::move_rss_feed_to_top(&mut ps.settings, source_index)
                }
                ReorderAction::Bottom => {
                    crate::settings::move_rss_feed_to_bottom(&mut ps.settings, source_index)
                }
                ReorderAction::Position => crate::settings::move_rss_feed_to_index(
                    &mut ps.settings,
                    source_index,
                    target_index,
                ),
            };
            if moved.is_some() {
                crate::settings::save_settings(ps.settings.clone());
            }
            moved
        })
        .flatten()
    };
    let new_index = new_index?;
    if move_vec_to_index(&mut root_items, source_index, new_index) {
        apply_root_order(hwnd, hwnd_tree, &root_items);
    }
    Some(new_index)
}

unsafe fn handle_reorder_action(hwnd: HWND, action: ReorderAction) {
    let Some(source_index) = selected_source_index(hwnd) else {
        return;
    };
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = with_state(parent, |ps| ps.settings.language).unwrap_or_default();
    let total = with_state(parent, |ps| ps.settings.rss_sources.len()).unwrap_or(0);
    if total == 0 {
        return;
    }
    if matches!(action, ReorderAction::Position) {
        show_reorder_dialog(hwnd, source_index, total);
        return;
    }
    let new_index = match action {
        ReorderAction::Up => apply_reorder_action(hwnd, source_index, action, 0),
        ReorderAction::Down => apply_reorder_action(hwnd, source_index, action, 0),
        ReorderAction::Top => apply_reorder_action(hwnd, source_index, action, 0),
        ReorderAction::Bottom => apply_reorder_action(hwnd, source_index, action, 0),
        ReorderAction::Position => None,
    };
    if let Some(new_index) = new_index {
        if new_index != source_index {
            let template = i18n::tr(language, "rss.reorder.moved_position");
            let message = template.replace("{x}", &(new_index + 1).to_string());
            announce_rss_status(&message);
        }
    }
}

#[derive(Clone, Copy)]
enum ArticleAction {
    OpenInBrowser,
    ShareFacebook,
    ShareTwitter,
    ShareWhatsApp,
}

unsafe fn selected_article_item(hwnd: HWND) -> Option<RssItem> {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return None;
    }
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
        return None;
    }
    with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Item(item)) => Some(item.clone()),
        _ => None,
    })
    .flatten()
}

unsafe fn handle_article_action(hwnd: HWND, action: ArticleAction) {
    let Some(item) = selected_article_item(hwnd) else {
        log_debug("rss_action kind=article action=unavailable reason=no_article");
        return;
    };
    let url = item.link.trim().to_string();
    if !is_valid_article_url(&url) {
        log_debug("rss_action kind=article action=unavailable reason=invalid_url");
        return;
    }
    let title = item.title.trim().to_string();
    let share_url = match action {
        ArticleAction::OpenInBrowser => {
            log_debug(&format!(
                "rss_action kind=article action=open_in_browser url=\"{}\"",
                url
            ));
            url.clone()
        }
        ArticleAction::ShareFacebook => {
            log_debug(&format!(
                "rss_action kind=article action=share_facebook url=\"{}\"",
                url
            ));
            format!(
                "https://www.facebook.com/sharer/sharer.php?u={}",
                percent_encode(&url)
            )
        }
        ArticleAction::ShareTwitter => {
            log_debug(&format!(
                "rss_action kind=article action=share_twitter url=\"{}\"",
                url
            ));
            let mut share = format!(
                "https://twitter.com/intent/tweet?url={}",
                percent_encode(&url)
            );
            if !title.is_empty() {
                share.push_str("&text=");
                share.push_str(&percent_encode(&title));
            }
            share
        }
        ArticleAction::ShareWhatsApp => {
            log_debug(&format!(
                "rss_action kind=article action=share_whatsapp url=\"{}\"",
                url
            ));
            let text = if title.is_empty() {
                url.clone()
            } else {
                format!("{}\n{}", title, url)
            };
            format!("https://wa.me/?text={}", percent_encode(&text))
        }
    };
    if let Err(err) = open_url_in_browser(&share_url) {
        log_debug(&format!(
            "rss_action kind=article action=browser_error error=\"{}\"",
            err
        ));
    }
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
        NotifyWinEvent(
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

        let text = match content_res {
            Ok(text) => text,
            Err(err) => {
                log_debug(&format!(
                    "rss_import_fallback url=\"{}\" error=\"{}\"",
                    url, err
                ));
                format!("{}\n\n{}", item.title, url)
            }
        };
        let msg = Box::new(ImportResult { text });
        let _ = PostMessageW(
            hwnd,
            WM_RSS_IMPORT_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        );
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
            out.push('\n');
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

unsafe fn show_reorder_dialog(parent_hwnd: HWND, source_index: usize, total: usize) {
    let existing = with_rss_state(parent_hwnd, |s| s.reorder_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide("NovapadRssReorder");
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
        lpfnWndProc: Some(reorder_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_rss_state(parent_hwnd, |s| {
        with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
    })
    .unwrap_or_default();
    let title = i18n::tr(language, "rss.context.reorder");
    let init_ptr = Box::into_raw(Box::new(ReorderDialogInit {
        parent: parent_hwnd,
        source_index,
        total,
    }));
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide(&title).as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_POPUP,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        360,
        160,
        parent_hwnd,
        None,
        hinstance,
        Some(init_ptr as *const _),
    );
    if hwnd.0 == 0 {
        let _ = Box::from_raw(init_ptr);
        return;
    }
    with_rss_state(parent_hwnd, |s| s.reorder_dialog = hwnd);
}

unsafe extern "system" fn reorder_control_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_CHAR {
        if wparam.0 as u16 == VK_TAB.0 {
            return LRESULT(0);
        }
    }
    if msg == WM_KEYDOWN {
        let id = GetDlgCtrlID(hwnd) as usize;
        let parent = GetParent(hwnd);
        let edit = GetDlgItem(parent, REORDER_EDIT_ID as i32);
        let ok = GetDlgItem(parent, REORDER_OK_ID as i32);
        let cancel = GetDlgItem(parent, REORDER_CANCEL_ID as i32);
        if wparam.0 as u16 == VK_TAB.0 {
            let shift = (GetKeyState(VK_SHIFT.0 as i32) & 0x8000u16 as i16) != 0;
            let next = if shift {
                if id == REORDER_EDIT_ID {
                    cancel
                } else if id == REORDER_CANCEL_ID {
                    ok
                } else {
                    edit
                }
            } else if id == REORDER_EDIT_ID {
                ok
            } else if id == REORDER_OK_ID {
                cancel
            } else {
                edit
            };
            SetFocus(next);
            return LRESULT(0);
        }
        if wparam.0 as u16 == VK_RETURN.0 {
            let target = if id == REORDER_CANCEL_ID {
                REORDER_CANCEL_ID
            } else {
                REORDER_OK_ID
            };
            SendMessageW(parent, WM_COMMAND, WPARAM(target), LPARAM(0));
            return LRESULT(0);
        }
        if wparam.0 as u16 == VK_ESCAPE.0 {
            SendMessageW(parent, WM_COMMAND, WPARAM(REORDER_CANCEL_ID), LPARAM(0));
            return LRESULT(0);
        }
    }
    let prev = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if prev == 0 {
        return DefWindowProcW(hwnd, msg, wparam, lparam);
    }
    CallWindowProcW(
        Some(std::mem::transmute::<
            isize,
            unsafe extern "system" fn(HWND, u32, WPARAM, LPARAM) -> LRESULT,
        >(prev)),
        hwnd,
        msg,
        wparam,
        lparam,
    )
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
