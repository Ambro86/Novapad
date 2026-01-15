use crate::accessibility::{from_wide, handle_accessibility, nvda_speak, to_wide};
use crate::editor_manager;
use crate::i18n;
use crate::settings::{self, Language, confirm_title};
use crate::tools::rss::{self, PodcastEpisode, RssSource, RssSourceType};
use crate::{log_debug, with_state};
use quick_xml::{Reader, events::Event};
use sha1::{Digest as Sha1Digest, Sha1};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::DataExchange::{
    COPYDATASTRUCT, CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Accessibility::NotifyWinEvent;
use windows::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, GetSaveFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY,
    OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows::Win32::UI::Controls::{
    HTREEITEM, NM_RETURN, NMTVKEYDOWN, TVGN_CARET, TVGN_CHILD, TVGN_NEXT, TVGN_PARENT, TVGN_ROOT,
    TVIF_CHILDREN, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVITEMEXW_CHILDREN, TVITEMW,
    TVM_DELETEITEM, TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_INSERTITEMW,
    TVM_SELECTITEM, TVM_SETITEMW, TVM_SORTCHILDRENCB, TVN_ITEMEXPANDINGW, TVN_KEYDOWN,
    TVN_SELCHANGEDW, TVSORTCB,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetFocus, VK_APPS, VK_DELETE, VK_ESCAPE, VK_F10, VK_LEFT,
    VK_RETURN, VK_RIGHT, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CB_ADDSTRING, CB_GETCURSEL, CB_SETCURSEL, CBS_DROPDOWNLIST, CHILDID_SELF,
    CallWindowProcW, CreateMenu, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu,
    DestroyWindow, EVENT_OBJECT_FOCUS, GetClientRect, GetDlgCtrlID, GetDlgItem, GetParent,
    GetWindowLongPtrW, GetWindowRect, HMENU, IDC_ARROW, IDYES, LB_ADDSTRING, LB_GETCURSEL,
    LB_RESETCONTENT, LB_SETCURSEL, LBN_DBLCLK, LBS_NOTIFY, MB_ICONINFORMATION, MB_OK, MB_YESNO,
    MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, MSG, MessageBoxW, OBJID_CLIENT, PostMessageW,
    RegisterClassW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW,
    TrackPopupMenu, WINDOW_STYLE, WM_CHAR, WM_COMMAND, WM_CONTEXTMENU, WM_COPYDATA, WM_CREATE,
    WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL, WM_NOTIFY, WM_SETFOCUS, WM_SETFONT,
    WM_SIZE, WNDCLASSW, WNDPROC, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_POPUP, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
};
use windows::core::{PCWSTR, PWSTR, w};

const PODCASTS_WINDOW_CLASS: &str = "NovapadPodcasts";
const PODCASTS_REORDER_CLASS: &str = "NovapadPodcastsReorder";
const PODCASTS_ADD_CLASS: &str = "NovapadPodcastsAdd";

const ID_TREE: usize = 12001;
const ID_SEARCH_LABEL: usize = 12005;
const ID_SEARCH_EDIT: usize = 12002;
const ID_SEARCH_PROVIDER: usize = 12011;
const ID_SEARCH_BUTTON: usize = 12006;
const ID_RESULTS: usize = 12003;
const ID_ADD_BUTTON: usize = 12004;
const ID_IMPORT_BUTTON: usize = 12009;
const ID_EXPORT_BUTTON: usize = 12010;
const ID_DELETE_BUTTON: usize = 12008;
const ID_CLOSE_BUTTON: usize = 12007;

const REORDER_EDIT_ID: usize = 12101;
const REORDER_OK_ID: usize = 12102;
const REORDER_CANCEL_ID: usize = 12103;

const ADD_URL_EDIT_ID: usize = 12201;
const ADD_OK_ID: usize = 12202;
const ADD_CANCEL_ID: usize = 12203;

const WM_PODCAST_FETCH_COMPLETE: u32 = windows::Win32::UI::WindowsAndMessaging::WM_USER + 310;
const WM_PODCAST_SEARCH_COMPLETE: u32 = windows::Win32::UI::WindowsAndMessaging::WM_USER + 311;
const WM_PODCAST_PLAY_READY: u32 = windows::Win32::UI::WindowsAndMessaging::WM_USER + 312;
const WM_PODCAST_PLAY_FAILED: u32 = windows::Win32::UI::WindowsAndMessaging::WM_USER + 313;

const EM_SETSEL: u32 = 0x00B1;
const EM_SCROLLCARET: u32 = 0x00B7;

const ID_CTX_UPDATE: usize = 13001;
const ID_CTX_REMOVE: usize = 13002;
const ID_CTX_COPY_URL: usize = 13003;
const ID_CTX_OPEN_FEED: usize = 13004;
const ID_CTX_REORDER_UP: usize = 13005;
const ID_CTX_REORDER_DOWN: usize = 13006;
const ID_CTX_REORDER_TOP: usize = 13007;
const ID_CTX_REORDER_BOTTOM: usize = 13008;
const ID_CTX_REORDER_POSITION: usize = 13009;

const ID_CTX_PLAY: usize = 13101;
const ID_CTX_OPEN_EPISODE: usize = 13102;
const ID_CTX_COPY_AUDIO: usize = 13103;
const ID_CTX_COPY_TITLE: usize = 13104;
const ID_CTX_DOWNLOAD_EPISODE: usize = 13105;

const ID_CTX_SUBSCRIBE: usize = 13201;
const ID_CTX_SEARCH_INFO: usize = 13202;
const ID_CTX_SEARCH_COPY_URL: usize = 13203;
const PODCAST_ADD_COPYDATA: usize = 0x504F4443;

#[derive(Clone)]
struct PodcastSearchResult {
    title: String,
    artist: String,
    feed_url: String,
}

#[derive(Clone, Copy)]
enum SearchProvider {
    Itunes,
    PodcastIndex,
}

struct PodcastWindowState {
    parent: HWND,
    language: Language,
    hwnd_tree: HWND,
    hwnd_search_label: HWND,
    hwnd_search: HWND,
    hwnd_search_provider: HWND,
    hwnd_search_button: HWND,
    hwnd_results: HWND,
    hwnd_add: HWND,
    hwnd_import: HWND,
    hwnd_export: HWND,
    hwnd_delete: HWND,
    hwnd_close: HWND,
    node_data: HashMap<isize, NodeData>,
    source_items: HashMap<isize, SourceItemsState>,
    pending_fetches: HashMap<String, isize>,
    search_results: Vec<PodcastSearchResult>,
    tree_proc: WNDPROC,
    search_proc: WNDPROC,
    reorder_dialog: HWND,
    last_selected: isize,
    pending_play: Option<String>,
}

#[derive(Clone)]
enum NodeData {
    Source(usize),
    Episode(PodcastEpisode),
}

struct SourceItemsState {
    items: Vec<PodcastEpisode>,
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

fn parse_single_path(buffer: &[u16]) -> Option<PathBuf> {
    let end = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
    if end == 0 {
        return None;
    }
    Some(PathBuf::from(String::from_utf16_lossy(&buffer[..end])))
}

fn open_opml_file_dialog(hwnd: HWND, language: Language, for_import: bool) -> Option<PathBuf> {
    let raw_filter = i18n::tr(language, "rss.import_filter");
    let filter = to_wide(&raw_filter.replace("\\0", "\0"));
    let mut buffer = vec![0u16; 4096];
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER
            | OFN_HIDEREADONLY
            | OFN_PATHMUSTEXIST
            | if for_import {
                OFN_FILEMUSTEXIST
            } else {
                OFN_OVERWRITEPROMPT
            },
        ..Default::default()
    };
    let success = if for_import {
        unsafe { GetOpenFileNameW(&mut ofn).as_bool() }
    } else {
        unsafe { GetSaveFileNameW(&mut ofn).as_bool() }
    };
    if !success {
        return None;
    }
    parse_single_path(&buffer)
}

fn normalize_podcast_key(url: &str) -> String {
    rss::normalize_url(url).to_ascii_lowercase()
}

fn import_podcast_sources_from_file(hwnd: HWND, path: &Path) -> Option<usize> {
    let bytes = std::fs::read(path).ok()?;
    let text = String::from_utf8_lossy(&bytes);
    let is_opml = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.eq_ignore_ascii_case("opml"))
        .unwrap_or(false)
        || text.to_ascii_lowercase().contains("<opml");
    let sources = if is_opml {
        parse_opml_sources(&text)
    } else {
        Vec::new()
    };
    let parent = unsafe { with_podcast_state(hwnd, |s| s.parent) }.unwrap_or(HWND(0));
    if parent.0 == 0 {
        return None;
    }
    let mut added = 0usize;
    unsafe {
        with_state(parent, |state| {
            let mut existing: HashSet<String> = state
                .settings
                .podcast_sources
                .iter()
                .map(|src| normalize_podcast_key(&src.url))
                .filter(|k| !k.is_empty())
                .collect();
            for (mut title, url_raw) in sources {
                let url = rss::normalize_url(&url_raw);
                if url.is_empty() {
                    continue;
                }
                let key = normalize_podcast_key(&url);
                if key.is_empty() || existing.contains(&key) {
                    continue;
                }
                if title.trim().is_empty() {
                    title = url.clone();
                }
                state.settings.podcast_sources.push(rss::RssSource {
                    title: title.clone(),
                    url: url.clone(),
                    kind: rss::RssSourceType::Feed,
                    user_title: title.trim() != url.trim(),
                    unread: false,
                    cache: rss::RssFeedCache::default(),
                });
                existing.insert(key);
                added += 1;
            }
            if added > 0 {
                crate::settings::save_settings(state.settings.clone());
            }
        });
    }
    if added > 0 {
        unsafe {
            reload_tree(hwnd);
        }
    }
    Some(added)
}

fn escape_opml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn export_podcast_sources_to_file(hwnd: HWND, path: &Path) -> Result<usize, String> {
    let parent = unsafe { with_podcast_state(hwnd, |s| s.parent) }.unwrap_or(HWND(0));
    if parent.0 == 0 {
        return Err("missing parent".to_string());
    }
    let sources = unsafe { with_state(parent, |state| state.settings.podcast_sources.clone()) }
        .unwrap_or_default();
    if sources.is_empty() {
        return Ok(0);
    }
    let mut file = File::create(path).map_err(|e| e.to_string())?;
    writeln!(
        file,
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<opml version=\"1.0\">\n<head>\n<title>Novapad Podcasts</title>\n</head>\n<body>"
    )
    .map_err(|e| e.to_string())?;
    for src in sources.iter() {
        let title = if src.title.trim().is_empty() {
            src.url.clone()
        } else {
            src.title.clone()
        };
        writeln!(
            file,
            "  <outline text=\"{}\" title=\"{}\" xmlUrl=\"{}\" />",
            escape_opml_attr(&title),
            escape_opml_attr(&title),
            escape_opml_attr(&src.url)
        )
        .map_err(|e| e.to_string())?;
    }
    writeln!(file, "</body>\n</opml>").map_err(|e| e.to_string())?;
    Ok(sources.len())
}

unsafe fn handle_import_opml(hwnd: HWND) {
    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    if let Some(path) = open_opml_file_dialog(hwnd, language, true) {
        if let Some(count) = import_podcast_sources_from_file(hwnd, &path) {
            if count > 0 {
                announce_status(&i18n::tr(language, "podcasts.imported"));
            }
        } else {
            let title = i18n::tr(language, "podcasts.window.title");
            let message = i18n::tr(language, "podcasts.import_failed");
            let _ = MessageBoxW(
                hwnd,
                PCWSTR(to_wide(&message).as_ptr()),
                PCWSTR(to_wide(&title).as_ptr()),
                MB_OK | MB_ICONINFORMATION,
            );
        }
    }
}

unsafe fn handle_export_opml(hwnd: HWND) {
    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    if let Some(path) = open_opml_file_dialog(hwnd, language, false) {
        match export_podcast_sources_to_file(hwnd, &path) {
            Ok(count) => {
                if count > 0 {
                    announce_status(&i18n::tr(language, "podcasts.exported"));
                }
            }
            Err(err) => {
                let title = i18n::tr(language, "podcasts.window.title");
                let message = format!("{}: {}", i18n::tr(language, "podcasts.export_failed"), err);
                let _ = MessageBoxW(
                    hwnd,
                    PCWSTR(to_wide(&message).as_ptr()),
                    PCWSTR(to_wide(&title).as_ptr()),
                    MB_OK | MB_ICONINFORMATION,
                );
            }
        }
    }
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

struct AddDialogInit {
    parent: HWND,
}

struct FetchResult {
    hitem: isize,
    source_index: usize,
    result: Result<rss::PodcastFetchOutcome, rss::FeedFetchError>,
}

struct SearchResultMsg {
    results: Vec<PodcastSearchResult>,
}

struct PlayReadyMsg {
    path: PathBuf,
    enclosure_url: String,
    title: String,
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_CHAR {
        return false;
    }
    if msg.message == WM_KEYDOWN {
        let key = msg.wParam.0 as u32;
        if key == VK_ESCAPE.0 as u32 {
            let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM(2), LPARAM(0));
            return true;
        }
        if key == VK_RETURN.0 as u32 {
            let (hwnd_tree, hwnd_results) =
                with_podcast_state(hwnd, |s| (s.hwnd_tree, s.hwnd_results))
                    .unwrap_or((HWND(0), HWND(0)));
            let focus = GetFocus();

            // Handle Enter on search results list
            if hwnd_results.0 != 0 && focus == hwnd_results {
                subscribe_selected_result(hwnd);
                return true;
            }

            // Handle Enter on tree view
            if hwnd_tree.0 != 0 && focus == hwnd_tree {
                if let Some(item) = selected_episode(hwnd) {
                    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    open_episode_in_player(hwnd, parent, &item);
                    return true;
                }
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

unsafe fn announce_status(message: &str) {
    log_debug(&format!("podcasts_status {}", message));
    let _ = nvda_speak(message);
}

unsafe fn ensure_rss_http(parent: HWND) {
    let config = with_state(parent, |s| rss::config_from_settings(&s.settings))
        .unwrap_or_else(|| rss::RssHttpConfig::default());
    if let Err(err) = rss::init_http(config) {
        log_debug(&format!("rss_http_init_error: {}", err));
    }
}

unsafe fn rss_fetch_config(parent: HWND) -> rss::RssFetchConfig {
    with_state(parent, |s| rss::fetch_config_from_settings(&s.settings))
        .unwrap_or_else(|| rss::RssFetchConfig::default())
}

unsafe fn open_url_in_browser(url: &str) -> Result<(), String> {
    let wide = to_wide(url);
    let result = ShellExecuteW(
        HWND(0),
        w!("open"),
        PCWSTR(wide.as_ptr()),
        PCWSTR::null(),
        PCWSTR::null(),
        windows::Win32::UI::WindowsAndMessaging::SW_SHOW,
    );
    if result.0 as isize <= 32 {
        return Err(format!("ShellExecute failed: {}", result.0));
    }
    Ok(())
}

unsafe fn copy_text_to_clipboard(hwnd: HWND, text: &str) {
    use windows::Win32::Foundation::HANDLE;
    use windows::Win32::System::Memory::{GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalUnlock};

    const CF_UNICODETEXT: u32 = 13;

    let content = to_wide(text);
    if content.is_empty() {
        return;
    }
    if OpenClipboard(hwnd).is_err() {
        return;
    }
    let _ = EmptyClipboard();
    let size = content.len() * std::mem::size_of::<u16>();
    let handle = match GlobalAlloc(GMEM_MOVEABLE, size) {
        Ok(handle) => handle,
        Err(_) => {
            let _ = CloseClipboard();
            return;
        }
    };
    if handle.0.is_null() {
        let _ = CloseClipboard();
        return;
    }
    let ptr = GlobalLock(handle) as *mut u16;
    if ptr.is_null() {
        let _ = CloseClipboard();
        return;
    }
    std::ptr::copy_nonoverlapping(content.as_ptr(), ptr, content.len());
    let _ = GlobalUnlock(handle);
    let _ = SetClipboardData(CF_UNICODETEXT, HANDLE(handle.0 as isize));
    let _ = CloseClipboard();
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

unsafe extern "system" fn podcast_tree_compare(
    lparam1: LPARAM,
    lparam2: LPARAM,
    _lparam_sort: LPARAM,
) -> i32 {
    let a = lparam1.0;
    let b = lparam2.0;
    a.cmp(&b) as i32
}

unsafe fn collect_root_items(hwnd_tree: HWND) -> Vec<HTREEITEM> {
    let mut items = Vec::new();
    let mut current = HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_ROOT as usize),
            LPARAM(0),
        )
        .0,
    );
    while current.0 != 0 {
        items.push(current);
        current = HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_GETNEXTITEM,
                WPARAM(TVGN_NEXT as usize),
                LPARAM(current.0),
            )
            .0,
        );
    }
    items
}

unsafe fn apply_root_order(hwnd: HWND, hwnd_tree: HWND, ordered_items: &[HTREEITEM]) {
    for (i, hitem) in ordered_items.iter().enumerate() {
        let mut item = TVITEMW {
            mask: TVIF_PARAM,
            lParam: LPARAM(i as isize),
            ..Default::default()
        };
        item.hItem = *hitem;
        let _ = SendMessageW(
            hwnd_tree,
            TVM_SETITEMW,
            WPARAM(0),
            LPARAM(&mut item as *mut _ as isize),
        );
    }
    with_podcast_state(hwnd, |s| {
        for (i, hitem) in ordered_items.iter().enumerate() {
            s.node_data.insert(hitem.0, NodeData::Source(i));
        }
    });
    let mut sort_cb = TVSORTCB {
        hParent: windows::Win32::UI::Controls::TVI_ROOT,
        lpfnCompare: Some(podcast_tree_compare),
        lParam: LPARAM(0),
    };
    let _ = SendMessageW(
        hwnd_tree,
        TVM_SORTCHILDRENCB,
        WPARAM(0),
        LPARAM(&mut sort_cb as *mut _ as isize),
    );
}

unsafe fn selected_tree_item(hwnd: HWND) -> HTREEITEM {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return HTREEITEM(0);
    }
    HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    )
}

unsafe fn selected_source_index(hwnd: HWND) -> Option<usize> {
    let hitem = selected_tree_item(hwnd);
    if hitem.0 == 0 {
        return None;
    }
    with_podcast_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten()
}

unsafe fn selected_source_name(hwnd: HWND) -> Option<String> {
    let index = selected_source_index(hwnd)?;
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return None;
    }
    with_state(parent, |ps| {
        ps.settings.podcast_sources.get(index).map(|src| {
            if src.title.trim().is_empty() {
                src.url.clone()
            } else {
                src.title.clone()
            }
        })
    })
    .unwrap_or(None)
}

fn update_delete_button_state(hwnd: HWND) {
    let enabled = unsafe { selected_source_name(hwnd).is_some() };
    unsafe {
        with_podcast_state(hwnd, |state| {
            if state.hwnd_delete.0 != 0 {
                EnableWindow(state.hwnd_delete, enabled);
            }
        });
    }
}

unsafe fn selected_episode(hwnd: HWND) -> Option<PodcastEpisode> {
    let hitem = selected_tree_item(hwnd);
    if hitem.0 == 0 {
        return None;
    }
    with_podcast_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Episode(item)) => Some(item.clone()),
        _ => None,
    })
    .flatten()
}

fn episode_key(item: &PodcastEpisode) -> String {
    if !item.guid.trim().is_empty() {
        return item.guid.trim().to_string();
    }
    if !item.link.trim().is_empty() {
        return item.link.trim().to_string();
    }
    item.title.trim().to_string()
}

unsafe fn load_episode_children(hwnd: HWND, hitem: HTREEITEM, source_index: usize, force: bool) {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let (parent, url, mut cache, force_uncached) = with_podcast_state(hwnd, |s| {
        let parent = s.parent;
        let empty_items = s
            .source_items
            .get(&hitem.0)
            .map(|state| state.items.is_empty())
            .unwrap_or(true);
        let (url, cache) = with_state(parent, |ps| {
            ps.settings
                .podcast_sources
                .get(source_index)
                .map(|src| (src.url.clone(), src.cache.clone()))
        })
        .unwrap_or(None)
        .unwrap_or((String::new(), rss::RssFeedCache::default()));
        (parent, url, cache, empty_items)
    })
    .unwrap_or((HWND(0), String::new(), rss::RssFeedCache::default(), true));
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }
    if url.trim().is_empty() {
        return;
    }
    if force_uncached {
        cache.etag = None;
        cache.last_modified = None;
    }

    let should_fetch = with_podcast_state(hwnd, |s| {
        if s.pending_fetches.contains_key(&url) {
            return false;
        }
        let state = s.source_items.get(&hitem.0);
        if state.is_none() {
            return true;
        }
        if force {
            return true;
        }
        state.map(|s| s.items.is_empty()).unwrap_or(true)
    })
    .unwrap_or(true);

    if !should_fetch {
        return;
    }

    with_podcast_state(hwnd, |s| {
        s.pending_fetches.insert(url.clone(), hitem.0);
    });

    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    let loading_txt = to_wide(&i18n::tr(language, "podcasts.loading"));

    let child = HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CHILD as usize),
            LPARAM(hitem.0),
        )
        .0,
    );
    if child.0 == 0 {
        let mut tvis_loading = TVINSERTSTRUCTW {
            hParent: hitem,
            hInsertAfter: windows::Win32::UI::Controls::TVI_LAST,
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

    let _ = SendMessageW(
        hwnd_tree,
        TVM_EXPAND,
        WPARAM(windows::Win32::UI::Controls::TVE_EXPAND.0 as usize),
        LPARAM(hitem.0),
    );
    let _ = SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));

    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let res = rt.block_on(rss::fetch_podcast_feed(
            &url,
            cache,
            rss_fetch_config(parent),
            false,
        ));
        let msg = Box::new(FetchResult {
            hitem: hitem.0,
            source_index,
            result: res,
        });
        let _ = PostMessageW(
            hwnd_copy,
            WM_PODCAST_FETCH_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        );
    });
}

unsafe fn update_source_cache(parent: HWND, source_index: usize, cache: rss::RssFeedCache) {
    let _ = with_state(parent, |ps| {
        if let Some(src) = ps.settings.podcast_sources.get_mut(source_index) {
            src.cache = cache;
            settings::save_settings(ps.settings.clone());
        }
    });
}

unsafe fn update_source_title(hwnd: HWND, hitem: HTREEITEM, source_index: usize, feed_title: &str) {
    let title = feed_title.trim();
    if title.is_empty() {
        return;
    }
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }
    let mut updated = None;
    let _ = with_state(parent, |ps| {
        if let Some(src) = ps.settings.podcast_sources.get_mut(source_index) {
            let looks_auto = src.title.trim().is_empty() || src.title == src.url;
            if !src.user_title && looks_auto {
                src.title = title.to_string();
                updated = Some(src.title.clone());
                settings::save_settings(ps.settings.clone());
            }
        }
    });
    let Some(updated) = updated else {
        return;
    };
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let title_wide = to_wide(&updated);
    let mut tvi = TVITEMW {
        mask: TVIF_TEXT,
        hItem: hitem,
        pszText: windows::core::PWSTR(title_wide.as_ptr() as *mut _),
        ..Default::default()
    };
    let _ = SendMessageW(
        hwnd_tree,
        TVM_SETITEMW,
        WPARAM(0),
        LPARAM(&mut tvi as *mut _ as isize),
    );
}

unsafe fn apply_episode_results(hwnd: HWND, hitem: HTREEITEM, items: Vec<PodcastEpisode>) {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }

    let existing_keys: HashSet<String> = with_podcast_state(hwnd, |s| {
        s.source_items
            .get(&hitem.0)
            .map(|state| state.items.iter().map(episode_key).collect())
            .unwrap_or_default()
    })
    .unwrap_or_default();

    let mut new_items = Vec::new();
    for item in items {
        if !existing_keys.contains(&episode_key(&item)) {
            new_items.push(item);
        }
    }

    let child = HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CHILD as usize),
            LPARAM(hitem.0),
        )
        .0,
    );
    if child.0 != 0 {
        let mut item = TVITEMW {
            mask: TVIF_TEXT,
            hItem: child,
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
        let mut buf = vec![0u16; 128];
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
            if text.trim()
                == i18n::tr(
                    with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                    "podcasts.loading",
                )
            {
                let _ = SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(child.0));
            }
        }
    }

    for item in new_items.iter() {
        let title = to_wide(&item.title);
        let mut tvis = TVINSERTSTRUCTW {
            hParent: hitem,
            hInsertAfter: windows::Win32::UI::Controls::TVI_LAST,
            Anonymous: windows::Win32::UI::Controls::TVINSERTSTRUCTW_0 {
                item: TVITEMW {
                    mask: TVIF_TEXT,
                    pszText: windows::core::PWSTR(title.as_ptr() as *mut _),
                    cchTextMax: title.len() as i32,
                    ..Default::default()
                },
            },
        };
        let inserted = HTREEITEM(
            SendMessageW(
                hwnd_tree,
                TVM_INSERTITEMW,
                WPARAM(0),
                LPARAM(&mut tvis as *mut _ as isize),
            )
            .0,
        );
        if inserted.0 != 0 {
            with_podcast_state(hwnd, |s| {
                s.node_data
                    .insert(inserted.0, NodeData::Episode(item.clone()));
            });
        }
    }

    with_podcast_state(hwnd, |s| {
        let state = s
            .source_items
            .entry(hitem.0)
            .or_insert(SourceItemsState { items: Vec::new() });
        state.items.extend(new_items);
    });
}

unsafe fn create_tree_item(hwnd_tree: HWND, title: &str, index: usize) -> HTREEITEM {
    let title_w = to_wide(title);
    let mut tvis = TVINSERTSTRUCTW {
        hParent: HTREEITEM(0),
        hInsertAfter: windows::Win32::UI::Controls::TVI_LAST,
        Anonymous: windows::Win32::UI::Controls::TVINSERTSTRUCTW_0 {
            item: TVITEMW {
                mask: TVIF_TEXT | TVIF_PARAM | TVIF_CHILDREN,
                pszText: windows::core::PWSTR(title_w.as_ptr() as *mut _),
                cchTextMax: title_w.len() as i32,
                cChildren: TVITEMEXW_CHILDREN(1),
                lParam: LPARAM(index as isize),
                ..Default::default()
            },
        },
    };
    HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_INSERTITEMW,
            WPARAM(0),
            LPARAM(&mut tvis as *mut _ as isize),
        )
        .0,
    )
}

unsafe fn reload_tree(hwnd: HWND) {
    let (hwnd_tree, sources) = with_podcast_state(hwnd, |s| {
        let sources =
            with_state(s.parent, |ps| ps.settings.podcast_sources.clone()).unwrap_or_default();
        (s.hwnd_tree, sources)
    })
    .unwrap_or((HWND(0), Vec::new()));
    if hwnd_tree.0 == 0 {
        return;
    }
    let _ = SendMessageW(
        hwnd_tree,
        TVM_DELETEITEM,
        WPARAM(0),
        LPARAM(windows::Win32::UI::Controls::TVI_ROOT.0),
    );
    with_podcast_state(hwnd, |s| {
        s.node_data.clear();
        s.source_items.clear();
    });

    for (i, src) in sources.iter().enumerate() {
        let title = if src.title.trim().is_empty() {
            src.url.clone()
        } else {
            src.title.clone()
        };
        let hitem = create_tree_item(hwnd_tree, &title, i);
        if hitem.0 != 0 {
            with_podcast_state(hwnd, |s| {
                s.node_data.insert(hitem.0, NodeData::Source(i));
            });
        }
    }

    let first = HTREEITEM(
        SendMessageW(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_ROOT as usize),
            LPARAM(0),
        )
        .0,
    );
    if first.0 != 0 {
        let _ = SendMessageW(
            hwnd_tree,
            TVM_SELECTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(first.0),
        );
    }
}

unsafe fn open_episode_in_player(hwnd: HWND, parent: HWND, episode: &PodcastEpisode) {
    let Some(url) = episode.enclosure_url.as_ref() else {
        let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
        announce_status(&i18n::tr(language, "podcasts.no_audio_url"));
        return;
    };
    let play_key = episode_key(episode);
    let should_start = with_podcast_state(hwnd, |s| {
        if s.pending_play.as_deref() == Some(play_key.as_str()) {
            return false;
        }
        s.pending_play = Some(play_key.clone());
        true
    })
    .unwrap_or(true);
    if !should_start {
        return;
    }
    if parent.0 != 0 {
        crate::set_pending_podcast_chapters_key(parent, Some(play_key.clone()));
        if !episode.podlove_chapters.is_empty() {
            crate::cache_podcast_chapters(
                parent,
                play_key.clone(),
                episode.podlove_chapters.clone(),
            );
        } else if let Some(chapters_url) = episode.chapters_url.clone() {
            let chapters_type = episode.chapters_type.clone();
            let should_fetch = match chapters_type
                .as_deref()
                .map(|t| t.trim().to_ascii_lowercase())
            {
                None => true,
                Some(kind) => kind == "application/json" || kind == "application/json+chapters",
            };
            if should_fetch {
                crate::prefetch_podcast_chapters(parent, play_key.clone(), chapters_url);
            }
        }
    }

    // Show loading message immediately so user knows action was triggered
    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    announce_status(&i18n::tr(language, "podcasts.loading"));

    if parent.0 != 0 {
        ensure_rss_http(parent);
    }
    let url = url.clone();
    let episode_title = episode.title.clone();
    let enclosure_type = episode.enclosure_type.clone();
    let parent_hwnd = parent;
    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let fetch_config = rss_fetch_config(parent_hwnd);
        let bytes = rt.block_on(rss::fetch_url_bytes(&url, fetch_config));
        let bytes = match bytes {
            Ok(b) => b,
            Err(err) => {
                log_debug(&format!("podcasts_download_error {}", err));
                let _ = PostMessageW(hwnd_copy, WM_PODCAST_PLAY_FAILED, WPARAM(0), LPARAM(0));
                return;
            }
        };
        let file_path = podcast_cache_path(&url, enclosure_type.as_deref());
        if let Some(parent_dir) = file_path.parent() {
            let _ = std::fs::create_dir_all(parent_dir);
        }
        if std::fs::write(&file_path, bytes).is_ok() {
            let msg = Box::new(PlayReadyMsg {
                path: file_path,
                enclosure_url: url.clone(),
                title: episode_title.clone(),
            });
            let _ = PostMessageW(
                hwnd_copy,
                WM_PODCAST_PLAY_READY,
                WPARAM(0),
                LPARAM(Box::into_raw(msg) as isize),
            );
        } else {
            let _ = PostMessageW(hwnd_copy, WM_PODCAST_PLAY_FAILED, WPARAM(0), LPARAM(0));
        }
    });
}

fn podcast_cache_path(url: &str, mime: Option<&str>) -> PathBuf {
    let mut hasher = sha2::Sha256::new();
    hasher.update(url.as_bytes());
    let hash = hex::encode(hasher.finalize());
    let url_ext = url
        .split('?')
        .next()
        .and_then(|s| Path::new(s).extension().and_then(|e| e.to_str()))
        .unwrap_or("");
    let ext = match mime.map(|m| m.to_ascii_lowercase()) {
        Some(mime) if mime.contains("mpeg") || mime.contains("mp3") => "mp3",
        _ => {
            if url_ext.is_empty() {
                "mp3"
            } else {
                url_ext
            }
        }
    };
    let filename = format!("podcast_{}.{}", &hash[..16], ext);
    settings::settings_dir().join("podcasts").join(filename)
}

unsafe fn add_podcast_source(parent: HWND, feed_url: &str, title: &str) -> Option<usize> {
    let normalized = rss::normalize_url(feed_url);
    if normalized.is_empty() {
        return None;
    }
    with_state(parent, |ps| {
        if ps
            .settings
            .podcast_sources
            .iter()
            .any(|src| rss::normalize_url(&src.url) == normalized)
        {
            return None;
        }
        let final_title = if title.trim().is_empty() {
            normalized.clone()
        } else {
            title.trim().to_string()
        };
        ps.settings.podcast_sources.push(RssSource {
            title: final_title,
            url: normalized,
            kind: RssSourceType::Feed,
            user_title: !title.trim().is_empty(),
            unread: false,
            cache: rss::RssFeedCache::default(),
        });
        settings::save_settings(ps.settings.clone());
        Some(ps.settings.podcast_sources.len() - 1)
    })
    .flatten()
}

unsafe fn update_search_results(hwnd: HWND, results: Vec<PodcastSearchResult>) {
    let hwnd_results = with_podcast_state(hwnd, |s| s.hwnd_results).unwrap_or(HWND(0));
    if hwnd_results.0 == 0 {
        return;
    }
    let _ = SendMessageW(hwnd_results, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if results.is_empty() {
        let text = to_wide(&i18n::tr(
            with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
            "podcasts.search.no_results",
        ));
        let _ = SendMessageW(
            hwnd_results,
            LB_ADDSTRING,
            WPARAM(0),
            LPARAM(text.as_ptr() as isize),
        );
        with_podcast_state(hwnd, |s| s.search_results = Vec::new());
        let _ = SendMessageW(hwnd_results, LB_SETCURSEL, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_results);
        return;
    }
    for item in &results {
        let label = if item.artist.trim().is_empty() {
            item.title.clone()
        } else {
            format!("{} - {}", item.title, item.artist)
        };
        let wide = to_wide(&label);
        let _ = SendMessageW(
            hwnd_results,
            LB_ADDSTRING,
            WPARAM(0),
            LPARAM(wide.as_ptr() as isize),
        );
    }
    with_podcast_state(hwnd, |s| s.search_results = results);
    let _ = SendMessageW(hwnd_results, LB_SETCURSEL, WPARAM(0), LPARAM(0));
    SetFocus(hwnd_results);
}

unsafe fn selected_search_provider(hwnd: HWND) -> SearchProvider {
    let combo = with_podcast_state(hwnd, |s| s.hwnd_search_provider).unwrap_or(HWND(0));
    if combo.0 == 0 {
        return SearchProvider::Itunes;
    }
    let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if sel == 1 {
        SearchProvider::PodcastIndex
    } else {
        SearchProvider::Itunes
    }
}

unsafe fn perform_search(hwnd: HWND, query: &str) {
    let trimmed = query.trim();
    if trimmed.is_empty() {
        return;
    }
    let hwnd_results = with_podcast_state(hwnd, |s| s.hwnd_results).unwrap_or(HWND(0));
    if hwnd_results.0 != 0 {
        let _ = SendMessageW(hwnd_results, LB_RESETCONTENT, WPARAM(0), LPARAM(0));
        let text = to_wide(&i18n::tr(
            with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
            "podcasts.loading",
        ));
        let _ = SendMessageW(
            hwnd_results,
            LB_ADDSTRING,
            WPARAM(0),
            LPARAM(text.as_ptr() as isize),
        );
        let _ = SendMessageW(hwnd_results, LB_SETCURSEL, WPARAM(0), LPARAM(0));
        SetFocus(hwnd_results);
    }
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }
    let provider = selected_search_provider(hwnd);
    let query = percent_encode(trimmed);
    let hwnd_copy = hwnd;
    std::thread::spawn(move || {
        let rt = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let mut results = Vec::new();
        match provider {
            SearchProvider::Itunes => {
                let url = format!(
                    "https://itunes.apple.com/search?media=podcast&term={}&limit=20",
                    query
                );
                let fetch_config = rss_fetch_config(parent);
                let bytes = rt.block_on(rss::fetch_url_bytes(&url, fetch_config));
                if let Ok(bytes) = bytes {
                    if let Ok(parsed) = serde_json::from_slice::<ItunesSearchResponse>(&bytes) {
                        for item in parsed.results {
                            if let Some(feed_url) = item.feed_url {
                                results.push(PodcastSearchResult {
                                    title: item.collection_name.unwrap_or_default(),
                                    artist: item.artist_name.unwrap_or_default(),
                                    feed_url,
                                });
                            }
                        }
                    }
                }
            }
            SearchProvider::PodcastIndex => {
                let key = std::env::var("PODCASTINDEX_API_KEY").ok();
                let secret = std::env::var("PODCASTINDEX_API_SECRET").ok();
                if let (Some(key), Some(secret)) = (key, secret) {
                    let auth_date = SystemTime::now()
                        .duration_since(UNIX_EPOCH)
                        .unwrap_or_default()
                        .as_secs()
                        .to_string();
                    let mut hasher = Sha1::new();
                    hasher.update(format!("{key}{secret}{auth_date}").as_bytes());
                    let hash = format!("{:x}", hasher.finalize());
                    let url = format!(
                        "https://api.podcastindex.org/api/1.0/search/byterm?q={}&max=20",
                        query
                    );
                    if let Ok(resp) = reqwest::blocking::Client::new()
                        .get(url)
                        .header("User-Agent", "Novapad")
                        .header("X-Auth-Date", auth_date)
                        .header("X-Auth-Key", key)
                        .header("Authorization", hash)
                        .send()
                    {
                        if let Ok(bytes) = resp.bytes() {
                            if let Ok(parsed) =
                                serde_json::from_slice::<PodcastIndexSearchResponse>(&bytes)
                            {
                                if let Some(feeds) = parsed.feeds {
                                    for feed in feeds {
                                        let feed_url = feed.feed_url.or(feed.url);
                                        if let Some(feed_url) = feed_url {
                                            results.push(PodcastSearchResult {
                                                title: feed.title.unwrap_or_default(),
                                                artist: feed
                                                    .author
                                                    .or(feed.owner_name)
                                                    .unwrap_or_default(),
                                                feed_url,
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                } else {
                    log_debug("PodcastIndex search skipped: missing API keys");
                }
            }
        }
        let msg = Box::new(SearchResultMsg { results });
        let _ = PostMessageW(
            hwnd_copy,
            WM_PODCAST_SEARCH_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        );
    });
}

unsafe fn subscribe_selected_result(hwnd: HWND) {
    let (parent, results, idx) = with_podcast_state(hwnd, |s| {
        let idx = SendMessageW(s.hwnd_results, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
        let results = s.search_results.clone();
        (s.parent, results, idx)
    })
    .unwrap_or((HWND(0), Vec::new(), -1));
    if idx < 0 || idx as usize >= results.len() {
        return;
    }
    let result = &results[idx as usize];
    let new_index = add_podcast_source(parent, &result.feed_url, &result.title);
    if let Some(index) = new_index {
        let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
        announce_status(&i18n::tr(language, "podcasts.added"));

        // Show confirmation dialog
        let title = i18n::tr(language, "podcasts.subscribed_title");
        let message = i18n::tr(language, "podcasts.subscribed_message");
        let _ = MessageBoxW(
            hwnd,
            PCWSTR(to_wide(&message).as_ptr()),
            PCWSTR(to_wide(&title).as_ptr()),
            MB_OK | MB_ICONINFORMATION,
        );

        let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
        if hwnd_tree.0 != 0 {
            let title = if result.title.trim().is_empty() {
                result.feed_url.clone()
            } else {
                result.title.clone()
            };
            let hitem = create_tree_item(hwnd_tree, &title, index);
            if hitem.0 != 0 {
                with_podcast_state(hwnd, |s| {
                    s.node_data.insert(hitem.0, NodeData::Source(index));
                });
                let _ = SendMessageW(
                    hwnd_tree,
                    TVM_SELECTITEM,
                    WPARAM(TVGN_CARET as usize),
                    LPARAM(hitem.0),
                );
                load_episode_children(hwnd, hitem, index, false);
            }
        }
    }
}

unsafe fn show_add_dialog(parent_hwnd: HWND) {
    let main_hwnd = with_podcast_state(parent_hwnd, |s| s.parent).unwrap_or(HWND(0));
    let existing = with_state(main_hwnd, |s| s.podcasts_add_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PODCASTS_ADD_CLASS);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            windows::Win32::UI::WindowsAndMessaging::LoadCursorW(None, IDC_ARROW)
                .unwrap_or_default()
                .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(add_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_podcast_state(parent_hwnd, |s| s.language).unwrap_or_default();
    let init_ptr = Box::into_raw(Box::new(AddDialogInit {
        parent: parent_hwnd,
    }));
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide(&i18n::tr(language, "podcasts.add_dialog.title")).as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_POPUP,
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        360,
        140,
        parent_hwnd,
        None,
        hinstance,
        Some(init_ptr as *const _),
    );
    if hwnd.0 == 0 {
        let _ = Box::from_raw(init_ptr);
        return;
    }
    with_state(main_hwnd, |s| s.podcasts_add_dialog = hwnd);
}

unsafe extern "system" fn add_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let init_ptr = (*cs).lpCreateParams as *mut AddDialogInit;
            let parent = if init_ptr.is_null() {
                HWND(0)
            } else {
                let init = Box::from_raw(init_ptr);
                init.parent
            };
            let language = with_podcast_state(parent, |s| s.language).unwrap_or_default();
            let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
            CreateWindowExW(
                Default::default(),
                w!("STATIC"),
                PCWSTR(to_wide(&i18n::tr(language, "podcasts.add_dialog.url_label")).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                10,
                10,
                320,
                16,
                hwnd,
                HMENU(1),
                hinstance,
                None,
            );
            CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("EDIT"),
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                10,
                28,
                320,
                24,
                hwnd,
                HMENU(ADD_URL_EDIT_ID as isize),
                hinstance,
                None,
            );
            CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&i18n::tr(language, "podcasts.add_dialog.ok")).as_ptr()),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WINDOW_STYLE(
                        windows::Win32::UI::WindowsAndMessaging::BS_DEFPUSHBUTTON as u32,
                    ),
                150,
                70,
                80,
                24,
                hwnd,
                HMENU(ADD_OK_ID as isize),
                hinstance,
                None,
            );
            CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&i18n::tr(language, "podcasts.add_dialog.cancel")).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                250,
                70,
                80,
                24,
                hwnd,
                HMENU(ADD_CANCEL_ID as isize),
                hinstance,
                None,
            );
            SetFocus(GetDlgItem(hwnd, ADD_URL_EDIT_ID as i32));
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            match id {
                1 => {
                    SendMessageW(hwnd, WM_COMMAND, WPARAM(ADD_OK_ID), LPARAM(0));
                    LRESULT(0)
                }
                ADD_OK_ID => {
                    let h_edit_url = GetDlgItem(hwnd, ADD_URL_EDIT_ID as i32);
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
                    let parent = GetParent(hwnd);
                    if !url.trim().is_empty() {
                        let payload = url.trim().to_string();
                        let url_wide = to_wide(&payload);
                        let cds = COPYDATASTRUCT {
                            dwData: PODCAST_ADD_COPYDATA,
                            cbData: (url_wide.len() * 2) as u32,
                            lpData: url_wide.as_ptr() as *mut _,
                        };
                        SendMessageW(
                            parent,
                            WM_COPYDATA,
                            WPARAM(hwnd.0 as usize),
                            LPARAM(&cds as *const _ as isize),
                        );
                    }
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                ADD_CANCEL_ID | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_DESTROY => {
            let parent = GetParent(hwnd);
            let main_hwnd = with_podcast_state(parent, |s| s.parent).unwrap_or(HWND(0));
            with_state(main_hwnd, |s| s.podcasts_add_dialog = HWND(0));
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

#[derive(serde::Deserialize)]
struct ItunesSearchResponse {
    #[serde(default)]
    results: Vec<ItunesSearchItem>,
}

#[derive(serde::Deserialize)]
struct ItunesSearchItem {
    #[serde(rename = "collectionName")]
    collection_name: Option<String>,
    #[serde(rename = "artistName")]
    artist_name: Option<String>,
    #[serde(rename = "feedUrl")]
    feed_url: Option<String>,
}

#[derive(serde::Deserialize)]
struct PodcastIndexSearchResponse {
    feeds: Option<Vec<PodcastIndexFeed>>,
}

#[derive(serde::Deserialize)]
struct PodcastIndexFeed {
    title: Option<String>,
    author: Option<String>,
    #[serde(rename = "ownerName")]
    owner_name: Option<String>,
    url: Option<String>,
    #[serde(rename = "feedUrl")]
    feed_url: Option<String>,
}

pub unsafe fn show_context_menu_from_keyboard(hwnd: HWND) {
    let mut pt = windows::Win32::Foundation::POINT::default();
    let _ = windows::Win32::UI::WindowsAndMessaging::GetCursorPos(&mut pt);
    show_context_menu(hwnd, pt.x, pt.y, false);
}

pub unsafe fn focus_library(hwnd: HWND) {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 != 0 {
        SetFocus(hwnd_tree);
    }
}

unsafe fn force_focus_editor_on_parent(parent: HWND) {
    if parent.0 == 0 {
        return;
    }
    SetForegroundWindow(parent);
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

unsafe fn show_context_menu(hwnd: HWND, x: i32, y: i32, use_hit_test: bool) {
    let (hwnd_tree, hwnd_results) =
        with_podcast_state(hwnd, |s| (s.hwnd_tree, s.hwnd_results)).unwrap_or((HWND(0), HWND(0)));
    if hwnd_tree.0 == 0 {
        return;
    }
    let focus = GetFocus();
    let target_list = focus == hwnd_results;
    if target_list {
        show_search_context_menu(hwnd, x, y, use_hit_test);
    } else {
        show_tree_context_menu(hwnd, x, y, use_hit_test);
    }
}

unsafe fn selected_search_result(hwnd: HWND) -> Option<PodcastSearchResult> {
    let (results, idx) = with_podcast_state(hwnd, |s| {
        let idx = SendMessageW(s.hwnd_results, LB_GETCURSEL, WPARAM(0), LPARAM(0)).0 as i32;
        (s.search_results.clone(), idx)
    })
    .unwrap_or((Vec::new(), -1));
    if idx < 0 || idx as usize >= results.len() {
        return None;
    }
    Some(results[idx as usize].clone())
}

unsafe fn trigger_search_from_edit(hwnd: HWND) {
    let (hwnd_search, hwnd_results) =
        with_podcast_state(hwnd, |s| (s.hwnd_search, s.hwnd_results)).unwrap_or((HWND(0), HWND(0)));
    if hwnd_search.0 == 0 {
        return;
    }
    let len = SendMessageW(
        hwnd_search,
        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXTLENGTH,
        WPARAM(0),
        LPARAM(0),
    )
    .0;
    let mut buf = vec![0u16; len as usize + 1];
    SendMessageW(
        hwnd_search,
        windows::Win32::UI::WindowsAndMessaging::WM_GETTEXT,
        WPARAM(buf.len()),
        LPARAM(buf.as_mut_ptr() as isize),
    );
    let query = from_wide(buf.as_ptr());
    perform_search(hwnd, &query);
    if hwnd_results.0 != 0 {
        SetFocus(hwnd_results);
    }
}

unsafe fn show_search_result_info(hwnd: HWND) {
    let result = match selected_search_result(hwnd) {
        Some(result) => result,
        None => return,
    };
    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    let title = i18n::tr(language, "podcasts.search.info_title");
    let body = i18n::tr_f(
        language,
        "podcasts.search.info_body",
        &[
            ("title", &result.title),
            ("artist", &result.artist),
            ("feed", &result.feed_url),
        ],
    );
    let _ = MessageBoxW(
        hwnd,
        PCWSTR(to_wide(&body).as_ptr()),
        PCWSTR(to_wide(&title).as_ptr()),
        windows::Win32::UI::WindowsAndMessaging::MB_OK,
    );
}

unsafe fn show_search_context_menu(hwnd: HWND, x: i32, y: i32, use_hit_test: bool) {
    let hwnd_results = with_podcast_state(hwnd, |s| s.hwnd_results).unwrap_or(HWND(0));
    if hwnd_results.0 == 0 {
        return;
    }
    let mut rect = windows::Win32::Foundation::RECT::default();
    if use_hit_test {
        if GetWindowRect(hwnd_results, &mut rect).is_ok() {
            if x < rect.left || x > rect.right || y < rect.top || y > rect.bottom {
                return;
            }
        }
    }
    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    let label = i18n::tr(language, "podcasts.context.subscribe");
    let info_label = i18n::tr(language, "podcasts.context.info");
    let copy_label = i18n::tr(language, "podcasts.context.copy_url");
    let menu = CreateMenu().unwrap_or(HMENU(0));
    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_CTX_SUBSCRIBE,
        PCWSTR(to_wide(&label).as_ptr()),
    );
    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_CTX_SEARCH_INFO,
        PCWSTR(to_wide(&info_label).as_ptr()),
    );
    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_CTX_SEARCH_COPY_URL,
        PCWSTR(to_wide(&copy_label).as_ptr()),
    );
    let cmd = TrackPopupMenu(
        menu,
        windows::Win32::UI::WindowsAndMessaging::TPM_RETURNCMD,
        x,
        y,
        0,
        hwnd,
        None,
    )
    .0 as usize;
    match cmd {
        ID_CTX_SUBSCRIBE => subscribe_selected_result(hwnd),
        ID_CTX_SEARCH_INFO => show_search_result_info(hwnd),
        ID_CTX_SEARCH_COPY_URL => {
            if let Some(result) = selected_search_result(hwnd) {
                copy_text_to_clipboard(hwnd, &result.feed_url);
            }
        }
        _ => {}
    }
    let _ = DestroyMenu(menu);
}

unsafe fn show_tree_context_menu(hwnd: HWND, x: i32, y: i32, use_hit_test: bool) {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let mut rect = windows::Win32::Foundation::RECT::default();
    if use_hit_test {
        if GetWindowRect(hwnd_tree, &mut rect).is_ok() {
            if x < rect.left || x > rect.right || y < rect.top || y > rect.bottom {
                return;
            }
        }
    }
    let hitem = selected_tree_item(hwnd);
    if hitem.0 == 0 {
        return;
    }
    let node = with_podcast_state(hwnd, |s| s.node_data.get(&hitem.0).cloned()).flatten();
    let language = with_podcast_state(hwnd, |s| s.language).unwrap_or_default();
    let menu = CreatePopupMenu().unwrap_or(HMENU(0));
    if menu.0 == 0 {
        return;
    }
    match node {
        Some(NodeData::Source(idx)) => {
            let update_label = i18n::tr(language, "podcasts.context.update");
            let remove_label = i18n::tr(language, "podcasts.context.remove");
            let reorder_label = i18n::tr(language, "podcasts.context.reorder");
            let reorder_up = i18n::tr(language, "rss.reorder.move_up");
            let reorder_down = i18n::tr(language, "rss.reorder.move_down");
            let reorder_top = i18n::tr(language, "rss.reorder.move_top");
            let reorder_bottom = i18n::tr(language, "rss.reorder.move_bottom");
            let reorder_position = i18n::tr(language, "rss.reorder.move_to_position");
            let copy_url = i18n::tr(language, "podcasts.context.copy_url");
            let open_feed = i18n::tr(language, "podcasts.context.open_feed");
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_UPDATE,
                PCWSTR(to_wide(&update_label).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_REMOVE,
                PCWSTR(to_wide(&remove_label).as_ptr()),
            );
            let total = with_podcast_state(hwnd, |s| {
                with_state(s.parent, |ps| ps.settings.podcast_sources.len()).unwrap_or(0)
            })
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
            let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_COPY_URL,
                PCWSTR(to_wide(&copy_url).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_OPEN_FEED,
                PCWSTR(to_wide(&open_feed).as_ptr()),
            );
        }
        Some(NodeData::Episode(_)) => {
            let play_label = i18n::tr(language, "podcasts.context.play");
            let open_label = i18n::tr(language, "podcasts.context.open_episode");
            let copy_audio = i18n::tr(language, "podcasts.context.copy_audio");
            let copy_title = i18n::tr(language, "podcasts.context.copy_title");
            let download_label = i18n::tr(language, "podcasts.context.download_episode");
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_PLAY,
                PCWSTR(to_wide(&play_label).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_OPEN_EPISODE,
                PCWSTR(to_wide(&open_label).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_COPY_AUDIO,
                PCWSTR(to_wide(&copy_audio).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_COPY_TITLE,
                PCWSTR(to_wide(&copy_title).as_ptr()),
            );
            let _ = AppendMenuW(
                menu,
                MF_STRING,
                ID_CTX_DOWNLOAD_EPISODE,
                PCWSTR(to_wide(&download_label).as_ptr()),
            );
        }
        None => {}
    }

    SetForegroundWindow(hwnd);
    let cmd = TrackPopupMenu(
        menu,
        windows::Win32::UI::WindowsAndMessaging::TPM_RETURNCMD,
        x,
        y,
        0,
        hwnd,
        None,
    )
    .0 as usize;
    let _ = PostMessageW(
        hwnd,
        windows::Win32::UI::WindowsAndMessaging::WM_NULL,
        WPARAM(0),
        LPARAM(0),
    );
    let _ = DestroyMenu(menu);
    match cmd {
        ID_CTX_UPDATE => handle_source_action(hwnd, SourceAction::Update),
        ID_CTX_REMOVE => handle_source_action(hwnd, SourceAction::Remove),
        ID_CTX_COPY_URL => handle_source_action(hwnd, SourceAction::CopyUrl),
        ID_CTX_OPEN_FEED => handle_source_action(hwnd, SourceAction::OpenFeed),
        ID_CTX_REORDER_UP => handle_reorder_action(hwnd, ReorderAction::Up),
        ID_CTX_REORDER_DOWN => handle_reorder_action(hwnd, ReorderAction::Down),
        ID_CTX_REORDER_TOP => handle_reorder_action(hwnd, ReorderAction::Top),
        ID_CTX_REORDER_BOTTOM => handle_reorder_action(hwnd, ReorderAction::Bottom),
        ID_CTX_REORDER_POSITION => handle_reorder_action(hwnd, ReorderAction::Position),
        ID_CTX_PLAY => handle_episode_action(hwnd, EpisodeAction::Play),
        ID_CTX_OPEN_EPISODE => handle_episode_action(hwnd, EpisodeAction::OpenEpisode),
        ID_CTX_COPY_AUDIO => handle_episode_action(hwnd, EpisodeAction::CopyAudio),
        ID_CTX_COPY_TITLE => handle_episode_action(hwnd, EpisodeAction::CopyTitle),
        ID_CTX_DOWNLOAD_EPISODE => handle_episode_action(hwnd, EpisodeAction::Download),
        ID_CTX_SUBSCRIBE => subscribe_selected_result(hwnd),
        _ => {}
    }
}

#[derive(Clone, Copy)]
enum SourceAction {
    Update,
    Remove,
    CopyUrl,
    OpenFeed,
}

unsafe fn handle_source_action(hwnd: HWND, verb: SourceAction) {
    let Some(source_index) = selected_source_index(hwnd) else {
        return;
    };
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    match verb {
        SourceAction::Update => {
            let hitem = selected_tree_item(hwnd);
            if hitem.0 != 0 {
                load_episode_children(hwnd, hitem, source_index, true);
                if parent.0 != 0 {
                    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
                    announce_status(&i18n::tr(language, "podcasts.updated"));
                }
            }
        }
        SourceAction::Remove => {
            let confirm = if parent.0 != 0 {
                let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
                let title = confirm_title(language);
                let msg = i18n::tr(language, "podcasts.remove_confirm");
                MessageBoxW(
                    hwnd,
                    PCWSTR(to_wide(&msg).as_ptr()),
                    PCWSTR(to_wide(&title).as_ptr()),
                    MB_YESNO,
                ) == IDYES
            } else {
                true
            };
            if !confirm {
                let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                if hwnd_tree.0 != 0 {
                    SetFocus(hwnd_tree);
                }
                update_delete_button_state(hwnd);
                return;
            }
            let removed = with_state(parent, |ps| {
                if source_index < ps.settings.podcast_sources.len() {
                    ps.settings.podcast_sources.remove(source_index);
                    settings::save_settings(ps.settings.clone());
                    true
                } else {
                    false
                }
            })
            .unwrap_or(false);
            if removed {
                let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
                announce_status(&i18n::tr(language, "podcasts.removed"));
                reload_tree(hwnd);
                update_delete_button_state(hwnd);
                let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                if hwnd_tree.0 != 0 {
                    SetFocus(hwnd_tree);
                    let first = HTREEITEM(
                        SendMessageW(
                            hwnd_tree,
                            TVM_GETNEXTITEM,
                            WPARAM(TVGN_ROOT as usize),
                            LPARAM(0),
                        )
                        .0,
                    );
                    if first.0 != 0 {
                        let _ = SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(first.0),
                        );
                    }
                }
            }
            {
                let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                if hwnd_tree.0 != 0 {
                    SetFocus(hwnd_tree);
                }
                update_delete_button_state(hwnd);
            }
        }
        SourceAction::CopyUrl => {
            let url = with_state(parent, |ps| {
                ps.settings
                    .podcast_sources
                    .get(source_index)
                    .map(|s| s.url.clone())
            })
            .unwrap_or(None)
            .unwrap_or_default();
            if !url.is_empty() {
                copy_text_to_clipboard(hwnd, &url);
            }
        }
        SourceAction::OpenFeed => {
            let url = with_state(parent, |ps| {
                ps.settings
                    .podcast_sources
                    .get(source_index)
                    .map(|s| s.url.clone())
            })
            .unwrap_or(None)
            .unwrap_or_default();
            if !url.is_empty() {
                let _ = open_url_in_browser(&url);
            }
        }
    }
}

#[derive(Clone, Copy)]
enum EpisodeAction {
    Play,
    OpenEpisode,
    CopyAudio,
    CopyTitle,
    Download,
}

unsafe fn handle_episode_action(hwnd: HWND, action: EpisodeAction) {
    let Some(item) = selected_episode(hwnd) else {
        return;
    };
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    match action {
        EpisodeAction::Play => open_episode_in_player(hwnd, parent, &item),
        EpisodeAction::OpenEpisode => {
            if !item.link.trim().is_empty() {
                let _ = open_url_in_browser(&item.link);
            }
        }
        EpisodeAction::CopyAudio => {
            if let Some(url) = item.enclosure_url {
                copy_text_to_clipboard(hwnd, &url);
            }
        }
        EpisodeAction::CopyTitle => copy_text_to_clipboard(hwnd, &item.title),
        EpisodeAction::Download => {
            if parent.0 == 0 {
                return;
            }
            crate::download_podcast_episode(
                parent,
                item.enclosure_url.clone(),
                Some(item.title.clone()),
                None,
                with_state(parent, |s| s.settings.language).unwrap_or_default(),
            );
        }
    }
}

unsafe fn apply_reorder_action(
    hwnd: HWND,
    source_index: usize,
    action: ReorderAction,
    target_index: usize,
) -> Option<usize> {
    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 || parent.0 == 0 {
        return None;
    }
    let mut root_items = collect_root_items(hwnd_tree);
    if source_index >= root_items.len() {
        return None;
    }
    let new_index = with_state(parent, |ps| {
        let moved = match action {
            ReorderAction::Up => settings::move_podcast_feed_up(&mut ps.settings, source_index),
            ReorderAction::Down => settings::move_podcast_feed_down(&mut ps.settings, source_index),
            ReorderAction::Top => {
                settings::move_podcast_feed_to_top(&mut ps.settings, source_index)
            }
            ReorderAction::Bottom => {
                settings::move_podcast_feed_to_bottom(&mut ps.settings, source_index)
            }
            ReorderAction::Position => {
                settings::move_podcast_feed_to_index(&mut ps.settings, source_index, target_index)
            }
        };
        if moved.is_some() {
            settings::save_settings(ps.settings.clone());
        }
        moved
    })
    .flatten();
    let Some(new_index) = new_index else {
        return None;
    };
    if move_vec_to_index(&mut root_items, source_index, new_index) {
        apply_root_order(hwnd, hwnd_tree, &root_items);
    }
    Some(new_index)
}

unsafe fn handle_reorder_action(hwnd: HWND, action: ReorderAction) {
    let Some(source_index) = selected_source_index(hwnd) else {
        return;
    };
    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = with_state(parent, |ps| ps.settings.language).unwrap_or_default();
    let total = with_state(parent, |ps| ps.settings.podcast_sources.len()).unwrap_or(0);
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
            announce_status(&message);
        }
    }
}

unsafe fn show_reorder_dialog(parent_hwnd: HWND, source_index: usize, total: usize) {
    let existing = with_podcast_state(parent_hwnd, |s| s.reorder_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PODCASTS_REORDER_CLASS);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            windows::Win32::UI::WindowsAndMessaging::LoadCursorW(None, IDC_ARROW)
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

    let language = with_podcast_state(parent_hwnd, |s| s.language).unwrap_or_default();
    let title = i18n::tr(language, "podcasts.context.reorder");
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
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        320,
        140,
        parent_hwnd,
        None,
        hinstance,
        Some(init_ptr as *const _),
    );
    if hwnd.0 == 0 {
        let _ = Box::from_raw(init_ptr);
        return;
    }
    with_podcast_state(parent_hwnd, |s| s.reorder_dialog = hwnd);
}

unsafe extern "system" fn reorder_control_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_CHAR {
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
    let prev = GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA);
    if prev == 0 {
        return DefWindowProcW(hwnd, msg, wparam, lparam);
    }
    CallWindowProcW(Some(std::mem::transmute(prev)), hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn reorder_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let init_ptr = (*cs).lpCreateParams as *mut ReorderDialogInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = &*init_ptr;
            let parent = init.parent;
            let source_index = init.source_index;
            let total = init.total;
            SetWindowLongPtrW(
                hwnd,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                init_ptr as isize,
            );
            let language = with_podcast_state(parent, |s| s.language).unwrap_or_default();
            let position_template = i18n::tr(language, "rss.reorder.position_of");
            let position_text = position_template
                .replace("{x}", &(source_index + 1).to_string())
                .replace("{n}", &total.to_string());
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
                280,
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
                32,
                280,
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
                280,
                24,
                hwnd,
                HMENU(REORDER_EDIT_ID as isize),
                hinstance,
                None,
            );
            let ok = CreateWindowExW(
                Default::default(),
                w!("BUTTON"),
                PCWSTR(to_wide(&ok_label).as_ptr()),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WINDOW_STYLE(
                        windows::Win32::UI::WindowsAndMessaging::BS_DEFPUSHBUTTON as u32,
                    ),
                130,
                92,
                70,
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
                210,
                92,
                70,
                24,
                hwnd,
                HMENU(REORDER_CANCEL_ID as isize),
                hinstance,
                None,
            );
            let prev = SetWindowLongPtrW(
                edit,
                windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
                reorder_control_subclass_proc as isize,
            );
            SetWindowLongPtrW(
                edit,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                prev,
            );
            let prev_ok = SetWindowLongPtrW(
                ok,
                windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
                reorder_control_subclass_proc as isize,
            );
            SetWindowLongPtrW(
                ok,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                prev_ok,
            );
            let prev_cancel = SetWindowLongPtrW(
                cancel,
                windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
                reorder_control_subclass_proc as isize,
            );
            SetWindowLongPtrW(
                cancel,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                prev_cancel,
            );
            let text = format!("{}", source_index + 1);
            let _ = SetWindowTextW(edit, PCWSTR(to_wide(&text).as_ptr()));
            SetFocus(edit);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            match id {
                REORDER_OK_ID | 1 => {
                    let ptr = GetWindowLongPtrW(
                        hwnd,
                        windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                    ) as *mut ReorderDialogInit;
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
                    let language =
                        with_podcast_state(init.parent, |s| s.language).unwrap_or_default();
                    let pos = match text.trim().parse::<usize>() {
                        Ok(v) if v > 0 => v,
                        _ => {
                            let message = i18n::tr(language, "rss.reorder.invalid_position");
                            announce_status(&message);
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
                            announce_status(&message);
                        }
                    }
                    let _ = DestroyWindow(hwnd);
                    focus_library(init.parent);
                    LRESULT(0)
                }
                REORDER_CANCEL_ID | 2 => {
                    let ptr = GetWindowLongPtrW(
                        hwnd,
                        windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                    ) as *mut ReorderDialogInit;
                    let parent = if ptr.is_null() {
                        HWND(0)
                    } else {
                        (*ptr).parent
                    };
                    let _ = DestroyWindow(hwnd);
                    if parent.0 != 0 {
                        focus_library(parent);
                    }
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
            let ptr =
                GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                    as *mut ReorderDialogInit;
            if !ptr.is_null() {
                let init = Box::from_raw(ptr);
                with_podcast_state(init.parent, |s| s.reorder_dialog = HWND(0));
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn podcast_tree_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_KEYDOWN
        || msg == windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYDOWN
        || msg == WM_CHAR
    {
        let key = wparam.0 as u32;
        if msg == WM_CHAR && key == VK_RETURN.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                if let Some(item) = selected_episode(parent) {
                    let main_hwnd = with_podcast_state(parent, |s| s.parent).unwrap_or(HWND(0));
                    open_episode_in_player(parent, main_hwnd, &item);
                    return LRESULT(0);
                }
            }
        }
        if key == VK_DELETE.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                handle_source_action(parent, SourceAction::Remove);
                return LRESULT(0);
            }
        }
        if key == VK_RIGHT.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                if let Some(idx) = selected_source_index(parent) {
                    let hitem = selected_tree_item(parent);
                    if hitem.0 != 0 {
                        load_episode_children(parent, hitem, idx, false);
                        let _ = SendMessageW(
                            hwnd,
                            TVM_EXPAND,
                            WPARAM(windows::Win32::UI::Controls::TVE_EXPAND.0 as usize),
                            LPARAM(hitem.0),
                        );
                        return LRESULT(0);
                    }
                }
            }
        }
        if key == VK_LEFT.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                let hitem = selected_tree_item(parent);
                if hitem.0 != 0 {
                    let parent_item = HTREEITEM(
                        SendMessageW(
                            hwnd,
                            TVM_GETNEXTITEM,
                            WPARAM(TVGN_PARENT as usize),
                            LPARAM(hitem.0),
                        )
                        .0,
                    );
                    if parent_item.0 != 0 {
                        let _ = SendMessageW(
                            hwnd,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(parent_item.0),
                        );
                        return LRESULT(0);
                    }
                    if selected_source_index(parent).is_some() {
                        let _ = SendMessageW(
                            hwnd,
                            TVM_EXPAND,
                            WPARAM(windows::Win32::UI::Controls::TVE_COLLAPSE.0 as usize),
                            LPARAM(hitem.0),
                        );
                        return LRESULT(0);
                    }
                }
            }
        }
        if key == VK_RETURN.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                if let Some(item) = selected_episode(parent) {
                    let main_hwnd = with_podcast_state(parent, |s| s.parent).unwrap_or(HWND(0));
                    open_episode_in_player(parent, main_hwnd, &item);
                    return LRESULT(0);
                }
                if let Some(idx) = selected_source_index(parent) {
                    let hitem = selected_tree_item(parent);
                    if hitem.0 != 0 {
                        load_episode_children(parent, hitem, idx, false);
                    }
                    return LRESULT(0);
                }
            }
        }
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
        with_podcast_state(parent, |s| s.tree_proc).unwrap_or(None)
    } else {
        None
    };
    if let Some(proc) = prev_proc {
        CallWindowProcW(Some(proc), hwnd, msg, wparam, lparam)
    } else {
        DefWindowProcW(hwnd, msg, wparam, lparam)
    }
}

unsafe extern "system" fn podcast_search_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_KEYDOWN || msg == windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYDOWN {
        let key = wparam.0 as u32;
        if key == VK_TAB.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                let (
                    hwnd_tree,
                    hwnd_search_button,
                    hwnd_results,
                    hwnd_add,
                    hwnd_import,
                    hwnd_export,
                    hwnd_close,
                ) = with_podcast_state(parent, |s| {
                    (
                        s.hwnd_tree,
                        s.hwnd_search_button,
                        s.hwnd_results,
                        s.hwnd_add,
                        s.hwnd_import,
                        s.hwnd_export,
                        s.hwnd_close,
                    )
                })
                .unwrap_or((
                    HWND(0),
                    HWND(0),
                    HWND(0),
                    HWND(0),
                    HWND(0),
                    HWND(0),
                    HWND(0),
                ));
                let prev = GetKeyState(VK_SHIFT.0 as i32) < 0;
                let target = if prev {
                    hwnd_tree
                } else if hwnd_search_button.0 != 0 {
                    hwnd_search_button
                } else if hwnd_results.0 != 0 {
                    hwnd_results
                } else if hwnd_add.0 != 0 {
                    hwnd_add
                } else if hwnd_import.0 != 0 {
                    hwnd_import
                } else if hwnd_export.0 != 0 {
                    hwnd_export
                } else {
                    hwnd_close
                };
                if target.0 != 0 {
                    SetFocus(target);
                    return LRESULT(0);
                }
            }
        }
        if key == VK_RETURN.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                trigger_search_from_edit(parent);
            }
            return LRESULT(0);
        }
    }
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_KEYUP
        || msg == windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYUP
    {
        let key = wparam.0 as u32;
        if key == VK_RETURN.0 as u32 {
            let parent = GetParent(hwnd);
            if parent.0 != 0 {
                trigger_search_from_edit(parent);
            }
            return LRESULT(0);
        }
    }
    if msg == WM_CHAR && wparam.0 as u32 == 13 {
        let parent = GetParent(hwnd);
        if parent.0 != 0 {
            trigger_search_from_edit(parent);
        }
        return LRESULT(0);
    }
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_SYSCHAR && wparam.0 as u32 == 13 {
        let parent = GetParent(hwnd);
        if parent.0 != 0 {
            trigger_search_from_edit(parent);
        }
        return LRESULT(0);
    }
    let parent = GetParent(hwnd);
    let prev_proc = if parent.0 != 0 {
        with_podcast_state(parent, |s| s.search_proc).unwrap_or(None)
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
            | WINDOW_STYLE(
                (windows::Win32::UI::Controls::TVS_HASLINES
                    | windows::Win32::UI::Controls::TVS_HASBUTTONS
                    | windows::Win32::UI::Controls::TVS_LINESATROOT
                    | windows::Win32::UI::Controls::TVS_SHOWSELALWAYS) as u32,
            ),
        10,
        10,
        460,
        280,
        hwnd,
        HMENU(ID_TREE as isize),
        hinstance,
        None,
    );
    if hwnd_tree.0 != 0 {
        let old = SetWindowLongPtrW(
            hwnd_tree,
            windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
            podcast_tree_wndproc as isize,
        );
        with_podcast_state(hwnd, |s| {
            s.tree_proc = std::mem::transmute::<isize, WNDPROC>(old)
        });
    }

    let hwnd_delete = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.delete_button",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        220,
        300,
        200,
        26,
        hwnd,
        HMENU(ID_DELETE_BUTTON as isize),
        hinstance,
        None,
    );

    let hwnd_search_label = CreateWindowExW(
        Default::default(),
        w!("STATIC"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.search.label",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE,
        10,
        310,
        460,
        16,
        hwnd,
        HMENU(ID_SEARCH_LABEL as isize),
        hinstance,
        None,
    );

    let hwnd_search = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("EDIT"),
        PCWSTR::null(),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        330,
        460,
        24,
        hwnd,
        HMENU(ID_SEARCH_EDIT as isize),
        hinstance,
        None,
    );
    if hwnd_search.0 != 0 {
        let old = SetWindowLongPtrW(
            hwnd_search,
            windows::Win32::UI::WindowsAndMessaging::GWLP_WNDPROC,
            podcast_search_wndproc as isize,
        );
        with_podcast_state(hwnd, |s| {
            s.search_proc = std::mem::transmute::<isize, WNDPROC>(old)
        });
    }

    let provider_itunes = i18n::tr(
        with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
        "podcasts.search.provider.itunes",
    );
    let provider_podcastindex = i18n::tr(
        with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
        "podcasts.search.provider.podcastindex",
    );
    let hwnd_search_provider = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("COMBOBOX"),
        PCWSTR::null(),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
        10,
        364,
        220,
        200,
        hwnd,
        HMENU(ID_SEARCH_PROVIDER as isize),
        hinstance,
        None,
    );
    if hwnd_search_provider.0 != 0 {
        let itunes_wide = to_wide(&provider_itunes);
        let podcastindex_wide = to_wide(&provider_podcastindex);
        let _ = SendMessageW(
            hwnd_search_provider,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(itunes_wide.as_ptr() as isize),
        );
        let _ = SendMessageW(
            hwnd_search_provider,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(podcastindex_wide.as_ptr() as isize),
        );
        let _ = SendMessageW(hwnd_search_provider, CB_SETCURSEL, WPARAM(0), LPARAM(0));
    }

    let hwnd_search_button = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.search.button",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        396,
        140,
        26,
        hwnd,
        HMENU(ID_SEARCH_BUTTON as isize),
        hinstance,
        None,
    );

    let hwnd_results = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("LISTBOX"),
        PCWSTR::null(),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(LBS_NOTIFY as u32),
        10,
        430,
        460,
        140,
        hwnd,
        HMENU(ID_RESULTS as isize),
        hinstance,
        None,
    );

    let hwnd_add = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.add_button",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        490,
        200,
        26,
        hwnd,
        HMENU(ID_ADD_BUTTON as isize),
        hinstance,
        None,
    );

    let hwnd_import = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.import_button",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        526,
        200,
        26,
        hwnd,
        HMENU(ID_IMPORT_BUTTON as isize),
        hinstance,
        None,
    );

    let hwnd_export = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.export_button",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        10,
        562,
        200,
        26,
        hwnd,
        HMENU(ID_EXPORT_BUTTON as isize),
        hinstance,
        None,
    );

    let hwnd_close = CreateWindowExW(
        Default::default(),
        w!("BUTTON"),
        PCWSTR(
            to_wide(&i18n::tr(
                with_podcast_state(hwnd, |s| s.language).unwrap_or_default(),
                "podcasts.close",
            ))
            .as_ptr(),
        ),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        430,
        490,
        200,
        26,
        hwnd,
        HMENU(ID_CLOSE_BUTTON as isize),
        hinstance,
        None,
    );

    with_podcast_state(hwnd, |s| {
        s.hwnd_tree = hwnd_tree;
        s.hwnd_search_label = hwnd_search_label;
        s.hwnd_search = hwnd_search;
        s.hwnd_search_provider = hwnd_search_provider;
        s.hwnd_search_button = hwnd_search_button;
        s.hwnd_results = hwnd_results;
        s.hwnd_add = hwnd_add;
        s.hwnd_import = hwnd_import;
        s.hwnd_export = hwnd_export;
        s.hwnd_delete = hwnd_delete;
        s.hwnd_close = hwnd_close;
    });

    let hfont = HFONT(
        windows::Win32::Graphics::Gdi::GetStockObject(
            windows::Win32::Graphics::Gdi::DEFAULT_GUI_FONT,
        )
        .0,
    );
    for ctrl in [
        hwnd_tree,
        hwnd_search_label,
        hwnd_search,
        hwnd_search_provider,
        hwnd_search_button,
        hwnd_results,
        hwnd_add,
        hwnd_import,
        hwnd_export,
        hwnd_close,
    ] {
        if ctrl.0 != 0 {
            let _ = SendMessageW(ctrl, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }
    }
}

unsafe fn resize_controls(hwnd: HWND) {
    let mut rect = windows::Win32::Foundation::RECT::default();
    if GetClientRect(hwnd, &mut rect).is_err() {
        return;
    }
    let width = (rect.right - rect.left).max(0);
    let height = (rect.bottom - rect.top).max(0);
    let margin = 10;
    let spacing = 8;
    let label_h = 16;
    let search_h = 24;
    let search_button_h = 26;
    let results_h = 140;
    let button_h = 26;
    let button_rows = 3;
    let tree_h = (height
        - margin * 2
        - spacing * 8
        - label_h
        - search_h
        - search_h
        - search_button_h
        - results_h
        - button_h * button_rows)
        .max(120);
    let controls = with_podcast_state(hwnd, |s| {
        (
            s.hwnd_tree,
            s.hwnd_search_label,
            s.hwnd_search,
            s.hwnd_search_provider,
            s.hwnd_search_button,
            s.hwnd_results,
            s.hwnd_add,
            s.hwnd_import,
            s.hwnd_export,
            s.hwnd_close,
        )
    })
    .unwrap_or((
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
        HWND(0),
    ));
    if controls.0 != HWND(0) {
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.0,
            margin,
            margin,
            width - margin * 2,
            tree_h,
            true,
        );
        let mut y = margin + tree_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.1,
            margin,
            y,
            width - margin * 2,
            label_h,
            true,
        );
        y += label_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.2,
            margin,
            y,
            width - margin * 2,
            search_h,
            true,
        );
        y += search_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.3,
            margin,
            y,
            width - margin * 2,
            search_h,
            true,
        );
        y += search_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.4,
            margin,
            y,
            200,
            search_button_h,
            true,
        );
        y += search_button_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.5,
            margin,
            y,
            width - margin * 2,
            results_h,
            true,
        );
        y += results_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.6, margin, y, 200, button_h, true,
        );
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.9,
            (width - margin - 200).max(margin),
            y,
            200,
            button_h,
            true,
        );
        y += button_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.7, margin, y, 200, button_h, true,
        );
        y += button_h + spacing;
        let _ = windows::Win32::UI::WindowsAndMessaging::MoveWindow(
            controls.8, margin, y, 200, button_h, true,
        );
    }
}

pub unsafe fn open(parent: HWND) {
    let exists = with_state(parent, |s| s.podcasts_window).unwrap_or(HWND(0));
    if exists.0 != 0 {
        SetForegroundWindow(exists);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(PODCASTS_WINDOW_CLASS);

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            windows::Win32::UI::WindowsAndMessaging::LoadCursorW(None, IDC_ARROW)
                .unwrap_or_default()
                .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(podcast_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
    let title = to_wide(&i18n::tr(language, "podcasts.window.title"));

    let hwnd = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        windows::Win32::UI::WindowsAndMessaging::CW_USEDEFAULT,
        520,
        560,
        parent,
        None,
        hinstance,
        Some(parent.0 as *const _),
    );

    if hwnd.0 != 0 {
        let _ = with_state(parent, |s| s.podcasts_window = hwnd);
        SetForegroundWindow(hwnd);
    }
}

unsafe extern "system" fn podcast_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let parent = HWND((*cs).lpCreateParams as isize);
            let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
            let state = Box::new(PodcastWindowState {
                parent,
                language,
                hwnd_tree: HWND(0),
                hwnd_search_label: HWND(0),
                hwnd_search: HWND(0),
                hwnd_search_provider: HWND(0),
                hwnd_search_button: HWND(0),
                hwnd_results: HWND(0),
                hwnd_add: HWND(0),
                hwnd_import: HWND(0),
                hwnd_export: HWND(0),
                hwnd_delete: HWND(0),
                hwnd_close: HWND(0),
                node_data: HashMap::new(),
                source_items: HashMap::new(),
                pending_fetches: HashMap::new(),
                search_results: Vec::new(),
                tree_proc: None,
                search_proc: None,
                reorder_dialog: HWND(0),
                last_selected: 0,
                pending_play: None,
            });
            SetWindowLongPtrW(
                hwnd,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                Box::into_raw(state) as isize,
            );
            create_controls(hwnd);
            reload_tree(hwnd);
            update_delete_button_state(hwnd);
            let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
            if hwnd_tree.0 != 0 {
                SetFocus(hwnd_tree);
            }
            LRESULT(0)
        }
        WM_SIZE => {
            resize_controls(hwnd);
            LRESULT(0)
        }
        WM_NOTIFY => {
            let nmhdr = &*(lparam.0 as *const windows::Win32::UI::Controls::NMHDR);
            if nmhdr.idFrom as usize == ID_TREE {
                if nmhdr.code == NM_RETURN {
                    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    if let Some(item) = selected_episode(hwnd) {
                        open_episode_in_player(hwnd, parent, &item);
                        return LRESULT(0);
                    }
                }
                if nmhdr.code == TVN_KEYDOWN {
                    let key = (lparam.0 as *const NMTVKEYDOWN).as_ref();
                    if let Some(key) = key
                        && key.wVKey == VK_RETURN.0 as u16
                    {
                        let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                        if let Some(item) = selected_episode(hwnd) {
                            open_episode_in_player(hwnd, parent, &item);
                            return LRESULT(0);
                        }
                    }
                }
                if nmhdr.code == TVN_ITEMEXPANDINGW {
                    let info = &*(lparam.0 as *const windows::Win32::UI::Controls::NMTREEVIEWW);
                    let hitem = info.itemNew.hItem;
                    if let Some(NodeData::Source(idx)) =
                        with_podcast_state(hwnd, |s| s.node_data.get(&hitem.0).cloned()).flatten()
                    {
                        load_episode_children(hwnd, hitem, idx, false);
                    }
                }
                if nmhdr.code == TVN_SELCHANGEDW {
                    let info = &*(lparam.0 as *const windows::Win32::UI::Controls::NMTREEVIEWW);
                    with_podcast_state(hwnd, |s| s.last_selected = info.itemNew.hItem.0);
                    update_delete_button_state(hwnd);
                }
            }
            LRESULT(0)
        }
        WM_CONTEXTMENU => {
            let x = (lparam.0 & 0xffff) as i16 as i32;
            let y = ((lparam.0 >> 16) & 0xffff) as i16 as i32;
            show_context_menu(hwnd, x, y, false);
            LRESULT(0)
        }
        WM_COMMAND => {
            let id = (wparam.0 & 0xffff) as usize;
            let code = ((wparam.0 >> 16) & 0xffff) as u16;
            match id {
                ID_ADD_BUTTON => {
                    show_add_dialog(hwnd);
                    LRESULT(0)
                }
                ID_IMPORT_BUTTON => {
                    handle_import_opml(hwnd);
                    LRESULT(0)
                }
                ID_EXPORT_BUTTON => {
                    handle_export_opml(hwnd);
                    LRESULT(0)
                }
                ID_CLOSE_BUTTON | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                ID_SEARCH_BUTTON => {
                    if code == windows::Win32::UI::WindowsAndMessaging::BN_CLICKED as u16 {
                        trigger_search_from_edit(hwnd);
                        return LRESULT(0);
                    }
                    LRESULT(0)
                }
                ID_RESULTS => {
                    if code == LBN_DBLCLK as u16 {
                        subscribe_selected_result(hwnd);
                        return LRESULT(0);
                    }
                    LRESULT(0)
                }
                ID_DELETE_BUTTON => {
                    handle_source_action(hwnd, SourceAction::Remove);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            let focus = GetFocus();
            let (hwnd_tree, hwnd_search, hwnd_search_provider, hwnd_results, hwnd_search_button) =
                with_podcast_state(hwnd, |s| {
                    (
                        s.hwnd_tree,
                        s.hwnd_search,
                        s.hwnd_search_provider,
                        s.hwnd_results,
                        s.hwnd_search_button,
                    )
                })
                .unwrap_or((HWND(0), HWND(0), HWND(0), HWND(0), HWND(0)));
            let key = wparam.0 as u32;
            if (focus == hwnd_search || focus == hwnd_search_provider) && key == VK_RETURN.0 as u32
            {
                if hwnd_search_button.0 != 0 {
                    let _ = SendMessageW(
                        hwnd_search_button,
                        windows::Win32::UI::WindowsAndMessaging::BM_CLICK,
                        WPARAM(0),
                        LPARAM(0),
                    );
                } else {
                    trigger_search_from_edit(hwnd);
                }
                return LRESULT(0);
            }
            if focus == hwnd_results && key == VK_RETURN.0 as u32 {
                subscribe_selected_result(hwnd);
                return LRESULT(0);
            }
            if focus == hwnd_tree && key == VK_RETURN.0 as u32 {
                if let Some(item) = selected_episode(hwnd) {
                    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    open_episode_in_player(hwnd, parent, &item);
                    return LRESULT(0);
                }
                if let Some(idx) = selected_source_index(hwnd) {
                    let hitem = selected_tree_item(hwnd);
                    if hitem.0 != 0 {
                        load_episode_children(hwnd, hitem, idx, false);
                    }
                    return LRESULT(0);
                }
            }
            if focus == hwnd_tree && key == VK_RIGHT.0 as u32 {
                if let Some(idx) = selected_source_index(hwnd) {
                    let hitem = selected_tree_item(hwnd);
                    if hitem.0 != 0 {
                        load_episode_children(hwnd, hitem, idx, false);
                        let _ = SendMessageW(
                            hwnd_tree,
                            TVM_EXPAND,
                            WPARAM(windows::Win32::UI::Controls::TVE_EXPAND.0 as usize),
                            LPARAM(hitem.0),
                        );
                    }
                    return LRESULT(0);
                }
            }
            if focus == hwnd_tree && key == VK_LEFT.0 as u32 {
                let hitem = selected_tree_item(hwnd);
                if hitem.0 != 0 {
                    let parent_item = HTREEITEM(
                        SendMessageW(
                            hwnd_tree,
                            TVM_GETNEXTITEM,
                            WPARAM(TVGN_PARENT as usize),
                            LPARAM(hitem.0),
                        )
                        .0,
                    );
                    if parent_item.0 != 0 {
                        let _ = SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(parent_item.0),
                        );
                        return LRESULT(0);
                    }
                    if selected_source_index(hwnd).is_some() {
                        let _ = SendMessageW(
                            hwnd_tree,
                            TVM_EXPAND,
                            WPARAM(windows::Win32::UI::Controls::TVE_COLLAPSE.0 as usize),
                            LPARAM(hitem.0),
                        );
                        return LRESULT(0);
                    }
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_COPYDATA => {
            let cds = &*(lparam.0 as *const COPYDATASTRUCT);
            if cds.dwData == PODCAST_ADD_COPYDATA {
                let url = from_wide(cds.lpData as *const u16);
                let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if let Some(index) = add_podcast_source(parent, &url, "") {
                    let language = with_state(parent, |s| s.settings.language).unwrap_or_default();
                    announce_status(&i18n::tr(language, "podcasts.added"));
                    let hwnd_tree = with_podcast_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                    if hwnd_tree.0 != 0 {
                        let hitem = create_tree_item(hwnd_tree, &url, index);
                        if hitem.0 != 0 {
                            with_podcast_state(hwnd, |s| {
                                s.node_data.insert(hitem.0, NodeData::Source(index));
                            });
                            let _ = SendMessageW(
                                hwnd_tree,
                                TVM_SELECTITEM,
                                WPARAM(TVGN_CARET as usize),
                                LPARAM(hitem.0),
                            );
                            load_episode_children(hwnd, hitem, index, false);
                        }
                    }
                }
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_PODCAST_FETCH_COMPLETE => {
            let ptr = lparam.0 as *mut FetchResult;
            if ptr.is_null() {
                return LRESULT(0);
            }
            let msg = Box::from_raw(ptr);
            with_podcast_state(hwnd, |s| {
                s.pending_fetches.retain(|_, h| *h != msg.hitem);
            });
            match msg.result {
                Ok(outcome) => {
                    let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    if !outcome.title.trim().is_empty() {
                        update_source_title(
                            hwnd,
                            HTREEITEM(msg.hitem),
                            msg.source_index,
                            &outcome.title,
                        );
                    }
                    update_source_cache(parent, msg.source_index, outcome.cache);
                    if outcome.not_modified {
                        apply_episode_results(hwnd, HTREEITEM(msg.hitem), Vec::new());
                    } else if !outcome.items.is_empty() {
                        apply_episode_results(hwnd, HTREEITEM(msg.hitem), outcome.items);
                    }
                }
                Err(err) => {
                    log_debug(&format!("podcasts_fetch_error {}", err));
                }
            }
            LRESULT(0)
        }
        WM_PODCAST_SEARCH_COMPLETE => {
            let ptr = lparam.0 as *mut SearchResultMsg;
            if ptr.is_null() {
                return LRESULT(0);
            }
            let msg = Box::from_raw(ptr);
            update_search_results(hwnd, msg.results);
            LRESULT(0)
        }
        WM_PODCAST_PLAY_READY => {
            let ptr = lparam.0 as *mut PlayReadyMsg;
            if ptr.is_null() {
                return LRESULT(0);
            }
            let msg = Box::from_raw(ptr);
            with_podcast_state(hwnd, |s| s.pending_play = None);
            let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            editor_manager::open_document(parent, &msg.path);
            if parent.0 != 0 {
                editor_manager::mark_current_document_from_rss(parent, true);
                crate::set_active_podcast_episode_info(
                    parent,
                    Some(msg.enclosure_url.clone()),
                    Some(msg.title.clone()),
                    Some(msg.path.clone()),
                );
                crate::menu::update_playback_menu(parent, true);
                crate::activate_pending_podcast_chapters(parent);
            }
            if parent.0 != 0 {
                SetForegroundWindow(parent);
                if let Some(hwnd_tab) = with_state(parent, |s| s.hwnd_tab) {
                    if hwnd_tab.0 != 0 {
                        SetFocus(hwnd_tab);
                    }
                }
            }
            LRESULT(0)
        }
        WM_PODCAST_PLAY_FAILED => {
            with_podcast_state(hwnd, |s| s.pending_play = None);
            let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                crate::set_pending_podcast_chapters_key(parent, None);
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            let parent = with_podcast_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                let _ = with_state(parent, |s| s.podcasts_window = HWND(0));
                force_focus_editor_on_parent(parent);
            }
            let ptr =
                GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                    as *mut PodcastWindowState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_NCDESTROY => DefWindowProcW(hwnd, msg, wparam, lparam),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_podcast_state<R>(
    hwnd: HWND,
    f: impl FnOnce(&mut PodcastWindowState) -> R,
) -> Option<R> {
    let ptr = GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
        as *mut PodcastWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

fn percent_encode(input: &str) -> String {
    use url::form_urlencoded::byte_serialize;
    byte_serialize(input.as_bytes()).collect()
}
