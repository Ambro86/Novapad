use crate::accessibility::{nvda_speak, to_wide};
use crate::app_windows::help_window;
use crate::settings::{ListDateDisplayMode, ListTimeDisplayMode};

use crate::editor_manager;
use crate::i18n;
use crate::log_debug;
use crate::tools::rss::{self, RssFeedCache, RssItem, RssSource, RssSourceType};
use crate::with_state;
use chrono::{Local, NaiveDate, TimeZone};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Write;
use std::mem;
use std::path::{Path, PathBuf};
use url::Url;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
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
    NM_RCLICK, NMHDR, NMTREEVIEWW, NMTVKEYDOWN, TVE_EXPAND, TVGN_CARET, TVGN_CHILD, TVGN_NEXT,
    TVGN_PARENT, TVGN_ROOT, TVHITTESTINFO, TVI_LAST, TVI_ROOT, TVIF_PARAM, TVIF_TEXT,
    TVINSERTSTRUCTW, TVINSERTSTRUCTW_0, TVITEMEXW_CHILDREN, TVITEMW, TVM_DELETEITEM,
    TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_HITTEST, TVM_INSERTITEMW,
    TVM_SELECTITEM, TVM_SETITEMW, TVM_SORTCHILDRENCB, TVN_ITEMEXPANDINGW, TVN_KEYDOWN,
    TVN_SELCHANGEDW, TVSORTCB, WC_BUTTON,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetFocus, GetKeyState, SetActiveWindow, SetFocus, VK_APPS, VK_CONTROL, VK_ESCAPE, VK_F10,
    VK_MENU, VK_RETURN, VK_SHIFT, VK_TAB,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BS_DEFPUSHBUTTON, CHILDID_SELF, CREATESTRUCTW, CW_USEDEFAULT, CallWindowProcW,
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, ES_AUTOHSCROLL,
    EVENT_OBJECT_FOCUS, GWLP_USERDATA, GWLP_WNDPROC, GetCursorPos, GetDlgItem, GetParent,
    GetWindowLongPtrW, GetWindowRect, HMENU, IDYES, KillTimer, MB_ICONINFORMATION, MB_ICONQUESTION,
    MB_OK, MB_YESNO, MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, MessageBoxW, OBJID_CLIENT,
    PostMessageW, RegisterClassW, SW_HIDE, SendMessageW, SetForegroundWindow, SetWindowLongPtrW,
    SetWindowTextW, ShowWindow, TrackPopupMenu, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CONTEXTMENU,
    WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY, WM_NEXTDLGCTL, WM_NOTIFY, WM_NULL,
    WM_SETFOCUS, WM_SETFONT, WM_SETREDRAW, WM_SYSKEYDOWN, WM_TIMER, WM_USER, WNDCLASSW, WNDPROC,
    WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_DLGMODALFRAME, WS_POPUP, WS_SYSMENU, WS_TABSTOP,
    WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR, w};

const RSS_WINDOW_CLASS: &str = "SonarpadRssWindow";

#[inline]
fn ignore_bool(_value: bool) {}
const ID_TREE: usize = 1001;
const ID_BTN_ADD: usize = 1002;
const ID_BTN_CLOSE: usize = 1003;
const ID_BTN_IMPORT: usize = 1004;
const ID_BTN_EXPORT: usize = 1005;
const ID_BTN_SEARCH: usize = 1006;
const ID_CTX_EDIT: usize = 1101;
const ID_CTX_DELETE: usize = 1102;
const ID_CTX_RETRY: usize = 1103;
const ID_CTX_UNDO_DELETE: usize = 1104;
const ID_CTX_REORDER_UP: usize = 1301;
const ID_CTX_REORDER_DOWN: usize = 1302;
const ID_CTX_REORDER_TOP: usize = 1303;
const ID_CTX_REORDER_BOTTOM: usize = 1304;
const ID_CTX_REORDER_POSITION: usize = 1305;
const ID_CTX_SORT_ASC: usize = 1306;
const ID_CTX_SORT_DESC: usize = 1307;
const ID_CTX_SORT_NEWEST: usize = 1308;
const ID_CTX_SORT_OLDEST: usize = 1309;
const ID_CTX_OPEN_BROWSER: usize = 1201;
const ID_CTX_SHARE_FACEBOOK: usize = 1202;
const ID_CTX_SHARE_TWITTER: usize = 1203;
const ID_CTX_SHARE_WHATSAPP: usize = 1204;
const ID_CTX_SHARE_EMAIL: usize = 1205;
const ID_CTX_PROPERTIES: usize = 1206;

const WM_RSS_FETCH_COMPLETE: u32 = WM_USER + 200;
const WM_RSS_IMPORT_COMPLETE: u32 = WM_USER + 201;
const WM_SHOW_ADD_DIALOG: u32 = WM_USER + 202;
const WM_CLEAR_ENTER_GUARD: u32 = WM_USER + 203;
const WM_CLEAR_ADD_GUARD: u32 = WM_USER + 204;
pub(crate) const WM_RSS_SHOW_CONTEXT: u32 = WM_USER + 205;
const WM_RSS_BACKGROUND_CHECK_COMPLETE: u32 = WM_USER + 206;
const WM_RSS_MARK_ITEM_READ_UI: u32 = WM_USER + 207;
const WM_RSS_SELECT_SOURCE_DELAYED: u32 = WM_USER + 208;
const ADD_GUARD_TIMER_ID: usize = 1;
const EM_REPLACESEL: u32 = 0x00C2;
const REORDER_EDIT_ID: usize = 1401;
const REORDER_OK_ID: usize = 1402;
const REORDER_CANCEL_ID: usize = 1403;
const SEARCH_EDIT_ID: usize = 1501;
const SEARCH_OK_ID: usize = 1502;
const SEARCH_CANCEL_ID: usize = 1503;

const FEED_EN_DATA: &str = include_str!("../../i18n/feed_en.txt");
const FEED_UK_DATA: &str = include_str!("../../i18n/feed_uk.txt");
const FEED_IT_DATA: &str = include_str!("../../i18n/feed_it.txt");
const FEED_ES_DATA: &str = include_str!("../../i18n/feed_es.txt");
const FEED_PT_DATA: &str = include_str!("../../i18n/feed_pt.txt");
const FEED_VI_DATA: &str = include_str!("../../i18n/feed_vi.txt");
const FEED_CS_DATA: &str = include_str!("../../i18n/feed_cs.txt");
const FEED_PL_DATA: &str = include_str!("../../i18n/feed_pl.txt");
const FEED_FR_DATA: &str = include_str!("../../i18n/feed_fr.txt");
const FEED_SR_DATA: &str = include_str!("../../i18n/feed_sr HR.txt");
const EM_SETSEL: u32 = 0x00B1;
const EM_LIMITTEXT: u32 = 0x00C5;
const INITIAL_LOAD_COUNT: usize = 5;
const LOAD_MORE_COUNT: usize = 5;

// Normalize article text before sending it to the editor:
// - collapse multiple blank lines to a single blank line
// - replace embedded NULs (which would truncate Win32 edit text)
fn normalize_article_text(s: &str) -> String {
    let decoded = decode_basic_html_entities(s);
    let no_nul: String = decoded
        .chars()
        .map(|c| if c == '\0' { ' ' } else { c })
        .collect();
    collapse_blank_lines(&no_nul)
}

#[cfg(test)]
mod tests {
    use super::{
        decode_basic_html_entities, format_google_news_source_title, normalize_rss_url_key,
    };

    #[test]
    fn normalize_rss_url_key_keeps_query_parameters() {
        let key = normalize_rss_url_key("https://servis.idnes.cz/rss.aspx?c=kultura");
        assert_eq!(key, "servis.idnes.cz/rss.aspx?c=kultura");
    }

    #[test]
    fn normalize_rss_url_key_removes_fragment_only() {
        let key = normalize_rss_url_key("https://servis.idnes.cz/rss.aspx?c=technet#section");
        assert_eq!(key, "servis.idnes.cz/rss.aspx?c=technet");
    }

    #[test]
    fn normalize_rss_url_key_normalizes_scheme_case_and_trailing_slash() {
        let key = normalize_rss_url_key("HTTP://EXAMPLE.COM/Feed/");
        assert_eq!(key, "example.com/feed");
    }

    #[test]
    fn format_google_news_source_title_capitalizes_each_word() {
        let title = format_google_news_source_title("elon musk");
        assert_eq!(title, "Elon Musk");
    }

    #[test]
    fn decode_basic_html_entities_decodes_common_italian_entities() {
        let text = "cos&igrave; si &egrave; visto, &laquo;ok&raquo;";
        let decoded = decode_basic_html_entities(text);
        assert_eq!(decoded, "così si è visto, «ok»");
    }
}

fn decode_basic_html_entities(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c != '&' {
            out.push(c);
            continue;
        }

        let mut entity = String::new();
        let mut ended_with_semicolon = false;
        while let Some(&next) = chars.peek() {
            chars.next();
            if next == ';' {
                ended_with_semicolon = true;
                break;
            }
            // Keep entity names bounded to avoid eating large text chunks.
            if entity.len() >= 16 {
                entity.push(next);
                break;
            }
            entity.push(next);
        }

        let decoded = if entity.starts_with("#x") || entity.starts_with("#X") {
            u32::from_str_radix(&entity[2..], 16)
                .ok()
                .and_then(char::from_u32)
        } else if let Some(num) = entity.strip_prefix('#') {
            num.parse::<u32>().ok().and_then(char::from_u32)
        } else {
            match entity.as_str() {
                "nbsp" => Some(' '),
                "amp" => Some('&'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                "lt" => Some('<'),
                "gt" => Some('>'),
                "laquo" => Some('«'),
                "raquo" => Some('»'),
                "hellip" => Some('…'),
                "ndash" => Some('–'),
                "mdash" => Some('—'),
                "rsquo" => Some('’'),
                "lsquo" => Some('‘'),
                "rdquo" => Some('”'),
                "ldquo" => Some('“'),
                // Latin entities commonly seen in Italian and other EU feeds.
                "agrave" => Some('à'),
                "egrave" => Some('è'),
                "igrave" => Some('ì'),
                "ograve" => Some('ò'),
                "ugrave" => Some('ù'),
                "aacute" => Some('á'),
                "eacute" => Some('é'),
                "iacute" => Some('í'),
                "oacute" => Some('ó'),
                "uacute" => Some('ú'),
                "Agrave" => Some('À'),
                "Egrave" => Some('È'),
                "Igrave" => Some('Ì'),
                "Ograve" => Some('Ò'),
                "Ugrave" => Some('Ù'),
                "Aacute" => Some('Á'),
                "Eacute" => Some('É'),
                "Iacute" => Some('Í'),
                "Oacute" => Some('Ó'),
                "Uacute" => Some('Ú'),
                _ => None,
            }
        };

        if let Some(ch) = decoded {
            out.push(ch);
        } else {
            out.push('&');
            out.push_str(&entity);
            if ended_with_semicolon {
                out.push(';');
            }
        }
    }
    out
}

fn normalize_rss_url_key(url: &str) -> String {
    let mut s = url.trim().to_string();
    if s.is_empty() {
        return s;
    }
    if s.get(..8)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("https://"))
    {
        s = s[8..].to_string();
    } else if s
        .get(..7)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("http://"))
    {
        s = s[7..].to_string();
    }
    if let Some((left, _)) = s.split_once('#') {
        s = left.to_string();
    }
    while s.ends_with('/') && s.len() > 1 {
        s.pop();
    }
    s.to_ascii_lowercase()
}

fn rss_source_display_title(
    source: &RssSource,
    language: crate::settings::Language,
    announce_unread: bool,
    unread_label_position: crate::settings::RssPodcastUnreadLabelPosition,
) -> String {
    let base_title = if source.title.trim().is_empty() {
        source.url.clone()
    } else {
        source.title.clone()
    };
    if announce_unread && source.unread {
        match unread_label_position {
            crate::settings::RssPodcastUnreadLabelPosition::Before => {
                format!("{}{}", i18n::tr(language, "rss.unread_prefix"), base_title)
            }
            crate::settings::RssPodcastUnreadLabelPosition::After => {
                format!("{base_title}{}", i18n::tr(language, "rss.unread_suffix"))
            }
        }
    } else {
        base_title
    }
}

#[derive(Clone, Copy)]
struct RssItemTitleContext {
    language: crate::settings::Language,
    announce_unread: bool,
    unread_label_position: crate::settings::RssPodcastUnreadLabelPosition,
    date_mode: ListDateDisplayMode,
    time_mode: ListTimeDisplayMode,
}

fn rss_item_display_title(
    title: &str,
    item_unread: bool,
    pub_date: Option<i64>,
    has_multiple_items_same_day: bool,
    ctx: RssItemTitleContext,
) -> String {
    let base = format!(
        "{title}{}",
        format_timestamp_for_list(
            pub_date,
            ctx.language,
            ctx.date_mode,
            ctx.time_mode,
            has_multiple_items_same_day,
        )
        .map(|ts| format!(". {ts}"))
        .unwrap_or_default()
    );
    if ctx.announce_unread && item_unread {
        return match ctx.unread_label_position {
            crate::settings::RssPodcastUnreadLabelPosition::Before => {
                format!(
                    "{}{}",
                    i18n::tr(ctx.language, "rss.item_unread_prefix"),
                    base
                )
            }
            crate::settings::RssPodcastUnreadLabelPosition::After => {
                format!("{base}{}", i18n::tr(ctx.language, "rss.item_unread_suffix"))
            }
        };
    }
    base
}

fn day_from_timestamp(timestamp: Option<i64>) -> Option<NaiveDate> {
    let ts = timestamp?;
    Local
        .timestamp_opt(ts, 0)
        .single()
        .map(|dt| dt.date_naive())
}

fn build_day_counts(items: &[RssItem]) -> HashMap<NaiveDate, usize> {
    let mut counts: HashMap<NaiveDate, usize> = HashMap::new();
    for item in items {
        if let Some(day) = day_from_timestamp(item.pub_date) {
            *counts.entry(day).or_insert(0) += 1;
        }
    }
    counts
}

fn has_multiple_items_same_day(
    pub_date: Option<i64>,
    day_counts: &HashMap<NaiveDate, usize>,
) -> bool {
    day_from_timestamp(pub_date)
        .and_then(|d| day_counts.get(&d).copied())
        .is_some_and(|count| count > 1)
}

fn format_timestamp_for_list(
    timestamp: Option<i64>,
    language: crate::settings::Language,
    date_mode: ListDateDisplayMode,
    time_mode: ListTimeDisplayMode,
    has_multiple_same_day: bool,
) -> Option<String> {
    let ts = timestamp?;
    let dt = Local.timestamp_opt(ts, 0).single()?;
    let (date_pattern, time_pattern) = match language {
        crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => ("%m/%d/%Y", "%I:%M %p"),
        crate::settings::Language::Italian => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Spanish => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Portuguese => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Swedish => ("%Y-%m-%d", "%H:%M"),
        crate::settings::Language::Vietnamese => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Czech => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::Polish => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::French => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Serbian => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::Ukrainian => ("%d.%m.%Y", "%H:%M"),
    };
    let show_date = matches!(date_mode, ListDateDisplayMode::Always);
    let show_time = match time_mode {
        ListTimeDisplayMode::Always => true,
        ListTimeDisplayMode::Never => false,
        ListTimeDisplayMode::OnlyIfMultipleSameDay => has_multiple_same_day,
    };
    if !show_date && !show_time {
        return None;
    }
    if show_date && !show_time {
        let now = Local::now().date_naive();
        let item_day = dt.date_naive();
        let day_diff = (now - item_day).num_days();
        if day_diff == 0 {
            return Some(i18n::tr(language, "rss.date.today"));
        }
        if day_diff == 1 {
            return Some(i18n::tr(language, "rss.date.yesterday"));
        }
        if day_diff == 2 {
            return Some(i18n::tr(language, "rss.date.day_before_yesterday"));
        }
        return Some(dt.format(date_pattern).to_string());
    }
    if !show_date && show_time {
        return Some(dt.format(time_pattern).to_string());
    }
    format_timestamp_for_language(timestamp, language)
}

fn format_timestamp_for_language(
    timestamp: Option<i64>,
    language: crate::settings::Language,
) -> Option<String> {
    let ts = timestamp?;
    let dt = Local.timestamp_opt(ts, 0).single()?;
    let (date_pattern, time_pattern) = match language {
        crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => ("%m/%d/%Y", "%I:%M %p"),
        crate::settings::Language::Italian => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Spanish => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Portuguese => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Swedish => ("%Y-%m-%d", "%H:%M"),
        crate::settings::Language::Vietnamese => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Czech => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::Polish => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::French => ("%d/%m/%Y", "%H:%M"),
        crate::settings::Language::Serbian => ("%d.%m.%Y", "%H:%M"),
        crate::settings::Language::Ukrainian => ("%d.%m.%Y", "%H:%M"),
    };
    let now = Local::now().date_naive();
    let item_day = dt.date_naive();
    let day_diff = (now - item_day).num_days();
    let time_text = dt.format(time_pattern).to_string();
    if day_diff == 0 {
        return Some(format!(
            "{} {time_text}",
            i18n::tr(language, "rss.date.today")
        ));
    }
    if day_diff == 1 {
        return Some(format!(
            "{} {time_text}",
            i18n::tr(language, "rss.date.yesterday")
        ));
    }
    if day_diff == 2 {
        return Some(format!(
            "{} {time_text}",
            i18n::tr(language, "rss.date.day_before_yesterday")
        ));
    }
    let full_pattern = format!("{date_pattern} {time_pattern}");
    Some(dt.format(&full_pattern).to_string())
}

fn show_selected_properties(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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

    let language = with_rss_state(hwnd, |s| {
        { with_state(s.parent, |ps| ps.settings.language) }.unwrap_or_default()
    })
    .unwrap_or_default();

    let (node, source, item_unread) = with_rss_state(hwnd, |s| {
        let node = s.node_data.get(&hitem.0).cloned();
        let source = match &node {
            Some(NodeData::Source(idx)) => {
                { with_state(s.parent, |ps| ps.settings.rss_sources.get(*idx).cloned()) }.flatten()
            }
            _ => None,
        };
        let item_unread = match &node {
            Some(NodeData::Item(item)) => {
                let key = rss_item_key(item);
                let parent_item = windows::Win32::UI::Controls::HTREEITEM(
                    crate::send_message_w_safe(
                        s.hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(TVGN_PARENT as usize),
                        LPARAM(hitem.0),
                    )
                    .0,
                );
                let unread = s
                    .source_items
                    .get(&parent_item.0)
                    .map(|state| !state.read_item_keys.contains(&key))
                    .unwrap_or(true);
                Some(unread)
            }
            _ => None,
        };
        (node, source, item_unread)
    })
    .unwrap_or((None, None, None));

    let mut lines: Vec<String> = Vec::new();
    match node {
        Some(NodeData::Source(_)) => {
            if let Some(src) = source {
                let source_type = match src.kind {
                    RssSourceType::Feed => i18n::tr(language, "properties.feed"),
                    RssSourceType::Site => i18n::tr(language, "properties.site"),
                    RssSourceType::Article => i18n::tr(language, "properties.article"),
                };
                let title_value = if src.title.trim().is_empty() {
                    src.url.clone()
                } else {
                    src.title
                };
                let status_value = if src.unread {
                    i18n::tr(language, "properties.new_items")
                } else {
                    i18n::tr(language, "properties.no_new_items")
                };
                lines.push(format!(
                    "{}: {}",
                    i18n::tr(language, "properties.type"),
                    i18n::tr(language, "properties.source")
                ));
                lines.push(format!(
                    "{}: {}",
                    i18n::tr(language, "properties.title"),
                    title_value
                ));
                lines.push(format!(
                    "{}: {}",
                    i18n::tr(language, "properties.kind"),
                    source_type
                ));
                lines.push(format!(
                    "{}: {}",
                    i18n::tr(language, "properties.url"),
                    src.url
                ));
                lines.push(format!(
                    "{}: {}",
                    i18n::tr(language, "properties.status"),
                    status_value
                ));
            }
        }
        Some(NodeData::Item(item)) => {
            let status_value = if item_unread.unwrap_or(true) {
                i18n::tr(language, "properties.unread")
            } else {
                i18n::tr(language, "properties.read")
            };
            let date_value = format_timestamp_for_language(item.pub_date, language)
                .unwrap_or_else(|| i18n::tr(language, "properties.not_available"));
            lines.push(format!(
                "{}: {}",
                i18n::tr(language, "properties.type"),
                i18n::tr(language, "properties.article")
            ));
            lines.push(format!(
                "{}: {}",
                i18n::tr(language, "properties.title"),
                item.title
            ));
            lines.push(format!(
                "{}: {}",
                i18n::tr(language, "properties.url"),
                item.link
            ));
            lines.push(format!(
                "{}: {}",
                i18n::tr(language, "properties.date"),
                date_value
            ));
            lines.push(format!(
                "{}: {}",
                i18n::tr(language, "properties.status"),
                status_value
            ));
        }
        None => return,
    }

    let body = lines.join("\r\n");
    help_window::open_readonly_text(hwnd, &i18n::tr(language, "properties.title_window"), &body);
}

fn percent_encode(input: &str) -> String {
    url::form_urlencoded::byte_serialize(input.as_bytes()).collect()
}

fn mailto_encode_component(input: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut out = String::with_capacity(input.len() * 3);
    for b in input.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' => {
                out.push(b as char)
            }
            _ => {
                out.push('%');
                out.push(HEX[(b >> 4) as usize] as char);
                out.push(HEX[(b & 0x0F) as usize] as char);
            }
        }
    }
    out
}

fn decode_mail_text_component(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0usize;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let h1 = bytes[i + 1];
            let h2 = bytes[i + 2];
            let v1 = (h1 as char).to_digit(16);
            let v2 = (h2 as char).to_digit(16);
            if let (Some(a), Some(b)) = (v1, v2) {
                out.push(((a << 4) | b) as u8);
                i += 3;
                continue;
            }
        }
        out.push(if bytes[i] == b'+' { b' ' } else { bytes[i] });
        i += 1;
    }
    decode_basic_html_entities(&String::from_utf8_lossy(&out))
        .trim()
        .to_string()
}

fn tr_or(language: crate::settings::Language, key: &str, fallback: &str) -> String {
    let translated = i18n::tr(language, key);
    if translated == key {
        fallback.to_string()
    } else {
        translated
    }
}

fn google_news_params(
    language: crate::settings::Language,
) -> (&'static str, &'static str, &'static str) {
    match language {
        crate::settings::Language::Italian => ("it", "IT", "IT:it"),
        crate::settings::Language::Spanish => ("es", "ES", "ES:es"),
        crate::settings::Language::Portuguese => ("pt", "PT", "PT:pt"),
        crate::settings::Language::Swedish => ("sv", "SE", "SE:sv"),
        crate::settings::Language::Vietnamese => ("vi", "VN", "VN:vi"),
        crate::settings::Language::Czech => ("cs", "CZ", "CZ:cs"),
        crate::settings::Language::Polish => ("pl", "PL", "PL:pl"),
        crate::settings::Language::French => ("fr", "FR", "FR:fr"),
        crate::settings::Language::Serbian => ("sr", "RS", "RS:sr"),
        crate::settings::Language::Ukrainian
        | crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => ("en", "US", "US:en"),
    }
}

fn build_google_news_rss_url(keyword: &str, language: crate::settings::Language) -> String {
    let (hl, gl, ceid) = google_news_params(language);
    let query = percent_encode(keyword.trim());
    format!(
        "https://news.google.com/rss/search?q={}&hl={}&gl={}&ceid={}",
        query, hl, gl, ceid
    )
}

fn format_google_news_source_title(keyword: &str) -> String {
    keyword
        .split_whitespace()
        .map(|word| {
            let mut chars = word.chars();
            if let Some(first) = chars.next() {
                let mut out = String::new();
                out.extend(first.to_uppercase());
                for ch in chars {
                    out.extend(ch.to_lowercase());
                }
                out
            } else {
                String::new()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
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
    reader.config_mut().trim_text(true);
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
                        .decode_and_unescape_value(reader.decoder())
                        .unwrap_or_default()
                        .to_string();
                    if key.eq_ignore_ascii_case(b"xmlUrl") {
                        url = value;
                    } else if title.is_empty()
                        && (key.eq_ignore_ascii_case(b"title") || key.eq_ignore_ascii_case(b"text"))
                    {
                        title = value;
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

fn open_export_opml_dialog(hwnd: HWND, language: crate::settings::Language) -> Option<PathBuf> {
    let filter_raw = i18n::tr(language, "rss.import_filter");
    let filter = to_wide(&filter_raw.replace("\\0", "\0"));
    let mut buffer = vec![0u16; 4096];
    let mut ofn = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: PCWSTR(filter.as_ptr()),
        lpstrFile: PWSTR(buffer.as_mut_ptr()),
        nMaxFile: buffer.len() as u32,
        Flags: OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY,
        ..Default::default()
    };
    if !unsafe { GetSaveFileNameW(&mut ofn).as_bool() } {
        return None;
    }
    parse_single_path(&buffer)
}

fn escape_opml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn export_sources_to_opml_file(hwnd: HWND, path: &Path) -> Result<usize, String> {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return Err("missing parent".to_string());
    }
    let sources =
        { with_state(parent, |state| state.settings.rss_sources.clone()) }.unwrap_or_default();
    if sources.is_empty() {
        return Ok(0);
    }

    let mut file = File::create(path).map_err(|e| e.to_string())?;
    writeln!(
        file,
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<opml version=\"1.0\">\n<head>\n<title>Sonarpad RSS</title>\n</head>\n<body>"
    )
    .map_err(|e| e.to_string())?;

    for src in &sources {
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

fn import_sources_from_file(hwnd: HWND, path: &Path) {
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
    let mut added = 0;
    if with_state(parent, |state| {
        let mut existing: std::collections::HashSet<String> = state
            .settings
            .rss_sources
            .iter()
            .map(|s| s.url.clone())
            .collect();
        for (title, url) in opml_sources {
            let key = url.clone();
            if existing.contains(&key) {
                continue;
            }
            state.settings.rss_sources.push(RssSource {
                title: title.clone(),
                url: url.clone(),
                kind: RssSourceType::Feed,
                user_title: title.trim() != url.trim(),
                unread: false,
                cache: RssFeedCache::default(),
                last_seen_guid: None,
                last_updated: None,
                removed_item_keys: Vec::new(),
                read_item_keys: Vec::new(),
            });
            existing.insert(key);
            added += 1;
        }
        if added > 0 {
            crate::settings::save_settings(state.settings.clone());
        }
    })
    .is_none()
    {
        crate::log_debug("Failed to access state in import_opml_file");
    }
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

fn select_newest_item_key(items: &[RssItem], removed_keys: &HashSet<String>) -> Option<String> {
    let mut best: Option<(i64, usize, String)> = None;
    for (idx, item) in items.iter().enumerate() {
        let key = rss_item_key(item);
        if removed_keys.contains(&key) {
            continue;
        }
        let ts = item.pub_date.unwrap_or(i64::MIN);
        match &best {
            Some((best_ts, best_idx, _))
                if ts < *best_ts || (ts == *best_ts && idx >= *best_idx) => {}
            _ => best = Some((ts, idx, key)),
        }
    }
    best.map(|(_, _, key)| key)
}

fn source_removed_keys_for_tree_item(
    hwnd: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
) -> HashSet<String> {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return HashSet::new();
    }
    let source_index = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten();
    let Some(idx) = source_index else {
        return HashSet::new();
    };
    {
        with_state(parent, |ps| {
            ps.settings
                .rss_sources
                .get(idx)
                .map(|src| src.removed_item_keys.iter().cloned().collect())
        })
        .flatten()
        .unwrap_or_default()
    }
}

fn sort_items_by_date_desc(items: &mut [RssItem]) {
    let mut indexed: Vec<(usize, RssItem)> = items.iter().cloned().enumerate().collect();
    indexed.sort_by(|(ia, a), (ib, b)| {
        let at = a.pub_date.unwrap_or(i64::MIN);
        let bt = b.pub_date.unwrap_or(i64::MIN);
        bt.cmp(&at).then_with(|| ia.cmp(ib))
    });
    for (dst, (_, item)) in indexed.into_iter().enumerate() {
        items[dst] = item;
    }
}

fn prune_persisted_read_keys_for_source(
    hwnd: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
) {
    let source_index = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten();
    let Some(source_index) = source_index else {
        return;
    };

    let current_item_keys: HashSet<String> = with_rss_state(hwnd, |s| {
        s.source_items
            .get(&hitem.0)
            .map(|state| state.items.iter().map(rss_item_key).collect())
            .unwrap_or_default()
    })
    .unwrap_or_default();

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    {
        with_state(parent, |ps| {
            if let Some(src) = ps.settings.rss_sources.get_mut(source_index) {
                let before = src.read_item_keys.len();
                src.read_item_keys.retain(|k| current_item_keys.contains(k));
                if src.read_item_keys.len() != before {
                    crate::settings::save_settings(ps.settings.clone());
                }
            }
        });
    }
}

unsafe extern "system" fn rss_tree_compare(
    lparam1: LPARAM,
    lparam2: LPARAM,
    _lparam_sort: LPARAM,
) -> i32 {
    crate::panic_guard::guard(
        "rss_tree_compare",
        || 0,
        || {
            let a = lparam1.0;
            let b = lparam2.0;
            a.cmp(&b) as i32
        },
    )
}

fn collect_root_items(hwnd_tree: HWND) -> Vec<windows::Win32::UI::Controls::HTREEITEM> {
    let mut items = Vec::new();
    let mut current = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_ROOT as usize),
            LPARAM(0),
        )
        .0,
    );
    while current.0 != 0 {
        items.push(current);
        current = windows::Win32::UI::Controls::HTREEITEM(
            crate::send_message_w_safe(
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

fn select_first_root_if_needed(hwnd: HWND, hwnd_tree: HWND) {
    if hwnd_tree.0 == 0 {
        return;
    }
    let current = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
        .0,
    );
    if current.0 != 0 {
        return;
    }
    let first = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
            hwnd_tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_ROOT as usize),
            LPARAM(0),
        )
        .0,
    );
    if first.0 != 0 {
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_SELECTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(first.0),
            );
            SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(first.0));
        }
        with_rss_state(hwnd, |s| s.last_selected = first.0);
    }
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
        crate::send_message_w_safe(
            hwnd_tree,
            TVM_SETITEMW,
            WPARAM(0),
            LPARAM(&mut item as *mut _ as isize),
        );
    }
    with_rss_state(hwnd, |s| {
        for (i, hitem) in ordered_items.iter().enumerate() {
            s.node_data.insert(hitem.0, NodeData::Source(i));
        }
    });
    let mut sort_cb = TVSORTCB {
        hParent: TVI_ROOT,
        lpfnCompare: Some(rss_tree_compare),
        lParam: LPARAM(0),
    };
    crate::send_message_w_safe(
        hwnd_tree,
        TVM_SORTCHILDRENCB,
        WPARAM(0),
        LPARAM(&mut sort_cb as *mut _ as isize),
    );
}

fn announce_rss_status(message: &str) {
    log_debug(&format!("rss_status {}", message));
    if !nvda_speak(message) {
        crate::log_debug("NVDA speak failed");
    }
}

fn rss_page_sizes(parent: HWND) -> (usize, usize) {
    {
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
    let config = {
        with_state(parent, |s| rss::config_from_settings(&s.settings))
            .unwrap_or_else(rss::RssHttpConfig::default)
    };
    if let Err(err) = rss::init_http(config) {
        log_debug(&format!("rss_http_init_error: {}", err));
    }
}

fn rss_fetch_config(parent: HWND) -> rss::RssFetchConfig {
    {
        with_state(parent, |s| rss::fetch_config_from_settings(&s.settings))
            .unwrap_or_else(rss::RssFetchConfig::default)
    }
}

fn default_feed_path(language: crate::settings::Language) -> Option<PathBuf> {
    let file_name = match language {
        crate::settings::Language::Ukrainian => "feed_uk.txt",
        crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => "feed_en.txt",
        crate::settings::Language::Italian => "feed_it.txt",
        crate::settings::Language::Spanish => "feed_es.txt",
        crate::settings::Language::Portuguese => "feed_pt.txt",
        crate::settings::Language::Swedish => "feed_en.txt",
        crate::settings::Language::Vietnamese => "feed_vi.txt",
        crate::settings::Language::Czech => "feed_cs.txt",
        crate::settings::Language::Polish => "feed_pl.txt",
        crate::settings::Language::French => "feed_fr.txt",
        crate::settings::Language::Serbian => "feed_sr HR.txt",
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
        crate::settings::Language::Ukrainian => FEED_UK_DATA,
        crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => FEED_EN_DATA,
        crate::settings::Language::Italian => FEED_IT_DATA,
        crate::settings::Language::Spanish => FEED_ES_DATA,
        crate::settings::Language::Portuguese => FEED_PT_DATA,
        crate::settings::Language::Swedish => FEED_EN_DATA,
        crate::settings::Language::Vietnamese => FEED_VI_DATA,
        crate::settings::Language::Czech => FEED_CS_DATA,
        crate::settings::Language::Polish => FEED_PL_DATA,
        crate::settings::Language::French => FEED_FR_DATA,
        crate::settings::Language::Serbian => FEED_SR_DATA,
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
        crate::settings::Language::Ukrainian
        | crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => settings
            .rss_default_en_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Swedish => settings
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
            .rss_default_vi_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Czech => settings
            .rss_default_cs_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Polish => settings
            .rss_default_pl_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::French => settings
            .rss_default_fr_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
        crate::settings::Language::Serbian => settings
            .rss_default_sr_keys
            .iter()
            .any(|k| normalize_rss_url_key(k) == key),
    }
}

fn apply_default_sources(
    rss_sources: &mut Vec<RssSource>,
    removed_list: &[String],
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
        if (current_default_keys.contains(&key) || stored_keys.contains(&key))
            && !seen_keys.insert(key)
        {
            return false;
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
            last_updated: None,
            removed_item_keys: Vec::new(),
            read_item_keys: Vec::new(),
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

fn ensure_default_sources(parent: HWND) {
    let language = { with_state(parent, |s| s.settings.language) }.unwrap_or_default();
    let defaults = load_default_feeds(language);
    if defaults.is_empty() {
        return;
    }
    {
        with_state(parent, |s| {
            let changed = match language {
                crate::settings::Language::Ukrainian
                | crate::settings::Language::English
                | crate::settings::Language::Lithuanian
                | crate::settings::Language::Chinese => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_en,
                    &mut s.settings.rss_default_en_keys,
                    &defaults,
                ),
                crate::settings::Language::Swedish => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_en,
                    &mut s.settings.rss_default_en_keys,
                    &defaults,
                ),
                crate::settings::Language::Italian => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_it,
                    &mut s.settings.rss_default_it_keys,
                    &defaults,
                ),
                crate::settings::Language::Spanish => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_es,
                    &mut s.settings.rss_default_es_keys,
                    &defaults,
                ),
                crate::settings::Language::Portuguese => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_pt,
                    &mut s.settings.rss_default_pt_keys,
                    &defaults,
                ),
                crate::settings::Language::Vietnamese => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_vi,
                    &mut s.settings.rss_default_vi_keys,
                    &defaults,
                ),
                crate::settings::Language::Czech => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_cs,
                    &mut s.settings.rss_default_cs_keys,
                    &defaults,
                ),
                crate::settings::Language::Polish => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_pl,
                    &mut s.settings.rss_default_pl_keys,
                    &defaults,
                ),
                crate::settings::Language::French => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_fr,
                    &mut s.settings.rss_default_fr_keys,
                    &defaults,
                ),
                crate::settings::Language::Serbian => apply_default_sources(
                    &mut s.settings.rss_sources,
                    &s.settings.rss_removed_default_sr,
                    &mut s.settings.rss_default_sr_keys,
                    &defaults,
                ),
            };
            if changed {
                crate::settings::save_settings(s.settings.clone());
            }
        });
    }
}

pub(crate) fn sync_default_sources_for_settings(
    settings: &mut crate::settings::AppSettings,
) -> bool {
    let language = settings.language;
    let defaults = load_default_feeds(language);
    if defaults.is_empty() {
        return false;
    }
    match language {
        crate::settings::Language::Ukrainian
        | crate::settings::Language::English
        | crate::settings::Language::Lithuanian
        | crate::settings::Language::Chinese => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_en,
            &mut settings.rss_default_en_keys,
            &defaults,
        ),
        crate::settings::Language::Swedish => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_en,
            &mut settings.rss_default_en_keys,
            &defaults,
        ),
        crate::settings::Language::Italian => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_it,
            &mut settings.rss_default_it_keys,
            &defaults,
        ),
        crate::settings::Language::Spanish => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_es,
            &mut settings.rss_default_es_keys,
            &defaults,
        ),
        crate::settings::Language::Portuguese => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_pt,
            &mut settings.rss_default_pt_keys,
            &defaults,
        ),
        crate::settings::Language::Vietnamese => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_vi,
            &mut settings.rss_default_vi_keys,
            &defaults,
        ),
        crate::settings::Language::Czech => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_cs,
            &mut settings.rss_default_cs_keys,
            &defaults,
        ),
        crate::settings::Language::Polish => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_pl,
            &mut settings.rss_default_pl_keys,
            &defaults,
        ),
        crate::settings::Language::French => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_fr,
            &mut settings.rss_default_fr_keys,
            &defaults,
        ),
        crate::settings::Language::Serbian => apply_default_sources(
            &mut settings.rss_sources,
            &settings.rss_removed_default_sr,
            &mut settings.rss_default_sr_keys,
            &defaults,
        ),
    }
}

struct RssWindowState {
    parent: HWND,
    hwnd_tree: HWND,
    hwnd_import: HWND,
    hwnd_export: HWND,
    node_data: HashMap<isize, NodeData>,
    pending_fetches: HashMap<String, isize>, // URL -> hItem
    source_items: HashMap<isize, SourceItemsState>,
    enter_guard: bool,
    add_guard: bool,
    reorder_dialog: HWND,
    search_dialog: HWND,
    pending_edit: Option<usize>,
    tree_proc: WNDPROC,
    last_selected: isize,
    removed_history: Vec<RssLastRemoved>,
    suppress_tree_selection_events: bool,
    suppress_focus_restore_once: bool,
}

#[derive(Clone)]
enum RssLastRemoved {
    Source {
        index: usize,
        source: RssSource,
        language: crate::settings::Language,
        default_removed_key_added: Option<String>,
    },
    Item {
        source_index: usize,
        item: RssItem,
        key: String,
        position: usize,
    },
}

#[derive(Clone)]
enum NodeData {
    Source(usize), // Index in settings
    Item(RssItem),
}

struct SourceItemsState {
    items: Vec<RssItem>,
    loaded: usize,
    read_item_keys: HashSet<String>,
}

struct AddDialogInit {
    parent: HWND,
    prefill_title: String,
    prefill_url: String,
    hide_url_field: bool,
}

struct SearchDialogInit {
    parent: HWND,
}

pub fn open(parent: HWND) {
    unsafe {
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

        if hwnd.0 != 0 && with_state(parent, |s| s.rss_window = hwnd).is_none() {
            crate::log_debug("Failed to set rss_window state");
            // IMPORTANT: do NOT disable the parent window.
            // If the parent is disabled, Windows (and NVDA) treat the editor as unavailable,
            // and SetFocus/SetForegroundWindow will not behave reliably.
        }
    }
}

pub fn show_context_menu_from_keyboard(hwnd: HWND) {
    let mut pt = POINT::default();
    crate::log_if_err!(unsafe { GetCursorPos(&mut pt) });
    show_rss_context_menu(hwnd, pt.x, pt.y, false);
}

pub fn focus_library(hwnd: HWND) {
    if hwnd.0 == 0 {
        return;
    }
    crate::set_foreground_window_safe(hwnd);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 != 0 {
        select_first_root_if_needed(hwnd, hwnd_tree);
        crate::set_focus_safe(hwnd_tree);
    }
}

fn show_rss_context_menu(hwnd: HWND, x: i32, y: i32, use_hit_test: bool) {
    unsafe {
        let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
        if hwnd_tree.0 == 0 {
            return;
        }

        let mut rect = RECT::default();
        if use_hit_test
            && GetWindowRect(hwnd_tree, &mut rect).is_ok()
            && (x < rect.left || x > rect.right || y < rect.top || y > rect.bottom)
        {
            return;
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
                SendMessageW(
                    hwnd_tree,
                    TVM_SELECTITEM,
                    WPARAM(TVGN_CARET as usize),
                    LPARAM(hitem.0),
                );
                SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));
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
        let remove_entry_label = i18n::tr(language, "dictionary.remove");
        let retry_label = i18n::tr(language, "rss.context.retry_now");
        let reorder_label = i18n::tr(language, "rss.context.reorder");
        let reorder_up = i18n::tr(language, "rss.reorder.move_up");
        let reorder_down = i18n::tr(language, "rss.reorder.move_down");
        let reorder_top = i18n::tr(language, "rss.reorder.move_top");
        let reorder_bottom = i18n::tr(language, "rss.reorder.move_bottom");
        let reorder_position = i18n::tr(language, "rss.reorder.move_to_position");
        let sort_asc = i18n::tr(language, "rss.reorder.title_asc");
        let sort_desc = i18n::tr(language, "rss.reorder.title_desc");
        let sort_newest = i18n::tr(language, "rss.reorder.date_newest");
        let sort_oldest = i18n::tr(language, "rss.reorder.date_oldest");
        let open_label = i18n::tr(language, "rss.context.open_browser");
        let facebook_label = i18n::tr(language, "rss.context.share_facebook");
        let twitter_label = i18n::tr(language, "rss.context.share_twitter");
        let whatsapp_label = i18n::tr(language, "rss.context.share_whatsapp");
        let email_label = i18n::tr(language, "rss.context.share_email");
        let properties_label = i18n::tr(language, "context.properties");
        let undo_label = i18n::tr(language, "edit.undo")
            .split('\t')
            .next()
            .unwrap_or_default()
            .to_string();
        let has_undo = with_rss_state(hwnd, |s| !s.removed_history.is_empty()).unwrap_or(false);

        if let Ok(menu) = CreatePopupMenu()
            && menu.0 != 0
        {
            if is_source {
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_EDIT,
                    PCWSTR(to_wide(&edit_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_DELETE,
                    PCWSTR(to_wide(&delete_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_RETRY,
                    PCWSTR(to_wide(&retry_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_PROPERTIES,
                    PCWSTR(to_wide(&properties_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) {}
                let undo_flags = if has_undo {
                    MF_STRING
                } else {
                    MF_STRING | MF_GRAYED
                };
                if let Err(_e) = AppendMenuW(
                    menu,
                    undo_flags,
                    ID_CTX_UNDO_DELETE,
                    PCWSTR(to_wide(&undo_label).as_ptr()),
                ) {}
                if let Some(idx) = source_index {
                    let total = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| ps.settings.rss_sources.len())
                    })
                    .flatten()
                    .unwrap_or(0);
                    let at_top = idx == 0;
                    let at_bottom = total == 0 || idx + 1 >= total;
                    if let Ok(submenu) = CreatePopupMenu()
                        && submenu.0 != 0
                    {
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
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            up_flags,
                            ID_CTX_REORDER_UP,
                            PCWSTR(to_wide(&reorder_up).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            down_flags,
                            ID_CTX_REORDER_DOWN,
                            PCWSTR(to_wide(&reorder_down).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            up_flags,
                            ID_CTX_REORDER_TOP,
                            PCWSTR(to_wide(&reorder_top).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            down_flags,
                            ID_CTX_REORDER_BOTTOM,
                            PCWSTR(to_wide(&reorder_bottom).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            MF_STRING,
                            ID_CTX_REORDER_POSITION,
                            PCWSTR(to_wide(&reorder_position).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(submenu, MF_SEPARATOR, 0, PCWSTR::null()) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            MF_STRING,
                            ID_CTX_SORT_ASC,
                            PCWSTR(to_wide(&sort_asc).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            MF_STRING,
                            ID_CTX_SORT_DESC,
                            PCWSTR(to_wide(&sort_desc).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            MF_STRING,
                            ID_CTX_SORT_NEWEST,
                            PCWSTR(to_wide(&sort_newest).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            submenu,
                            MF_STRING,
                            ID_CTX_SORT_OLDEST,
                            PCWSTR(to_wide(&sort_oldest).as_ptr()),
                        ) {}
                        if let Err(_e) = AppendMenuW(
                            menu,
                            MF_POPUP,
                            submenu.0 as usize,
                            PCWSTR(to_wide(&reorder_label).as_ptr()),
                        ) {}
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
                if let Err(_e) = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_OPEN_BROWSER,
                    PCWSTR(to_wide(&open_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_FACEBOOK,
                    PCWSTR(to_wide(&facebook_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_TWITTER,
                    PCWSTR(to_wide(&twitter_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_WHATSAPP,
                    PCWSTR(to_wide(&whatsapp_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    flags,
                    ID_CTX_SHARE_EMAIL,
                    PCWSTR(to_wide(&email_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_PROPERTIES,
                    PCWSTR(to_wide(&properties_label).as_ptr()),
                ) {}
                if let Err(_e) = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) {}
                if let Err(_e) = AppendMenuW(
                    menu,
                    MF_STRING,
                    ID_CTX_DELETE,
                    PCWSTR(to_wide(&remove_entry_label).as_ptr()),
                ) {}
                let undo_flags = if has_undo {
                    MF_STRING
                } else {
                    MF_STRING | MF_GRAYED
                };
                if let Err(_e) = AppendMenuW(
                    menu,
                    undo_flags,
                    ID_CTX_UNDO_DELETE,
                    PCWSTR(to_wide(&undo_label).as_ptr()),
                ) {}
            }
            SetForegroundWindow(hwnd);
            if !TrackPopupMenu(
                menu,
                windows::Win32::UI::WindowsAndMessaging::TPM_RIGHTBUTTON,
                x,
                y,
                0,
                hwnd,
                None,
            )
            .as_bool()
            {
                crate::log_debug("TrackPopupMenu failed");
            }
            if let Err(e) = PostMessageW(hwnd, WM_NULL, WPARAM(0), LPARAM(0)) {
                crate::log_debug(&format!("Failed to post WM_NULL: {}", e));
            }
            crate::log_if_err!(DestroyMenu(menu));
        }
    }
}

unsafe extern "system" fn rss_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        crate::panic_guard::guard(
            "rss_wndproc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || rss_wndproc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn rss_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let parent = HWND((*cs).lpCreateParams as isize);

                let state = Box::new(RssWindowState {
                    parent,
                    hwnd_tree: HWND(0),
                    hwnd_import: HWND(0),
                    hwnd_export: HWND(0),
                    node_data: HashMap::new(),
                    pending_fetches: HashMap::new(),
                    source_items: HashMap::new(),
                    enter_guard: false,
                    add_guard: false,
                    reorder_dialog: HWND(0),
                    search_dialog: HWND(0),
                    pending_edit: None,
                    tree_proc: None,
                    last_selected: 0,
                    removed_history: Vec::new(),
                    suppress_tree_selection_events: false,
                    suppress_focus_restore_once: false,
                });
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);

                create_controls(hwnd);
                ensure_default_sources(parent);
                reload_tree(hwnd);
                let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                if hwnd_tree.0 != 0 {
                    select_first_root_if_needed(hwnd, hwnd_tree);
                    SetFocus(hwnd_tree);
                }

                // Start background check for new articles on all feeds
                start_background_unread_check(hwnd);

                LRESULT(0)
            }
            WM_DESTROY => {
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if parent.0 != 0 {
                    if with_state(parent, |s| s.rss_window = HWND(0)).is_none() {
                        crate::log_debug("Failed to reset rss_window state");
                    }
                    // Parent was never disabled; just bring it to front as a convenience.
                    // Only focus editor if not in player mode (audiobook)
                    if !crate::editor_manager::is_current_audiobook(parent) {
                        force_focus_editor_on_parent(parent);
                    }
                }
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut RssWindowState;
                if !ptr.is_null() {
                    let parent = (*ptr).parent;
                    let hwnd_tree = (*ptr).hwnd_tree;
                    if hwnd_tree.0 != 0
                        && let Some(proc) = (*ptr).tree_proc
                    {
                        let proc_ptr = proc as usize;
                        SetWindowLongPtrW(hwnd_tree, GWLP_WNDPROC, proc_ptr as isize);
                    }
                    if parent.0 != 0 {
                        force_focus_editor_on_parent(parent);
                    }
                    let _unused_box = Box::from_raw(ptr);
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                match id {
                    ID_BTN_CLOSE | 2 => {
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                        LRESULT(0)
                    }
                    ID_BTN_ADD => {
                        // Direct activation (mouse click, or Space/Enter if the button sends BN_CLICKED).
                        // Guard against the common case where Enter also generates IDOK (1),
                        // which would otherwise open the dialog twice and crash.
                        let already = with_rss_state(hwnd, |s| s.add_guard).unwrap_or(false);
                        if !already {
                            with_rss_state(hwnd, |s| s.add_guard = true);
                            if let Err(_e) =
                                PostMessageW(hwnd, WM_SHOW_ADD_DIALOG, WPARAM(0), LPARAM(0))
                            {
                                crate::log_debug(&format!("Error: {:?}", _e));
                            }
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
                    ID_BTN_SEARCH => {
                        show_rss_search_dialog(hwnd);
                        LRESULT(0)
                    }
                    ID_BTN_EXPORT => {
                        let language = with_rss_state(hwnd, |s| {
                            with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
                        })
                        .unwrap_or_default();
                        if let Some(path) = open_export_opml_dialog(hwnd, language) {
                            match export_sources_to_opml_file(hwnd, &path) {
                                Ok(count) => {
                                    if count > 0 {
                                        announce_rss_status(&i18n::tr(language, "rss.exported"));
                                    }
                                }
                                Err(err) => {
                                    let title = i18n::tr(language, "rss.window.title");
                                    let message = format!(
                                        "{}: {}",
                                        i18n::tr(language, "rss.export_failed"),
                                        err
                                    );
                                    MessageBoxW(
                                        hwnd,
                                        PCWSTR(to_wide(&message).as_ptr()),
                                        PCWSTR(to_wide(&title).as_ptr()),
                                        MB_OK | MB_ICONINFORMATION,
                                    );
                                }
                            }
                        }
                        LRESULT(0)
                    }
                    1 => {
                        // IDOK (Enter key often triggers this generic command)
                        let focus = GetFocus();
                        let btn_add = GetDlgItem(hwnd, ID_BTN_ADD as i32);
                        let btn_search = GetDlgItem(hwnd, ID_BTN_SEARCH as i32);
                        let btn_import = GetDlgItem(hwnd, ID_BTN_IMPORT as i32);
                        let btn_export = GetDlgItem(hwnd, ID_BTN_EXPORT as i32);
                        let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));

                        if focus == btn_add {
                            let already = with_rss_state(hwnd, |s| s.add_guard).unwrap_or(false);
                            if !already {
                                with_rss_state(hwnd, |s| s.add_guard = true);
                                if let Err(_e) =
                                    PostMessageW(hwnd, WM_SHOW_ADD_DIALOG, WPARAM(0), LPARAM(0))
                                {
                                    crate::log_debug(&format!("Error: {:?}", _e));
                                }
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

                        if focus == btn_search {
                            show_rss_search_dialog(hwnd);
                            return LRESULT(0);
                        }

                        if focus == btn_export {
                            let language = with_rss_state(hwnd, |s| {
                                with_state(s.parent, |ps| ps.settings.language).unwrap_or_default()
                            })
                            .unwrap_or_default();
                            if let Some(path) = open_export_opml_dialog(hwnd, language) {
                                match export_sources_to_opml_file(hwnd, &path) {
                                    Ok(count) => {
                                        if count > 0 {
                                            announce_rss_status(&i18n::tr(
                                                language,
                                                "rss.exported",
                                            ));
                                        }
                                    }
                                    Err(err) => {
                                        let title = i18n::tr(language, "rss.window.title");
                                        let message = format!(
                                            "{}: {}",
                                            i18n::tr(language, "rss.export_failed"),
                                            err
                                        );
                                        MessageBoxW(
                                            hwnd,
                                            PCWSTR(to_wide(&message).as_ptr()),
                                            PCWSTR(to_wide(&title).as_ptr()),
                                            MB_OK | MB_ICONINFORMATION,
                                        );
                                    }
                                }
                            }
                            return LRESULT(0);
                        }

                        if focus == hwnd_tree {
                            let already = with_rss_state(hwnd, |s| s.enter_guard).unwrap_or(false);
                            if !already {
                                let shift_down =
                                    GetKeyState(VK_SHIFT.0 as i32) & 0x8000u16 as i16 != 0;
                                with_rss_state(hwnd, |s| s.enter_guard = true);
                                if let Err(_e) =
                                    PostMessageW(hwnd, WM_CLEAR_ENTER_GUARD, WPARAM(0), LPARAM(0))
                                {
                                    crate::log_debug(&format!("Error: {:?}", _e));
                                }
                                handle_enter_action(hwnd, shift_down);
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
                    ID_CTX_UNDO_DELETE => {
                        undo_last_delete(hwnd);
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
                    ID_CTX_SORT_ASC => {
                        handle_sort_action(hwnd, crate::settings::SortOrder::TitleAsc);
                        LRESULT(0)
                    }
                    ID_CTX_SORT_DESC => {
                        handle_sort_action(hwnd, crate::settings::SortOrder::TitleDesc);
                        LRESULT(0)
                    }
                    ID_CTX_SORT_NEWEST => {
                        handle_sort_action(hwnd, crate::settings::SortOrder::DateNewest);
                        LRESULT(0)
                    }
                    ID_CTX_SORT_OLDEST => {
                        handle_sort_action(hwnd, crate::settings::SortOrder::DateOldest);
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
                    ID_CTX_SHARE_EMAIL => {
                        handle_article_action(hwnd, ArticleAction::ShareEmail);
                        LRESULT(0)
                    }
                    ID_CTX_PROPERTIES => {
                        show_selected_properties(hwnd);
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
                                last_updated: None,
                                removed_item_keys: Vec::new(),
                                read_item_keys: Vec::new(),
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
                            let action_val = (*pnmtv).action.0;
                            if (action_val & TVE_EXPAND.0) != 0 {
                                let hitem = (*pnmtv).itemNew.hItem;
                                handle_expand(hwnd, hitem);
                            }
                            LRESULT(0)
                        }
                        TVN_SELCHANGEDW => {
                            if with_rss_state(hwnd, |s| s.suppress_tree_selection_events)
                                .unwrap_or(false)
                            {
                                return LRESULT(0);
                            }
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
                                if GetKeyState(VK_MENU.0 as i32) < 0 {
                                    show_selected_properties(hwnd);
                                    return LRESULT(1);
                                }
                                let shift_down =
                                    GetKeyState(VK_SHIFT.0 as i32) & 0x8000u16 as i16 != 0;
                                with_rss_state(hwnd, |s| s.enter_guard = true);
                                if let Err(_e) =
                                    PostMessageW(hwnd, WM_CLEAR_ENTER_GUARD, WPARAM(0), LPARAM(0))
                                {
                                    crate::log_debug(&format!("Error: {:?}", _e));
                                }
                                handle_enter_action(hwnd, shift_down);
                                LRESULT(1)
                            } else if (*ptvkd).wVKey
                                == windows::Win32::UI::Input::KeyboardAndMouse::VK_DELETE.0
                            {
                                handle_delete(hwnd);
                                LRESULT(1)
                            } else if (*ptvkd).wVKey == 'Z' as u16
                                && GetKeyState(VK_CONTROL.0 as i32) < 0
                            {
                                undo_last_delete(hwnd);
                                LRESULT(1)
                            } else if (*ptvkd).wVKey == VK_F10.0
                                && GetKeyState(VK_SHIFT.0 as i32) < 0
                            {
                                let hwnd_tree =
                                    with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                                if hwnd_tree.0 != 0
                                    && let Err(_e) = PostMessageW(
                                        hwnd,
                                        WM_CONTEXTMENU,
                                        WPARAM(hwnd_tree.0 as usize),
                                        LPARAM(-1),
                                    )
                                {}
                                LRESULT(1)
                            } else if (*ptvkd).wVKey == VK_APPS.0 {
                                let hwnd_tree =
                                    with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                                if hwnd_tree.0 != 0
                                    && let Err(_e) = PostMessageW(
                                        hwnd,
                                        WM_CONTEXTMENU,
                                        WPARAM(hwnd_tree.0 as usize),
                                        LPARAM(-1),
                                    )
                                {}
                                LRESULT(1)
                            } else if (*ptvkd).wVKey == VK_ESCAPE.0 {
                                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                                LRESULT(1)
                            } else {
                                LRESULT(0)
                            }
                        }
                        NM_RCLICK => {
                            let mut pt = POINT::default();
                            crate::log_if_err!(GetCursorPos(&mut pt));
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
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
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
                    if let Err(_e) = PostMessageW(
                        hwnd,
                        WM_CONTEXTMENU,
                        WPARAM(hwnd_tree.0 as usize),
                        LPARAM(-1),
                    ) {}
                    return LRESULT(0);
                }
                if key == 'Z' as u32 && GetKeyState(VK_CONTROL.0 as i32) < 0 {
                    undo_last_delete(hwnd);
                    return LRESULT(0);
                }
                if key == 'C' as u32 && GetKeyState(VK_CONTROL.0 as i32) < 0 {
                    ignore_bool(handle_rss_quick_copy(hwnd));
                    return LRESULT(0);
                }
                if key == u32::from(VK_RETURN.0) && GetKeyState(VK_MENU.0 as i32) < 0 {
                    show_selected_properties(hwnd);
                    return LRESULT(0);
                }
                if key == u32::from(VK_ESCAPE.0) {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
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
                    crate::log_if_err!(GetCursorPos(&mut pt));
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
            WM_RSS_MARK_ITEM_READ_UI => {
                let ptr = lparam.0 as *mut MarkItemReadUiMessage;
                let msg_data = *Box::from_raw(ptr);
                let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                if hwnd_tree.0 == 0 {
                    return LRESULT(0);
                }
                let hitem = windows::Win32::UI::Controls::HTREEITEM(msg_data.hitem);
                let item = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
                    Some(NodeData::Item(item)) => Some(item.clone()),
                    _ => None,
                })
                .flatten();
                if let Some(item) = item
                    && rss_item_key(&item) == msg_data.item_key
                {
                    let (
                        language,
                        announce_unread,
                        unread_label_position,
                        rss_date_mode,
                        rss_time_mode,
                    ) = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| {
                            (
                                ps.settings.language,
                                ps.settings.announce_unread_rss_podcast_items,
                                ps.settings.rss_podcast_unread_label_position,
                                ps.settings.rss_articles_date_display,
                                ps.settings.rss_articles_time_display,
                            )
                        })
                        .unwrap_or((
                            crate::settings::Language::English,
                            true,
                            crate::settings::RssPodcastUnreadLabelPosition::Before,
                            ListDateDisplayMode::Always,
                            ListTimeDisplayMode::Always,
                        ))
                    })
                    .unwrap_or((
                        crate::settings::Language::English,
                        true,
                        crate::settings::RssPodcastUnreadLabelPosition::Before,
                        ListDateDisplayMode::Always,
                        ListTimeDisplayMode::Always,
                    ));
                    let day_counts = with_rss_state(hwnd, |s| {
                        s.source_items
                            .values()
                            .find(|state| {
                                state
                                    .items
                                    .iter()
                                    .any(|x| rss_item_key(x) == msg_data.item_key)
                            })
                            .map(|state| build_day_counts(&state.items))
                            .unwrap_or_default()
                    })
                    .unwrap_or_default();
                    let same_day = has_multiple_items_same_day(item.pub_date, &day_counts);
                    let title_ctx = RssItemTitleContext {
                        language,
                        announce_unread,
                        unread_label_position,
                        date_mode: rss_date_mode,
                        time_mode: rss_time_mode,
                    };
                    let updated = rss_item_display_title(
                        &item.title,
                        false,
                        item.pub_date,
                        same_day,
                        title_ctx,
                    );
                    let text = to_wide(&updated);
                    let mut tv_item = TVITEMW {
                        mask: TVIF_TEXT,
                        hItem: hitem,
                        pszText: windows::core::PWSTR(text.as_ptr() as *mut _),
                        cchTextMax: text.len() as i32,
                        ..Default::default()
                    };
                    SendMessageW(
                        hwnd_tree,
                        TVM_SETITEMW,
                        WPARAM(0),
                        LPARAM(&mut tv_item as *mut _ as isize),
                    );
                }
                LRESULT(0)
            }
            WM_RSS_SELECT_SOURCE_DELAYED => {
                select_source_by_index(hwnd, wparam.0);
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
                if wparam.0 == ADD_GUARD_TIMER_ID {
                    with_rss_state(hwnd, |s| s.add_guard = false);
                    if let Err(e) = KillTimer(hwnd, ADD_GUARD_TIMER_ID) {
                        crate::log_debug(&format!("Failed to kill ADD_GUARD_TIMER: {}", e));
                    }
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }

            WM_RSS_IMPORT_COMPLETE => {
                let ptr = lparam.0 as *mut ImportResult;
                let res = Box::from_raw(ptr);

                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                let language =
                    with_state(parent, |state| state.settings.language).unwrap_or_default();
                let rss_title = i18n::tr(language, "rss.temp_title");
                let hwnd_edit = editor_manager::get_or_create_rss_document(parent, &rss_title);

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
                if let Err(_e) = PostMessageW(hwnd, WM_CLEAR_ADD_GUARD, WPARAM(0), LPARAM(0)) {
                    crate::log_debug(&format!("Error: {:?}", _e));
                }
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
                    let suppress_restore =
                        with_rss_state(hwnd, |s| mem::take(&mut s.suppress_focus_restore_once))
                            .unwrap_or(false);
                    if suppress_restore {
                        return LRESULT(0);
                    }
                    let last = with_rss_state(hwnd, |s| s.last_selected).unwrap_or(0);
                    if last != 0 {
                        SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(last),
                        );
                        SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(last));
                    }
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

unsafe extern "system" fn reorder_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "reorder_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || reorder_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn reorder_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
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
                SendMessageW(edit, EM_LIMITTEXT, WPARAM(6), LPARAM(0));
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
                let proc_ptr = reorder_control_subclass_proc as *const () as usize;
                for control in [edit, ok, cancel] {
                    let prev = SetWindowLongPtrW(control, GWLP_WNDPROC, proc_ptr as isize);
                    SetWindowLongPtrW(control, GWLP_USERDATA, prev);
                }
                SetFocus(edit);
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
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
                        let text = String::from_utf16_lossy(&buf[..len as usize]);
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
                        ) && new_index != init.source_index
                        {
                            let template = i18n::tr(language, "rss.reorder.moved_position");
                            let message = template.replace("{x}", &(new_index + 1).to_string());
                            announce_rss_status(&message);
                        }
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                        LRESULT(0)
                    }
                    REORDER_CANCEL_ID | 2 => {
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
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
}

fn with_rss_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut RssWindowState) -> R,
{
    let ptr = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA) as *mut RssWindowState;
    if ptr.is_null() {
        None
    } else {
        Some(unsafe { f(&mut *ptr) })
    }
}

unsafe extern "system" fn rss_tree_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "rss_tree_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || rss_tree_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn rss_tree_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        if msg == windows::Win32::UI::WindowsAndMessaging::WM_CHAR
            && wparam.0 as u32 == 3
            && GetKeyState(VK_CONTROL.0 as i32) < 0
        {
            return LRESULT(0);
        }
        if msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN {
            let key = wparam.0 as u32;
            if key == 'C' as u32 && GetKeyState(VK_CONTROL.0 as i32) < 0 {
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    ignore_bool(handle_rss_quick_copy(parent));
                    return LRESULT(0);
                }
            }
            if key == u32::from(VK_RETURN.0) && GetKeyState(VK_MENU.0 as i32) < 0 {
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    show_selected_properties(parent);
                    return LRESULT(0);
                }
            }
            if key == u32::from(VK_APPS.0)
                || (key == u32::from(VK_F10.0) && GetKeyState(VK_SHIFT.0 as i32) < 0)
            {
                let parent = GetParent(hwnd);
                if parent.0 != 0 {
                    if let Err(_e) =
                        PostMessageW(parent, WM_CONTEXTMENU, WPARAM(hwnd.0 as usize), LPARAM(-1))
                    {
                        crate::log_debug(&format!("Error: {:?}", _e));
                    }
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
}

fn create_controls(hwnd: HWND) {
    unsafe {
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
                    windows::Win32::UI::Controls::TVS_HASLINES
                        | windows::Win32::UI::Controls::TVS_HASBUTTONS
                        | windows::Win32::UI::Controls::TVS_LINESATROOT
                        | windows::Win32::UI::Controls::TVS_SHOWSELALWAYS,
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
            let proc_ptr = rss_tree_wndproc as *const () as usize;
            let old = SetWindowLongPtrW(hwnd_tree, GWLP_WNDPROC, proc_ptr as isize);
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
            90,
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
            105,
            520,
            90,
            30,
            hwnd,
            HMENU(ID_BTN_IMPORT as isize),
            hinstance,
            None,
        );

        let export_label = i18n::tr(language, "rss.tree.export_opml");
        let export_label = if export_label == "rss.tree.export_opml" {
            "Export OPML...".to_string()
        } else {
            export_label
        };
        let hwnd_export = CreateWindowExW(
            Default::default(),
            WC_BUTTON,
            PCWSTR(to_wide(&export_label).as_ptr()),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            200,
            520,
            90,
            30,
            hwnd,
            HMENU(ID_BTN_EXPORT as isize),
            hinstance,
            None,
        );

        let hwnd_close = CreateWindowExW(
            Default::default(),
            WC_BUTTON,
            PCWSTR(to_wide(&i18n::tr(language, "rss.tree.close")).as_ptr()),
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            295,
            520,
            80,
            30,
            hwnd,
            HMENU(ID_BTN_CLOSE as isize),
            hinstance,
            None,
        );

        with_rss_state(hwnd, |s| {
            s.hwnd_tree = hwnd_tree;
            s.hwnd_import = hwnd_import;
            s.hwnd_export = hwnd_export;
        });

        let hfont = with_rss_state(hwnd, |s| {
            with_state(s.parent, |ps| ps.hfont).unwrap_or(HFONT(0))
        })
        .unwrap_or(HFONT(0));
        if hfont.0 != 0 {
            SendMessageW(hwnd_tree, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            SendMessageW(hwnd_add, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            SendMessageW(hwnd_import, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            SendMessageW(hwnd_export, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
            SendMessageW(hwnd_close, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }

        SetFocus(hwnd_tree);
    }
}

fn reload_tree(hwnd: HWND) {
    let (hwnd_tree, sources, language, announce_unread, unread_label_position) =
        match with_rss_state(hwnd, |s| {
            (
                s.hwnd_tree,
                { with_state(s.parent, |ps| ps.settings.rss_sources.clone()) },
                { with_state(s.parent, |ps| ps.settings.language) },
                { with_state(s.parent, |ps| ps.settings.announce_unread_rss_podcast_items) },
                { with_state(s.parent, |ps| ps.settings.rss_podcast_unread_label_position) },
            )
        }) {
            Some((
                t,
                Some(src),
                Some(language),
                Some(announce_unread),
                Some(unread_label_position),
            )) => (t, src, language, announce_unread, unread_label_position),
            _ => return,
        };

    crate::send_message_w_safe(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(TVI_ROOT.0));

    with_rss_state(hwnd, |s| {
        s.node_data.clear();
        s.source_items.clear();
    });

    for (i, source) in sources.into_iter().enumerate() {
        let title = to_wide(&rss_source_display_title(
            &source,
            language,
            announce_unread,
            unread_label_position,
        ));
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
        let hitem = crate::send_message_w_safe(
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

fn schedule_delayed_source_select(hwnd: HWND, source_index: usize) {
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_secs(2));
        if let Err(e) = crate::post_message_w_safe(
            hwnd,
            WM_RSS_SELECT_SOURCE_DELAYED,
            WPARAM(source_index),
            LPARAM(0),
        ) {
            crate::log_debug(&format!(
                "Failed to post delayed RSS source selection: {}",
                e
            ));
        }
    });
}

fn select_source_by_index(hwnd: HWND, source_index: usize) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let target_hitem = with_rss_state(hwnd, |s| {
        s.node_data.iter().find_map(|(&h, node)| match node {
            NodeData::Source(i) if *i == source_index => {
                Some(windows::Win32::UI::Controls::HTREEITEM(h))
            }
            _ => None,
        })
    })
    .flatten();
    if let Some(target) = target_hitem {
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_SELECTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(target.0),
            );
            SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(target.0));
        }
        if crate::get_focus_safe() != hwnd_tree {
            with_rss_state(hwnd, |s| s.suppress_focus_restore_once = true);
            crate::set_focus_safe(hwnd_tree);
        }
    }
}

fn update_source_tree_title(
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
    crate::send_message_w_safe(
        hwnd_tree,
        TVM_SETITEMW,
        WPARAM(0),
        LPARAM(&mut tvi as *mut _ as isize),
    );
}

fn set_source_unread(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM, unread: bool) {
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
        let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
        let title_opt = {
            with_state(parent, |ps| {
                if let Some(src) = ps.settings.rss_sources.get_mut(idx) {
                    let mut changed = false;
                    if src.unread != unread {
                        src.unread = unread;
                        changed = true;
                    }
                    // Update last_seen_guid when marking as read
                    if let Some(ref key) = first_item_key
                        && src.last_seen_guid.as_ref() != Some(key)
                    {
                        src.last_seen_guid = Some(key.clone());
                        changed = true;
                    }
                    if changed {
                        let title = rss_source_display_title(
                            src,
                            language,
                            ps.settings.announce_unread_rss_podcast_items,
                            ps.settings.rss_podcast_unread_label_position,
                        );
                        crate::settings::save_settings(ps.settings.clone());
                        return Some(title);
                    }
                }
                None
            })
        }
        .flatten();
        (hwnd_tree, title_opt)
    })
    .unwrap_or((HWND(0), None));
    if let Some(title) = title_opt {
        update_source_tree_title(hwnd_tree, hitem, &title);
    }
}

fn handle_expand(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM) {
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
            {
                with_state(s.parent, |ps| {
                    ps.settings
                        .rss_sources
                        .get(*idx)
                        .map(|src| (src.url.clone(), src.kind.clone(), src.cache.clone(), true))
                })
            }
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
        let first_child = crate::send_message_w_safe(
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
            crate::send_message_w_safe(
                hwnd_tree,
                TVM_INSERTITEMW,
                WPARAM(0),
                LPARAM(&mut tvis_loading as *mut _ as isize),
            );
        }
        // Force visual expansion now.
        unsafe {
            SendMessageW(
                hwnd_tree,
                TVM_EXPAND,
                WPARAM(TVE_EXPAND.0 as usize),
                LPARAM(hitem.0),
            );
            SendMessageW(hwnd_tree, TVM_ENSUREVISIBLE, WPARAM(0), LPARAM(hitem.0));
        }
    }

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let fetch_config = if parent.0 != 0 {
        rss_fetch_config(parent)
    } else {
        rss::RssFetchConfig::default()
    };
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    // UI: "Refresh feeds" should trigger this fetch for the selected source.
    std::thread::spawn(move || {
        let rt = match tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
        {
            Ok(rt) => rt,
            Err(e) => {
                crate::log_debug(&format!("Failed to build tokio runtime: {}", e));
                return;
            }
        };
        let res = rt.block_on(rss::fetch_and_parse(
            &url_clone,
            source_kind,
            cache,
            fetch_config,
            false,
            language,
        ));
        let msg = Box::new(FetchResult {
            hitem: hitem.0,
            result: res,
        });
        if let Err(_e) = crate::post_message_w_safe(
            hwnd,
            WM_RSS_FETCH_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        ) {}
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
fn start_background_unread_check(hwnd: HWND) {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let sources: Vec<(
        usize,
        String,
        rss::RssSourceType,
        rss::RssFeedCache,
        HashSet<String>,
    )> = {
        with_state(parent, |ps| {
            ps.settings
                .rss_sources
                .iter()
                .enumerate()
                .map(|(i, src)| {
                    (
                        i,
                        src.url.clone(),
                        src.kind.clone(),
                        src.cache.clone(),
                        src.removed_item_keys.iter().cloned().collect(),
                    )
                })
                .collect()
        })
    }
    .unwrap_or_default();

    if sources.is_empty() {
        return;
    }

    let fetch_config = rss_fetch_config(parent);
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    ensure_rss_http(parent);

    let hwnd_raw = hwnd.0;
    std::thread::spawn(move || {
        let rt = match tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
        {
            Ok(rt) => rt,
            Err(e) => {
                crate::log_debug(&format!("Failed to build tokio runtime: {}", e));
                return;
            }
        };

        rt.block_on(async {
            // Process feeds concurrently but not all at once (limit concurrency)
            let semaphore = std::sync::Arc::new(tokio::sync::Semaphore::new(4));
            let mut handles = Vec::new();

            for (idx, url, kind, cache, removed_keys) in sources {
                let sem = semaphore.clone();
                let cfg = fetch_config;
                let hwnd_val = hwnd_raw;

                let handle = tokio::spawn(async move {
                    let _permit = sem.acquire().await.ok()?;
                    let result =
                        rss::fetch_and_parse(&url, kind, cache, cfg, false, language).await;
                    if let Ok(outcome) = result {
                        let newest_key = select_newest_item_key(&outcome.items, &removed_keys);
                        let msg = Box::new(BackgroundCheckResult {
                            source_idx: idx,
                            newest_item_key: newest_key,
                        });
                        if let Err(e) = crate::post_message_w_safe(
                            HWND(hwnd_val),
                            WM_RSS_BACKGROUND_CHECK_COMPLETE,
                            WPARAM(0),
                            LPARAM(Box::into_raw(msg) as isize),
                        ) {
                            crate::log_debug(&format!(
                                "Failed to post WM_RSS_BACKGROUND_CHECK_COMPLETE: {}",
                                e
                            ));
                        }
                    }
                    Some(())
                });
                handles.push(handle);
            }

            for h in handles {
                if let Err(e) = h.await {
                    crate::log_debug(&format!("Background check handle await failed: {}", e));
                }
            }
        });
    });
}

/// Process background check result - update unread state if new articles detected
fn process_background_check_result(hwnd: HWND, res: BackgroundCheckResult) {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if parent.0 == 0 {
        return;
    }

    let Some(newest_key) = res.newest_item_key else {
        return;
    };

    // Check if this is a new article compared to last_seen_guid
    let should_mark_unread = {
        with_state(parent, |ps| {
            ps.settings
                .rss_sources
                .get(res.source_idx)
                .map(|src| match &src.last_seen_guid {
                    Some(last_seen) => last_seen != &newest_key,
                    None => true, // Never seen before
                })
                .unwrap_or(false)
        })
    }
    .unwrap_or(false);

    if should_mark_unread {
        // Find the tree item for this source and mark it unread
        let hitem_opt = with_rss_state(hwnd, |s| {
            for (&h, node) in &s.node_data {
                if let NodeData::Source(idx) = node
                    && *idx == res.source_idx
                {
                    return Some(windows::Win32::UI::Controls::HTREEITEM(h));
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

fn process_fetch_result(hwnd: HWND, res: FetchResult) {
    unsafe {
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
                                let looks_auto =
                                    src.title.trim().is_empty() || src.title == src.url;
                                if !src.user_title
                                    && !keep_default_title
                                    && looks_auto
                                    && !outcome.title.is_empty()
                                {
                                    src.title = outcome.title.clone();
                                }
                                final_title = rss_source_display_title(
                                    src,
                                    lang,
                                    ps.settings.announce_unread_rss_podcast_items,
                                    ps.settings.rss_podcast_unread_label_position,
                                );
                                if src.kind != outcome.kind {
                                    src.kind = outcome.kind;
                                }
                                src.cache = outcome.cache.clone();
                                let max_ts = outcome.items.iter().filter_map(|i| i.pub_date).max();
                                if let Some(ts) = max_ts {
                                    src.last_updated = Some(ts);
                                }
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
                            SendMessageW(
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
                                let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
                                let text = String::from_utf16_lossy(&buf[..len]);
                                if text.trim() == "Loading..." {
                                    SendMessageW(
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
                let existing = with_rss_state(hwnd, |s| s.source_items.contains_key(&hitem.0))
                    .unwrap_or(false);

                let removed_keys = source_removed_keys_for_tree_item(hwnd, hitem);

                if existing {
                    with_rss_state(hwnd, |s| {
                        let Some(state) = s.source_items.get_mut(&hitem.0) else {
                            return;
                        };
                        loaded_before = state.loaded;
                        total_before = state.items.len();
                        let mut seen: HashSet<String> =
                            state.items.iter().map(rss_item_key).collect();
                        for item in outcome.items {
                            let key = rss_item_key(&item);
                            if removed_keys.contains(&key) {
                                continue;
                            }
                            if seen.insert(key) {
                                state.items.push(item);
                                appended += 1;
                            }
                        }
                        // Keep a consistent chronological order for all RSS feeds.
                        // Google News often requires this, but standard feeds can also arrive unsorted.
                        sort_items_by_date_desc(&mut state.items);
                    });
                    if appended > 0 {
                        log_debug(&format!(
                            "rss_ui_batch start source={} append_count={} loaded_before={} total_before={}",
                            hitem.0, appended, loaded_before, total_before
                        ));
                        if loaded_before >= total_before {
                            let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                            let (_initial_count, next_count) = rss_page_sizes(parent);
                            load_more_items(hwnd, hitem, next_count);
                        }
                        log_debug(&format!(
                            "rss_ui_batch end source={} appended={}",
                            hitem.0, appended
                        ));
                        set_source_unread(hwnd, hitem, true);
                    } else {
                        // Feed opened/refreshed with no newly appended items: clear "new items" flag.
                        // This keeps source-level status aligned with real incoming updates.
                        set_source_unread(hwnd, hitem, false);
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

                    let saved_read_item_keys: HashSet<String> = with_rss_state(hwnd, |s| {
                        let source_index = match s.node_data.get(&hitem.0) {
                            Some(NodeData::Source(index)) => Some(*index),
                            _ => None,
                        };
                        with_state(s.parent, |ps| {
                            source_index
                                .and_then(|index| ps.settings.rss_sources.get(index))
                                .map(|src| src.read_item_keys.iter().cloned().collect())
                        })
                        .flatten()
                        .unwrap_or_default()
                    })
                    .unwrap_or_default();

                    with_rss_state(hwnd, |s| {
                        let mut items: Vec<RssItem> = outcome
                            .items
                            .into_iter()
                            .filter(|item| !removed_keys.contains(&rss_item_key(item)))
                            .collect();
                        // Keep a consistent chronological order for all RSS feeds.
                        sort_items_by_date_desc(&mut items);
                        s.source_items.insert(
                            hitem.0,
                            SourceItemsState {
                                items,
                                loaded: 0,
                                read_item_keys: saved_read_item_keys,
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
                prune_persisted_read_keys_for_source(hwnd, hitem);
            }
            Err(e) => {
                let (message, cache) = match e {
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

                let has_items = with_rss_state(hwnd, |s| s.source_items.contains_key(&hitem.0))
                    .unwrap_or(false);
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
}

fn handle_selection_changed(hwnd: HWND, hitem: windows::Win32::UI::Controls::HTREEITEM) {
    if hitem.0 == 0 {
        return;
    }
    with_rss_state(hwnd, |s| s.last_selected = hitem.0);
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return;
    }
    let parent = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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
        crate::send_message_w_safe(
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
            crate::send_message_w_safe(
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
        load_more_items(hwnd, parent, next_count);
    }
}

fn load_more_items(
    hwnd: HWND,
    hitem: windows::Win32::UI::Controls::HTREEITEM,
    batch: usize,
) -> usize {
    // UI: "Load more titles" can call this to append the next page locally.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return 0;
    }
    let (language, announce_unread, unread_label_position, rss_date_mode, rss_time_mode) =
        with_rss_state(hwnd, |s| {
            {
                with_state(s.parent, |ps| {
                    (
                        ps.settings.language,
                        ps.settings.announce_unread_rss_podcast_items,
                        ps.settings.rss_podcast_unread_label_position,
                        ps.settings.rss_articles_date_display,
                        ps.settings.rss_articles_time_display,
                    )
                })
            }
            .unwrap_or((
                crate::settings::Language::English,
                true,
                crate::settings::RssPodcastUnreadLabelPosition::Before,
                ListDateDisplayMode::Always,
                ListTimeDisplayMode::Always,
            ))
        })
        .unwrap_or((
            crate::settings::Language::English,
            true,
            crate::settings::RssPodcastUnreadLabelPosition::Before,
            ListDateDisplayMode::Always,
            ListTimeDisplayMode::Always,
        ));

    crate::send_message_w_safe(hwnd_tree, WM_SETREDRAW, WPARAM(0), LPARAM(0));
    let (inserted, loaded_after, total_after) = with_rss_state(hwnd, |s| {
        let Some(state) = s.source_items.get_mut(&hitem.0) else {
            return (0usize, 0usize, 0usize);
        };
        if state.loaded >= state.items.len() {
            return (0usize, state.loaded, state.items.len());
        }
        let mut inserted = 0usize;
        let mut idx = state.loaded;
        let day_counts = build_day_counts(&state.items);
        while idx < state.items.len() && inserted < batch {
            let item = &state.items[idx];
            idx += 1;
            if item.title.trim().is_empty() {
                continue;
            }
            let item_unread = !state.read_item_keys.contains(&rss_item_key(item));
            let title_ctx = RssItemTitleContext {
                language,
                announce_unread,
                unread_label_position,
                date_mode: rss_date_mode,
                time_mode: rss_time_mode,
            };
            let display_title = rss_item_display_title(
                &item.title,
                item_unread,
                item.pub_date,
                has_multiple_items_same_day(item.pub_date, &day_counts),
                title_ctx,
            );
            let text = to_wide(&display_title);
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
            let hchild = crate::send_message_w_safe(
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
    crate::send_message_w_safe(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
    if inserted > 0 {
        log_debug(&format!(
            "rss_ui_batch append source={} inserted={} loaded={} total={}",
            hitem.0, inserted, loaded_after, total_after
        ));
    }
    inserted
}

fn handle_enter_action(hwnd: HWND, open_in_browser: bool) {
    // UI: Enter imports the article, Shift+Enter opens it in the browser.
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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
        let item_key = rss_item_key(&item);
        with_rss_state(hwnd, |s| {
            let parent = windows::Win32::UI::Controls::HTREEITEM(
                crate::send_message_w_safe(
                    s.hwnd_tree,
                    TVM_GETNEXTITEM,
                    WPARAM(TVGN_PARENT as usize),
                    LPARAM(hitem.0),
                )
                .0,
            );
            if parent.0 != 0
                && let Some(state) = s.source_items.get_mut(&parent.0)
            {
                state.read_item_keys.insert(item_key.clone());
            }
            if parent.0 != 0
                && let Some(NodeData::Source(source_index)) = s.node_data.get(&parent.0)
            {
                {
                    with_state(s.parent, |ps| {
                        if let Some(src) = ps.settings.rss_sources.get_mut(*source_index)
                            && !src.read_item_keys.iter().any(|k| k == &item_key)
                        {
                            src.read_item_keys.push(item_key.clone());
                            const MAX_PERSISTED_READ_KEYS: usize = 5000;
                            if src.read_item_keys.len() > MAX_PERSISTED_READ_KEYS {
                                let overflow = src.read_item_keys.len() - MAX_PERSISTED_READ_KEYS;
                                src.read_item_keys.drain(0..overflow);
                            }
                            crate::settings::save_settings(ps.settings.clone());
                        }
                    });
                }
            }
        });
        let delayed_key = rss_item_key(&item);
        let delayed_hitem = hitem.0;
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_secs(2));
            let payload = Box::new(MarkItemReadUiMessage {
                hitem: delayed_hitem,
                item_key: delayed_key,
            });
            let payload_ptr = Box::into_raw(payload);
            if let Err(e) = crate::post_message_w_safe(
                hwnd,
                WM_RSS_MARK_ITEM_READ_UI,
                WPARAM(0),
                LPARAM(payload_ptr as isize),
            ) {
                let _payload_owner = unsafe { Box::from_raw(payload_ptr) };
                crate::log_debug(&format!("Failed to post WM_RSS_MARK_ITEM_READ_UI: {}", e));
            }
        });
        if open_in_browser {
            handle_article_action(hwnd, ArticleAction::OpenInBrowser);
        } else {
            import_item(hwnd, item);
        }
    } else {
        crate::send_message_w_safe(
            hwnd_tree,
            TVM_EXPAND,
            WPARAM(TVE_EXPAND.0 as usize),
            LPARAM(hitem.0),
        );
    }
}

fn handle_delete(hwnd: HWND) {
    unsafe {
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

        let selected_node = with_rss_state(hwnd, |s| s.node_data.get(&hitem.0).cloned()).flatten();

        if let Some(NodeData::Source(idx)) = selected_node {
            let source_info = with_rss_state(hwnd, |s| {
                with_state(s.parent, |ps| ps.settings.rss_sources.get(idx).cloned()).flatten()
            })
            .flatten();
            let (title, url) = source_info
                .as_ref()
                .map(|src| (src.title.clone(), src.url.clone()))
                .unwrap_or_default();

            // Localize message and title
            let (language, require_confirm) = with_rss_state(hwnd, |s| {
                with_state(s.parent, |ps| {
                    (
                        ps.settings.language,
                        matches!(
                            ps.settings.rss_delete_confirm_mode,
                            crate::settings::RssDeleteConfirmMode::Feed
                                | crate::settings::RssDeleteConfirmMode::Both
                        ),
                    )
                })
                .unwrap_or((crate::settings::Language::default(), true))
            })
            .unwrap_or((crate::settings::Language::default(), true));
            let msg_template = i18n::tr(language, "rss.delete_confirm");
            let msg_text = msg_template.replace("{title}", &title);
            let caption = i18n::tr(language, "rss.delete_title");

            let confirmed = if require_confirm {
                MessageBoxW(
                    hwnd,
                    PCWSTR(to_wide(&msg_text).as_ptr()),
                    PCWSTR(to_wide(&caption).as_ptr()),
                    MB_YESNO | MB_ICONQUESTION,
                ) == IDYES
            } else {
                true
            };
            if confirmed {
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                let mut default_removed_key_added: Option<String> = None;
                let mut delayed_target_source_idx: Option<usize> = None;
                with_state(parent, |ps| {
                    if matches!(
                        language,
                        crate::settings::Language::English
                            | crate::settings::Language::Swedish
                            | crate::settings::Language::Italian
                            | crate::settings::Language::Spanish
                            | crate::settings::Language::Portuguese
                            | crate::settings::Language::Vietnamese
                            | crate::settings::Language::Czech
                            | crate::settings::Language::Polish
                            | crate::settings::Language::French
                            | crate::settings::Language::Serbian
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
                                    crate::settings::Language::Ukrainian
                                    | crate::settings::Language::English
                                    | crate::settings::Language::Lithuanian
                                    | crate::settings::Language::Chinese => {
                                        &mut ps.settings.rss_removed_default_en
                                    }
                                    crate::settings::Language::Swedish => {
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
                                        &mut ps.settings.rss_removed_default_vi
                                    }
                                    crate::settings::Language::Czech => {
                                        &mut ps.settings.rss_removed_default_en
                                    }
                                    crate::settings::Language::Polish => {
                                        &mut ps.settings.rss_removed_default_pl
                                    }
                                    crate::settings::Language::French => {
                                        &mut ps.settings.rss_removed_default_fr
                                    }
                                    crate::settings::Language::Serbian => {
                                        &mut ps.settings.rss_removed_default_sr
                                    }
                                };
                                let already =
                                    removed_list.iter().any(|u| normalize_rss_url_key(u) == key);
                                if !already {
                                    removed_list.push(key.clone());
                                    default_removed_key_added = Some(key);
                                }
                            }
                        }
                    }
                    ps.settings.rss_sources.remove(idx);
                    delayed_target_source_idx = if ps.settings.rss_sources.is_empty() {
                        None
                    } else if idx > 0 {
                        Some(idx - 1)
                    } else {
                        Some(0)
                    };
                    crate::settings::save_settings(ps.settings.clone());
                });
                with_rss_state(hwnd, |s| s.suppress_tree_selection_events = true);
                SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(0), LPARAM(0));
                reload_tree(hwnd);
                SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
                with_rss_state(hwnd, |s| s.suppress_tree_selection_events = false);
                if let Some(source) = source_info {
                    with_rss_state(hwnd, |s| {
                        s.removed_history.push(RssLastRemoved::Source {
                            index: idx,
                            source,
                            language,
                            default_removed_key_added,
                        });
                    });
                }
                announce_rss_status(&i18n::tr(language, "rss.removed"));
                if let Some(target_idx) = delayed_target_source_idx {
                    schedule_delayed_source_select(hwnd, target_idx);
                }
            }
        } else if let Some(NodeData::Item(item)) = selected_node {
            let (language, require_confirm) = with_rss_state(hwnd, |s| {
                with_state(s.parent, |ps| {
                    (
                        ps.settings.language,
                        matches!(
                            ps.settings.rss_delete_confirm_mode,
                            crate::settings::RssDeleteConfirmMode::Article
                                | crate::settings::RssDeleteConfirmMode::Both
                        ),
                    )
                })
                .unwrap_or((crate::settings::Language::default(), true))
            })
            .unwrap_or((crate::settings::Language::default(), true));
            let title = if item.title.trim().is_empty() {
                item.link.clone()
            } else {
                item.title.clone()
            };
            let msg_template = i18n::tr(language, "rss.delete_confirm");
            let msg_text = msg_template.replace("{title}", &title);
            let caption = i18n::tr(language, "rss.delete_title");
            let confirmed = if require_confirm {
                MessageBoxW(
                    hwnd,
                    PCWSTR(to_wide(&msg_text).as_ptr()),
                    PCWSTR(to_wide(&caption).as_ptr()),
                    MB_YESNO | MB_ICONQUESTION,
                ) == IDYES
            } else {
                true
            };
            if confirmed {
                let parent_item = windows::Win32::UI::Controls::HTREEITEM(
                    SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(TVGN_PARENT as usize),
                        LPARAM(hitem.0),
                    )
                    .0,
                );
                let next_sibling = windows::Win32::UI::Controls::HTREEITEM(
                    SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(windows::Win32::UI::Controls::TVGN_NEXT as usize),
                        LPARAM(hitem.0),
                    )
                    .0,
                );
                let prev_sibling = windows::Win32::UI::Controls::HTREEITEM(
                    SendMessageW(
                        hwnd_tree,
                        TVM_GETNEXTITEM,
                        WPARAM(windows::Win32::UI::Controls::TVGN_PREVIOUS as usize),
                        LPARAM(hitem.0),
                    )
                    .0,
                );
                let key = rss_item_key(&item);
                let mut source_idx_for_undo: Option<usize> = None;
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if parent.0 != 0 {
                    let source_index =
                        with_rss_state(hwnd, |s| match s.node_data.get(&parent_item.0) {
                            Some(NodeData::Source(idx)) => Some(*idx),
                            _ => None,
                        })
                        .flatten();
                    if let Some(source_idx) = source_index {
                        source_idx_for_undo = Some(source_idx);
                        with_state(parent, |ps| {
                            if let Some(src) = ps.settings.rss_sources.get_mut(source_idx)
                                && !src.removed_item_keys.iter().any(|k| k == &key)
                            {
                                src.removed_item_keys.push(key.clone());
                                crate::settings::save_settings(ps.settings.clone());
                            }
                        });
                    }
                }
                let mut removed_position: Option<usize> = None;
                with_rss_state(hwnd, |s| {
                    s.node_data.remove(&hitem.0);
                    if parent_item.0 != 0
                        && let Some(state) = s.source_items.get_mut(&parent_item.0)
                        && let Some(pos) = state.items.iter().position(|x| rss_item_key(x) == key)
                    {
                        removed_position = Some(pos);
                        state.items.remove(pos);
                        state.loaded = state.loaded.saturating_sub(1);
                    }
                });
                SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(hitem.0));
                let next_exists = next_sibling.0 != 0
                    && with_rss_state(hwnd, |s| s.node_data.contains_key(&next_sibling.0))
                        .unwrap_or(false);
                let prev_exists = prev_sibling.0 != 0
                    && with_rss_state(hwnd, |s| s.node_data.contains_key(&prev_sibling.0))
                        .unwrap_or(false);
                if next_exists {
                    SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(next_sibling.0),
                    );
                } else if prev_exists {
                    SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(prev_sibling.0),
                    );
                } else if parent_item.0 != 0 {
                    SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(parent_item.0),
                    );
                }
                if let (Some(source_index), Some(position)) =
                    (source_idx_for_undo, removed_position)
                {
                    with_rss_state(hwnd, |s| {
                        s.removed_history.push(RssLastRemoved::Item {
                            source_index,
                            item,
                            key,
                            position,
                        });
                    });
                }
                announce_rss_status(&i18n::tr(language, "rss.article_removed"));
            }
        }
        if hwnd_tree.0 != 0 && GetFocus() != hwnd_tree {
            with_rss_state(hwnd, |s| s.suppress_focus_restore_once = true);
            SetFocus(hwnd_tree);
        }
    }
}

fn undo_last_delete(hwnd: HWND) {
    unsafe {
        let Some(last_removed) = with_rss_state(hwnd, |s| s.removed_history.pop()).flatten() else {
            return;
        };

        match last_removed {
            RssLastRemoved::Source {
                index,
                source,
                language,
                default_removed_key_added,
            } => {
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if parent.0 == 0 {
                    return;
                }
                let mut restored_index = index;
                with_state(parent, |ps| {
                    let insert_at = index.min(ps.settings.rss_sources.len());
                    restored_index = insert_at;
                    ps.settings.rss_sources.insert(insert_at, source);
                    if let Some(key) = default_removed_key_added {
                        let removed_list = match language {
                            crate::settings::Language::Ukrainian
                            | crate::settings::Language::English
                            | crate::settings::Language::Lithuanian
                            | crate::settings::Language::Chinese => {
                                &mut ps.settings.rss_removed_default_en
                            }
                            crate::settings::Language::Swedish => {
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
                                &mut ps.settings.rss_removed_default_vi
                            }
                            crate::settings::Language::Czech => {
                                &mut ps.settings.rss_removed_default_en
                            }
                            crate::settings::Language::Polish => {
                                &mut ps.settings.rss_removed_default_pl
                            }
                            crate::settings::Language::French => {
                                &mut ps.settings.rss_removed_default_fr
                            }
                            crate::settings::Language::Serbian => {
                                &mut ps.settings.rss_removed_default_sr
                            }
                        };
                        removed_list.retain(|u| normalize_rss_url_key(u) != key);
                    }
                    crate::settings::save_settings(ps.settings.clone());
                })
                .unwrap_or(());

                let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                with_rss_state(hwnd, |s| s.suppress_tree_selection_events = true);
                if hwnd_tree.0 != 0 {
                    SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(0), LPARAM(0));
                }
                reload_tree(hwnd);
                if hwnd_tree.0 != 0 {
                    SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
                    with_rss_state(hwnd, |s| s.suppress_tree_selection_events = false);
                    if GetFocus() != hwnd_tree {
                        with_rss_state(hwnd, |s| s.suppress_focus_restore_once = true);
                        SetFocus(hwnd_tree);
                    }
                } else {
                    with_rss_state(hwnd, |s| s.suppress_tree_selection_events = false);
                }
                schedule_delayed_source_select(hwnd, restored_index);
            }
            RssLastRemoved::Item {
                source_index,
                item,
                key,
                position,
            } => {
                let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                if parent.0 != 0 {
                    with_state(parent, |ps| {
                        if let Some(src) = ps.settings.rss_sources.get_mut(source_index) {
                            src.removed_item_keys.retain(|k| k != &key);
                            crate::settings::save_settings(ps.settings.clone());
                        }
                    });
                }

                let source_hitem = with_rss_state(hwnd, |s| {
                    s.node_data.iter().find_map(|(&h, node)| match node {
                        NodeData::Source(i) if *i == source_index => {
                            Some(windows::Win32::UI::Controls::HTREEITEM(h))
                        }
                        _ => None,
                    })
                })
                .flatten();

                let Some(source_hitem) = source_hitem else {
                    return;
                };

                let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
                let mut show_in_tree = false;
                with_rss_state(hwnd, |s| {
                    if let Some(state) = s.source_items.get_mut(&source_hitem.0) {
                        let insert_at = position.min(state.items.len());
                        if state.loaded > 0 && insert_at <= state.loaded {
                            state.loaded += 1;
                            show_in_tree = true;
                        }
                        state.items.insert(insert_at, item.clone());
                    }
                });

                if show_in_tree && hwnd_tree.0 != 0 {
                    with_rss_state(hwnd, |s| s.suppress_tree_selection_events = true);
                    // Rebuild loaded children to preserve exact order after undo.
                    SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(0), LPARAM(0));
                    SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(source_hitem.0),
                    );
                    loop {
                        let child = windows::Win32::UI::Controls::HTREEITEM(
                            SendMessageW(
                                hwnd_tree,
                                TVM_GETNEXTITEM,
                                WPARAM(TVGN_CHILD as usize),
                                LPARAM(source_hitem.0),
                            )
                            .0,
                        );
                        if child.0 == 0 {
                            break;
                        }
                        with_rss_state(hwnd, |s| {
                            s.node_data.remove(&child.0);
                        });
                        SendMessageW(hwnd_tree, TVM_DELETEITEM, WPARAM(0), LPARAM(child.0));
                    }

                    let mut restored_hitem = windows::Win32::UI::Controls::HTREEITEM(0);
                    let (
                        language,
                        announce_unread,
                        unread_label_position,
                        rss_date_mode,
                        rss_time_mode,
                    ) = with_rss_state(hwnd, |s| {
                        with_state(s.parent, |ps| {
                            (
                                ps.settings.language,
                                ps.settings.announce_unread_rss_podcast_items,
                                ps.settings.rss_podcast_unread_label_position,
                                ps.settings.rss_articles_date_display,
                                ps.settings.rss_articles_time_display,
                            )
                        })
                        .unwrap_or((
                            crate::settings::Language::English,
                            true,
                            crate::settings::RssPodcastUnreadLabelPosition::Before,
                            ListDateDisplayMode::Always,
                            ListTimeDisplayMode::Always,
                        ))
                    })
                    .unwrap_or((
                        crate::settings::Language::English,
                        true,
                        crate::settings::RssPodcastUnreadLabelPosition::Before,
                        ListDateDisplayMode::Always,
                        ListTimeDisplayMode::Always,
                    ));
                    with_rss_state(hwnd, |s| {
                        if let Some(state) = s.source_items.get(&source_hitem.0) {
                            let day_counts = build_day_counts(&state.items);
                            for entry in state.items.iter().take(state.loaded) {
                                if entry.title.trim().is_empty() {
                                    continue;
                                }
                                let item_unread =
                                    !state.read_item_keys.contains(&rss_item_key(entry));
                                let title_ctx = RssItemTitleContext {
                                    language,
                                    announce_unread,
                                    unread_label_position,
                                    date_mode: rss_date_mode,
                                    time_mode: rss_time_mode,
                                };
                                let display_title = rss_item_display_title(
                                    &entry.title,
                                    item_unread,
                                    entry.pub_date,
                                    has_multiple_items_same_day(entry.pub_date, &day_counts),
                                    title_ctx,
                                );
                                let text = to_wide(&display_title);
                                let mut tvis = TVINSERTSTRUCTW {
                                    hParent: source_hitem,
                                    hInsertAfter: TVI_LAST,
                                    Anonymous: TVINSERTSTRUCTW_0 {
                                        item: TVITEMW {
                                            mask: TVIF_TEXT
                                                | TVIF_PARAM
                                                | windows::Win32::UI::Controls::TVIF_CHILDREN,
                                            pszText: windows::core::PWSTR(text.as_ptr() as *mut _),
                                            cChildren: TVITEMEXW_CHILDREN(if entry.is_folder {
                                                1
                                            } else {
                                                0
                                            }),
                                            lParam: LPARAM(0),
                                            ..Default::default()
                                        },
                                    },
                                };
                                let hchild = windows::Win32::UI::Controls::HTREEITEM(
                                    SendMessageW(
                                        hwnd_tree,
                                        TVM_INSERTITEMW,
                                        WPARAM(0),
                                        LPARAM(&mut tvis as *mut _ as isize),
                                    )
                                    .0,
                                );
                                if hchild.0 != 0 {
                                    s.node_data.insert(hchild.0, NodeData::Item(entry.clone()));
                                    if rss_item_key(entry) == key {
                                        restored_hitem = hchild;
                                    }
                                }
                            }
                        }
                    });
                    SendMessageW(hwnd_tree, WM_SETREDRAW, WPARAM(1), LPARAM(0));
                    with_rss_state(hwnd, |s| s.suppress_tree_selection_events = false);
                    if restored_hitem.0 != 0 {
                        SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(restored_hitem.0),
                        );
                        SendMessageW(
                            hwnd_tree,
                            TVM_ENSUREVISIBLE,
                            WPARAM(0),
                            LPARAM(restored_hitem.0),
                        );
                    } else {
                        SendMessageW(
                            hwnd_tree,
                            TVM_SELECTITEM,
                            WPARAM(TVGN_CARET as usize),
                            LPARAM(source_hitem.0),
                        );
                    }
                    if GetFocus() != hwnd_tree {
                        with_rss_state(hwnd, |s| s.suppress_focus_restore_once = true);
                        SetFocus(hwnd_tree);
                    }
                } else if hwnd_tree.0 != 0 {
                    SendMessageW(
                        hwnd_tree,
                        TVM_SELECTITEM,
                        WPARAM(TVGN_CARET as usize),
                        LPARAM(source_hitem.0),
                    );
                    if GetFocus() != hwnd_tree {
                        with_rss_state(hwnd, |s| s.suppress_focus_restore_once = true);
                        SetFocus(hwnd_tree);
                    }
                }
            }
        }
    }
}

fn handle_edit_source(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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
        {
            with_state(s.parent, |ps| {
                ps.settings
                    .rss_sources
                    .get(idx)
                    .map(|src| (src.title.clone(), src.url.clone()))
            })
        }
        .flatten()
    })
    .flatten();
    let (title, url) = source_info.unwrap_or_default();
    if url.trim().is_empty() {
        return;
    }

    let main_hwnd = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let existing = { with_state(main_hwnd, |s| s.rss_add_dialog) }.unwrap_or(HWND(0));
    if existing.0 != 0 {
        crate::set_foreground_window_safe(existing);
        return;
    }

    with_rss_state(hwnd, |s| s.pending_edit = Some(idx));
    show_add_dialog_with_prefill(hwnd, title, url);
}

fn handle_retry_now(hwnd: HWND) {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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

    let source_info = with_rss_state(hwnd, |s| match s.node_data.get(&hitem.0) {
        Some(NodeData::Source(idx)) => {
            with_state(s.parent, |ps| {
                ps.settings
                    .rss_sources
                    .get(*idx)
                    .map(|src| (src.url.clone(), src.kind.clone(), src.cache.clone()))
            })
        }
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
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    let url_clone = url.clone();
    std::thread::spawn(move || {
        let rt = match tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
        {
            Ok(rt) => rt,
            Err(e) => {
                crate::log_debug(&format!("Failed to build tokio runtime: {}", e));
                return;
            }
        };
        let res = rt.block_on(rss::fetch_and_parse(
            &url_clone,
            source_kind,
            cache,
            fetch_config,
            true,
            language,
        ));
        let msg = Box::new(FetchResult {
            hitem: hitem.0,
            result: res,
        });
        if let Err(_e) = crate::post_message_w_safe(
            hwnd,
            WM_RSS_FETCH_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        ) {}
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
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return None;
    }
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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
        Some(NodeData::Source(idx)) => Some(*idx),
        _ => None,
    })
    .flatten()
}

fn apply_reorder_action(
    hwnd: HWND,
    source_index: usize,
    action: ReorderAction,
    target_index: usize,
) -> Option<usize> {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 || parent.0 == 0 {
        return None;
    }
    let mut root_items = collect_root_items(hwnd_tree);
    if source_index >= root_items.len() {
        return None;
    }
    let new_index = {
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

fn handle_reorder_action(hwnd: HWND, action: ReorderAction) {
    let Some(source_index) = selected_source_index(hwnd) else {
        return;
    };
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    let total = { with_state(parent, |ps| ps.settings.rss_sources.len()) }.unwrap_or(0);
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
    if let Some(new_index) = new_index
        && new_index != source_index
    {
        let template = i18n::tr(language, "rss.reorder.moved_position");
        let message = template.replace("{x}", &(new_index + 1).to_string());
        announce_rss_status(&message);
    }
}

fn handle_sort_action(hwnd: HWND, order: crate::settings::SortOrder) {
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    {
        with_state(parent, |ps| {
            crate::settings::sort_rss_sources(&mut ps.settings, order);
            crate::settings::save_settings(ps.settings.clone());
        });
        reload_tree(hwnd);
    }
}

#[derive(Clone, Copy)]
enum ArticleAction {
    OpenInBrowser,
    ShareFacebook,
    ShareTwitter,
    ShareWhatsApp,
    ShareEmail,
}

fn selected_article_item(hwnd: HWND) -> Option<RssItem> {
    let hwnd_tree = with_rss_state(hwnd, |s| s.hwnd_tree).unwrap_or(HWND(0));
    if hwnd_tree.0 == 0 {
        return None;
    }
    let hitem = windows::Win32::UI::Controls::HTREEITEM(
        crate::send_message_w_safe(
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

fn copy_text_to_clipboard(hwnd: HWND, text: &str) {
    use windows::Win32::Foundation::HANDLE;
    use windows::Win32::System::Memory::{GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalUnlock};

    const CF_UNICODETEXT: u32 = 13;

    let content = to_wide(text);
    if content.is_empty() {
        return;
    }
    if unsafe { OpenClipboard(hwnd) }.is_err() {
        return;
    }
    if let Err(e) = unsafe { EmptyClipboard() } {
        crate::log_debug(&format!("EmptyClipboard failed: {}", e));
    }
    let size = content.len() * std::mem::size_of::<u16>();
    let handle = match unsafe { GlobalAlloc(GMEM_MOVEABLE, size) } {
        Ok(handle) => handle,
        Err(_) => {
            if let Err(e) = unsafe { CloseClipboard() } {
                crate::log_debug(&format!("CloseClipboard failed: {}", e));
            }
            return;
        }
    };
    if handle.0.is_null() {
        if let Err(e) = unsafe { CloseClipboard() } {
            crate::log_debug(&format!("CloseClipboard failed: {}", e));
        }
        return;
    }
    let ptr = unsafe { GlobalLock(handle) as *mut u16 };
    if ptr.is_null() {
        if let Err(e) = unsafe { CloseClipboard() } {
            crate::log_debug(&format!("CloseClipboard failed: {}", e));
        }
        return;
    }
    unsafe {
        std::ptr::copy_nonoverlapping(content.as_ptr(), ptr, content.len());
    }
    crate::log_if_err!(unsafe { GlobalUnlock(handle) });
    if let Err(e) = unsafe { SetClipboardData(CF_UNICODETEXT, HANDLE(handle.0 as isize)) } {
        crate::log_debug(&format!("SetClipboardData failed: {}", e));
    }
    if let Err(e) = unsafe { CloseClipboard() } {
        crate::log_debug(&format!("CloseClipboard failed: {}", e));
    }
}

fn fetch_full_article_text_for_quick_copy(parent: HWND, item: &RssItem) -> String {
    if parent.0 == 0 {
        return clean_html_for_quick_copy(&item.description);
    }
    ensure_rss_http(parent);
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    let rt = match tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
    {
        Ok(rt) => rt,
        Err(_) => return clean_html_for_quick_copy(&item.description),
    };
    let text = match rt.block_on(crate::tools::rss::fetch_article_text(
        &item.link,
        &item.title,
        &item.description,
        language,
    )) {
        Ok(res) => res,
        Err(_) => item.description.clone(),
    };
    let normalized = normalize_article_text(&text);
    if normalized.trim().is_empty() {
        clean_html_for_quick_copy(&item.description)
    } else {
        normalized
    }
}

fn rss_quick_copy_text(
    parent: HWND,
    item: &RssItem,
    mode: crate::settings::RssQuickCopyMode,
) -> String {
    let title = item.title.trim();
    let url = item.link.trim();
    match mode {
        crate::settings::RssQuickCopyMode::Title => title.to_string(),
        crate::settings::RssQuickCopyMode::Url => url.to_string(),
        crate::settings::RssQuickCopyMode::Content => {
            fetch_full_article_text_for_quick_copy(parent, item)
        }
        crate::settings::RssQuickCopyMode::All => {
            let content = fetch_full_article_text_for_quick_copy(parent, item);
            let mut parts = Vec::new();
            if !title.is_empty() {
                parts.push(title.to_string());
            }
            if !url.is_empty() {
                parts.push(url.to_string());
            }
            if !content.is_empty() {
                parts.push(content);
            }
            parts.join("\r\n")
        }
    }
}

fn clean_html_for_quick_copy(input: &str) -> String {
    let mut out = String::with_capacity(input.len());
    let mut in_tag = false;
    let mut prev_space = false;

    for ch in input.chars() {
        match ch {
            '<' => {
                in_tag = true;
            }
            '>' => {
                in_tag = false;
                if !prev_space {
                    out.push(' ');
                    prev_space = true;
                }
            }
            _ if in_tag => {}
            '\r' | '\n' | '\t' => {
                if !prev_space {
                    out.push(' ');
                    prev_space = true;
                }
            }
            _ => {
                out.push(ch);
                prev_space = ch.is_whitespace();
            }
        }
    }

    normalize_article_text(out.trim())
}

fn handle_rss_quick_copy(hwnd: HWND) -> bool {
    let Some(item) = selected_article_item(hwnd) else {
        return false;
    };
    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let mode = { with_state(parent, |ps| ps.settings.rss_quick_copy_mode) }.unwrap_or_default();
    let language = { with_state(parent, |ps| ps.settings.language) }.unwrap_or_default();
    let text = rss_quick_copy_text(parent, &item, mode);
    if text.trim().is_empty() {
        return false;
    }
    copy_text_to_clipboard(hwnd, &text);
    announce_rss_status(&i18n::tr(language, "rss.copied"));
    true
}

fn handle_article_action(hwnd: HWND, action: ArticleAction) {
    let Some(item) = selected_article_item(hwnd) else {
        log_debug("rss_action kind=article action=unavailable reason=no_article");
        return;
    };
    let url = item.link.trim().to_string();
    if !is_valid_article_url(&url) {
        log_debug("rss_action kind=article action=unavailable reason=invalid_url");
        return;
    }
    if matches!(action, ArticleAction::OpenInBrowser)
        && crate::tools::rss::is_google_news_article_url(&url)
    {
        std::thread::spawn(move || {
            let final_url = match crate::tools::rss::resolve_google_news_article_url_blocking(&url)
            {
                Ok(Some(decoded)) => decoded,
                Ok(None) => url.clone(),
                Err(err) => {
                    log_debug(&format!(
                        "rss_action kind=article action=google_news_resolve_failed error=\"{}\"",
                        err
                    ));
                    url.clone()
                }
            };
            if let Err(err) = crate::audio_utils::open_url_in_browser(&final_url) {
                log_debug(&format!(
                    "rss_action kind=article action=browser_error error=\"{}\"",
                    err
                ));
            }
        });
        return;
    }
    let title = item.title.trim().to_string();
    let language = with_rss_state(hwnd, |s| {
        { with_state(s.parent, |ps| ps.settings.language) }.unwrap_or_default()
    })
    .unwrap_or_default();
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
        ArticleAction::ShareEmail => {
            log_debug(&format!(
                "rss_action kind=article action=share_email url=\"{}\"",
                url
            ));
            let subject_raw = if title.is_empty() {
                i18n::tr(language, "rss.email.default_subject")
            } else {
                title.clone()
            };
            let subject = decode_mail_text_component(&subject_raw);
            let intro = decode_mail_text_component(&i18n::tr(language, "rss.email.body_intro"));
            let body = format!("\r\n{}\r\n{}\r\n", intro, url);
            format!(
                "mailto:?subject={}&body={}",
                mailto_encode_component(&subject),
                mailto_encode_component(&body)
            )
        }
    };
    if let Err(err) = crate::audio_utils::open_url_in_browser(&share_url) {
        log_debug(&format!(
            "rss_action kind=article action=browser_error error=\"{}\"",
            err
        ));
    }
}

fn force_focus_editor_on_parent(parent: HWND) {
    if parent.0 == 0 {
        return;
    }
    unsafe {
        SetForegroundWindow(parent);
        SetActiveWindow(parent);
        SendMessageW(parent, WM_SETFOCUS, WPARAM(0), LPARAM(0));
    }
    if crate::get_active_edit(parent).is_none() {
        crate::send_message_w_safe(
            parent,
            WM_COMMAND,
            WPARAM(crate::menu::IDM_FILE_NEW),
            LPARAM(0),
        );
    }
    if let Some(hwnd_edit) = crate::get_active_edit(parent) {
        unsafe {
            SetFocus(hwnd_edit);
            SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            SendMessageW(
                parent,
                WM_NEXTDLGCTL,
                WPARAM(hwnd_edit.0 as usize),
                LPARAM(1),
            );
        }
        // Re-assert focus after dialog navigation to help NVDA settle on the edit control.
        unsafe {
            SetFocus(hwnd_edit);
            SendMessageW(hwnd_edit, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            NotifyWinEvent(
                EVENT_OBJECT_FOCUS,
                hwnd_edit,
                OBJID_CLIENT.0,
                CHILDID_SELF as i32,
            );
        }
    }
    crate::send_message_w_safe(parent, WM_SETFOCUS, WPARAM(0), LPARAM(0));
    if let Err(_e) =
        crate::post_message_w_safe(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0))
    {
        crate::log_debug(&format!("Error: {:?}", _e));
    }
}

fn import_item(hwnd: HWND, item: RssItem) {
    let url = item.link.clone();

    let parent = with_rss_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = if parent.0 != 0 {
        { with_state(parent, |state| state.settings.language) }.unwrap_or_default()
    } else {
        crate::settings::Language::default()
    };
    if parent.0 != 0 {
        ensure_rss_http(parent);
    }

    std::thread::spawn(move || {
        let rt = match tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
        {
            Ok(rt) => rt,
            Err(e) => {
                crate::log_debug(&format!("Failed to build tokio runtime: {}", e));
                return;
            }
        };
        let content_res = rt.block_on(crate::tools::rss::fetch_article_text(
            &url,
            &item.title,
            &item.description,
            language,
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
        if let Err(_e) = crate::post_message_w_safe(
            hwnd,
            WM_RSS_IMPORT_COMPLETE,
            WPARAM(0),
            LPARAM(Box::into_raw(msg) as isize),
        ) {}
    });
}

struct ImportResult {
    text: String,
}

struct MarkItemReadUiMessage {
    hitem: isize,
    item_key: String,
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

fn show_reorder_dialog(parent_hwnd: HWND, source_index: usize, total: usize) {
    let existing = with_rss_state(parent_hwnd, |s| s.reorder_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        crate::set_foreground_window_safe(existing);
        return;
    }
    let hinstance = unsafe { HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0) };
    let class_name = to_wide("SonarpadRssReorder");
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            unsafe {
                windows::Win32::UI::WindowsAndMessaging::LoadCursorW(
                    None,
                    windows::Win32::UI::WindowsAndMessaging::IDC_ARROW,
                )
            }
            .unwrap_or_default()
            .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(reorder_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let language = with_rss_state(parent_hwnd, |s| {
        { with_state(s.parent, |ps| ps.settings.language) }.unwrap_or_default()
    })
    .unwrap_or_default();
    let title = i18n::tr(language, "rss.context.reorder");
    let init_ptr = Box::into_raw(Box::new(ReorderDialogInit {
        parent: parent_hwnd,
        source_index,
        total,
    }));
    let hwnd = unsafe {
        CreateWindowExW(
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
        )
    };
    if hwnd.0 == 0 {
        let _unused_box = unsafe { Box::from_raw(init_ptr) };
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
    unsafe {
        crate::panic_guard::guard(
            "reorder_control_subclass_proc",
            || DefWindowProcW(hwnd, msg, wparam, lparam),
            || reorder_control_subclass_proc_inner(hwnd, msg, wparam, lparam),
        )
    }
}

fn reorder_control_subclass_proc_inner(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == windows::Win32::UI::WindowsAndMessaging::WM_CHAR && wparam.0 as u16 == VK_TAB.0 {
        return LRESULT(0);
    }
    if msg == WM_KEYDOWN {
        let id = crate::get_dlg_ctrl_id_safe(hwnd);
        let parent = crate::get_parent_safe(hwnd);
        let (edit_id, ok_id, cancel_id) =
            if id == REORDER_EDIT_ID || id == REORDER_OK_ID || id == REORDER_CANCEL_ID {
                (REORDER_EDIT_ID, REORDER_OK_ID, REORDER_CANCEL_ID)
            } else if id == SEARCH_EDIT_ID || id == SEARCH_OK_ID || id == SEARCH_CANCEL_ID {
                (SEARCH_EDIT_ID, SEARCH_OK_ID, SEARCH_CANCEL_ID)
            } else {
                (0, 0, 0)
            };
        if edit_id == 0 {
            let prev = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA);
            if prev == 0 {
                return crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam);
            }
            return unsafe {
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
            };
        }
        let edit = crate::get_dlg_item_safe(parent, edit_id as i32);
        let ok = crate::get_dlg_item_safe(parent, ok_id as i32);
        let cancel = crate::get_dlg_item_safe(parent, cancel_id as i32);
        if wparam.0 as u16 == VK_TAB.0 {
            let shift = (crate::get_key_state_safe(VK_SHIFT.0 as i32) & 0x8000u16 as i16) != 0;
            let next = if shift {
                if id == edit_id {
                    cancel
                } else if id == cancel_id {
                    ok
                } else {
                    edit
                }
            } else if id == edit_id {
                ok
            } else if id == ok_id {
                cancel
            } else {
                edit
            };
            crate::set_focus_safe(next);
            return LRESULT(0);
        }
        if wparam.0 as u16 == VK_RETURN.0 {
            let target = if id == cancel_id { cancel_id } else { ok_id };
            crate::send_message_w_safe(parent, WM_COMMAND, WPARAM(target), LPARAM(0));
            return LRESULT(0);
        }
        if wparam.0 as u16 == VK_ESCAPE.0 {
            crate::send_message_w_safe(parent, WM_COMMAND, WPARAM(cancel_id), LPARAM(0));
            return LRESULT(0);
        }
    }
    let prev = crate::get_window_long_ptr_w_safe(hwnd, GWLP_USERDATA);
    if prev == 0 {
        return crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam);
    }
    unsafe {
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
}

fn show_add_dialog(parent_hwnd: HWND) {
    show_add_dialog_with_prefill(parent_hwnd, String::new(), String::new());
}

fn show_rss_search_dialog(parent_hwnd: HWND) {
    let exists = with_rss_state(parent_hwnd, |s| s.search_dialog).unwrap_or(HWND(0));
    if exists.0 != 0 {
        crate::set_foreground_window_safe(exists);
        return;
    }

    let hinstance = unsafe { HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0) };
    let class_name = to_wide("SonarpadRssSearchKeyword");
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            unsafe {
                windows::Win32::UI::WindowsAndMessaging::LoadCursorW(
                    None,
                    windows::Win32::UI::WindowsAndMessaging::IDC_ARROW,
                )
            }
            .unwrap_or_default()
            .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(search_keyword_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let main_hwnd = with_rss_state(parent_hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = { with_state(main_hwnd, |s| s.settings.language) }.unwrap_or_default();
    let title = tr_or(language, "rss.search_dialog.title", "Search RSS by keyword");
    let init_ptr = Box::into_raw(Box::new(SearchDialogInit {
        parent: parent_hwnd,
    }));
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(to_wide(&title).as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_POPUP,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            420,
            170,
            parent_hwnd,
            None,
            hinstance,
            Some(init_ptr as *const _),
        )
    };
    if hwnd.0 == 0 {
        let _unused_box = unsafe { Box::from_raw(init_ptr) };
        return;
    }
    with_rss_state(parent_hwnd, |s| s.search_dialog = hwnd);
}

unsafe extern "system" fn search_keyword_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    crate::panic_guard::guard(
        "search_keyword_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || search_keyword_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn search_keyword_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*cs).lpCreateParams as *mut SearchDialogInit;
                let parent = if init_ptr.is_null() {
                    HWND(0)
                } else {
                    let init = Box::from_raw(init_ptr);
                    init.parent
                };
                let main_hwnd = with_rss_state(parent, |s| s.parent).unwrap_or(HWND(0));
                let language = with_state(main_hwnd, |s| s.settings.language).unwrap_or_default();
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, parent.0);

                let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
                let keyword_label = tr_or(language, "rss.search_dialog.keyword_label", "Keyword:");
                CreateWindowExW(
                    Default::default(),
                    w!("STATIC"),
                    PCWSTR(to_wide(&keyword_label).as_ptr()),
                    WS_CHILD | WS_VISIBLE,
                    10,
                    12,
                    390,
                    18,
                    hwnd,
                    HMENU(1601),
                    hinstance,
                    None,
                );
                let edit = CreateWindowExW(
                    WS_EX_CLIENTEDGE,
                    w!("EDIT"),
                    PCWSTR::null(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
                    10,
                    34,
                    390,
                    24,
                    hwnd,
                    HMENU(SEARCH_EDIT_ID as isize),
                    hinstance,
                    None,
                );
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.ok")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                    220,
                    80,
                    85,
                    28,
                    hwnd,
                    HMENU(SEARCH_OK_ID as isize),
                    hinstance,
                    None,
                );
                CreateWindowExW(
                    Default::default(),
                    WC_BUTTON,
                    PCWSTR(to_wide(&i18n::tr(language, "rss.dialog.cancel")).as_ptr()),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    315,
                    80,
                    85,
                    28,
                    hwnd,
                    HMENU(SEARCH_CANCEL_ID as isize),
                    hinstance,
                    None,
                );
                let ok = GetDlgItem(hwnd, SEARCH_OK_ID as i32);
                let cancel = GetDlgItem(hwnd, SEARCH_CANCEL_ID as i32);
                let proc_ptr = reorder_control_subclass_proc as *const () as usize;
                for control in [edit, ok, cancel] {
                    let prev = SetWindowLongPtrW(control, GWLP_WNDPROC, proc_ptr as isize);
                    SetWindowLongPtrW(control, GWLP_USERDATA, prev);
                }
                SetFocus(edit);
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
                if id == SEARCH_CANCEL_ID || id == 2 {
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    return LRESULT(0);
                }
                if id == SEARCH_OK_ID || id == 1 {
                    let parent = HWND(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
                    let main_hwnd = with_rss_state(parent, |s| s.parent).unwrap_or(HWND(0));
                    let language =
                        with_state(main_hwnd, |s| s.settings.language).unwrap_or_default();
                    let edit = GetDlgItem(hwnd, SEARCH_EDIT_ID as i32);
                    let mut buf = vec![0u16; 1024];
                    let len =
                        windows::Win32::UI::WindowsAndMessaging::GetWindowTextW(edit, &mut buf)
                            as usize;
                    let keyword = String::from_utf16_lossy(&buf[..len]).trim().to_string();
                    if keyword.is_empty() {
                        let title =
                            tr_or(language, "rss.search_dialog.title", "Search RSS by keyword");
                        let message = tr_or(
                            language,
                            "rss.search_dialog.empty_keyword",
                            "Please enter a keyword.",
                        );
                        MessageBoxW(
                            hwnd,
                            PCWSTR(to_wide(&message).as_ptr()),
                            PCWSTR(to_wide(&title).as_ptr()),
                            MB_OK | MB_ICONINFORMATION,
                        );
                        return LRESULT(0);
                    }
                    let url = build_google_news_rss_url(&keyword, language);
                    let source_title = format_google_news_source_title(&keyword);
                    show_add_dialog_with_prefill_options(parent, source_title, url, true);
                    crate::log_if_err!(crate::destroy_window_safe(hwnd));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
            WM_CLOSE => {
                crate::log_if_err!(crate::destroy_window_safe(hwnd));
                LRESULT(0)
            }
            WM_NCDESTROY => {
                let parent = HWND(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
                if parent.0 != 0 {
                    with_rss_state(parent, |s| s.search_dialog = HWND(0));
                }
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn show_add_dialog_with_prefill(parent_hwnd: HWND, title: String, url: String) {
    show_add_dialog_with_prefill_options(parent_hwnd, title, url, false);
}

fn show_add_dialog_with_prefill_options(
    parent_hwnd: HWND,
    title: String,
    url: String,
    hide_url_field: bool,
) {
    let hinstance = unsafe { HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0) };
    let class_name = to_wide("SonarpadInput");

    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(
            unsafe {
                windows::Win32::UI::WindowsAndMessaging::LoadCursorW(
                    None,
                    windows::Win32::UI::WindowsAndMessaging::IDC_ARROW,
                )
            }
            .unwrap_or_default()
            .0,
        ),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(input_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let main_hwnd = with_rss_state(parent_hwnd, |s| s.parent).unwrap_or(HWND(0));
    let language = { with_state(main_hwnd, |s| s.settings.language) }.unwrap_or_default();
    let init_ptr = Box::into_raw(Box::new(AddDialogInit {
        parent: parent_hwnd,
        prefill_title: title,
        prefill_url: url,
        hide_url_field,
    }));
    let hwnd = unsafe {
        CreateWindowExW(
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
        )
    };
    if hwnd.0 == 0 {
        let _unused_box = unsafe { Box::from_raw(init_ptr) };
        return;
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
    crate::panic_guard::guard(
        "input_wndproc",
        || crate::def_window_proc_w_safe(hwnd, msg, wparam, lparam),
        || input_wndproc_inner(hwnd, msg, wparam, lparam),
    )
}

fn input_wndproc_inner(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                let cs = lparam.0 as *const CREATESTRUCTW;
                let init_ptr = (*cs).lpCreateParams as *mut AddDialogInit;
                let (parent, prefill_title, prefill_url, hide_url_field): (
                    HWND,
                    String,
                    String,
                    bool,
                ) = if init_ptr.is_null() {
                    (HWND(0), String::new(), String::new(), false)
                } else {
                    let init = Box::from_raw(init_ptr);
                    (
                        init.parent,
                        init.prefill_title,
                        init.prefill_url,
                        init.hide_url_field,
                    )
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
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
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
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
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
                if !prefill_url.trim().is_empty()
                    && let Err(_e) = SetWindowTextW(
                        GetDlgItem(hwnd, 101),
                        PCWSTR(to_wide(&prefill_url).as_ptr()),
                    )
                {}
                if !prefill_title.trim().is_empty()
                    && let Err(_e) = SetWindowTextW(
                        GetDlgItem(hwnd, 104),
                        PCWSTR(to_wide(&prefill_title).as_ptr()),
                    )
                {}
                if hide_url_field {
                    ShowWindow(GetDlgItem(hwnd, 105), SW_HIDE);
                    ShowWindow(GetDlgItem(hwnd, 101), SW_HIDE);
                    SetFocus(GetDlgItem(hwnd, 104));
                } else {
                    SetFocus(GetDlgItem(hwnd, 101));
                }
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = wparam.0 & 0xffff;
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
                        let url = String::from_utf16_lossy(&buf[..len as usize]);

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
                        let title = String::from_utf16_lossy(&tbuf[..tlen as usize]);

                        if !url.trim().is_empty() {
                            let parent = windows::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
                            let main_hwnd = with_rss_state(parent, |s| s.parent).unwrap_or(HWND(0));
                            let language =
                                with_state(main_hwnd, |s| s.settings.language).unwrap_or_default();
                            let mut source_url = url.trim().to_string();
                            let mut source_title = title.trim().to_string();
                            if !is_valid_article_url(&source_url) {
                                source_url = build_google_news_rss_url(&source_url, language);
                                if source_title.is_empty() {
                                    source_title = format_google_news_source_title(url.trim());
                                }
                            }
                            if source_title.is_empty() {
                                source_title = source_url.clone();
                            }

                            let payload = format!("{}\n{}", source_title, source_url);
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
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
                        LRESULT(0)
                    }
                    103 | 2 => {
                        crate::log_if_err!(crate::destroy_window_safe(hwnd));
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
}
