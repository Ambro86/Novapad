#!/usr/bin/env python3
"""Rewrite bare Win32 API call sites to *_safe helpers, then demote pure-safe unsafe blocks.

Pass 1: identifier( -> helper( at call sites (not imports, not *_safe helper bodies).
Pass 2: unwrap `unsafe { panic_guard::guard(...) }` when fallback is already safe.
Pass 3: demote `unsafe { body }` when body no longer needs unsafe (conservative).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "src"

# API name -> helper base name (crate:: prefix added outside main.rs)
API_TO_HELPER = {
    "SendMessageW": "send_message_w_safe",
    "PostMessageW": "post_message_w_safe",
    "DestroyWindow": "destroy_window_safe",
    "SetFocus": "set_focus_safe",
    "GetFocus": "get_focus_safe",
    "SetForegroundWindow": "set_foreground_window_safe",
    "GetForegroundWindow": "get_foreground_window_safe",
    "ShowWindow": "show_window_safe",
    "EnableWindow": "enable_window_safe",
    "SetWindowTextW": "set_window_text_w_safe",
    "GetWindowTextW": "get_window_text_w_safe",
    "GetWindowTextLengthW": "get_window_text_length_w_safe",
    "GetWindowLongPtrW": "get_window_long_ptr_w_safe",
    "SetWindowLongPtrW": "set_window_long_ptr_w_safe",
    "DefWindowProcW": "def_window_proc_w_safe",
    "CallWindowProcW": "call_window_proc_w_safe",
    "GetParent": "get_parent_safe",
    "GetDlgItem": "get_dlg_item_safe",
    "KillTimer": "kill_timer_safe",
    "MessageBoxW": "message_box_w_safe",
    "GetClientRect": "get_client_rect_safe",
    "GetWindowRect": "get_window_rect_safe",
    "GetKeyState": "get_key_state_safe",
    "GetMenu": "get_menu_safe",
    "IsDialogMessageW": "is_dialog_message_w_safe",
    "GetClassNameW": "get_class_name_w_safe",
    "GetNextDlgTabItem": "get_next_dlg_tab_item_safe",
    "FindWindowW": "find_window_w_safe",
    "GetCursorPos": "get_cursor_pos_safe",
    "CreateMenu": "create_menu_safe",
    "DestroyMenu": "destroy_menu_safe",
    "AppendMenuW": "append_menu_w_safe",
    "CheckMenuItem": "check_menu_item_safe",
    "TrackPopupMenu": "track_popup_menu_safe",
    "OpenClipboard": "open_clipboard_safe",
    "CloseClipboard": "close_clipboard_safe",
    "EmptyClipboard": "empty_clipboard_safe",
    "SetClipboardData": "set_clipboard_data_safe",
    "GetSaveFileNameW": "get_save_file_name_w_safe",
    "GetOpenFileNameW": "get_open_file_name_w_safe",
    "RegisterClassW": "register_class_w_safe",
    "GetLastError": "get_last_error_safe",
    "LoadCursorW": "load_cursor_w_safe",
    "CreateWindowExW": "create_window_ex_w_safe",
    "MoveWindow": "move_window_safe",
    "SetTimer": "set_timer_safe",
    "SetWindowPos": "set_window_pos_safe",
    "InvalidateRect": "invalidate_rect_safe",
    "TranslateMessage": "translate_message_safe",
    "DispatchMessageW": "dispatch_message_w_safe",
    "GetMessageW": "get_message_w_safe",
    # CreatePopupMenu omitted: helper returns HMENU, raw often used as Result via ?
    "GetStockObject": "get_stock_object_safe",
}

# Patterns that mean a block still needs unsafe (conservative).
NEEDS_UNSAFE_RES = [
    re.compile(r"\bas\s*\*const\b"),
    re.compile(r"\bas\s*\*mut\b"),
    re.compile(r"\btransmute\b"),
    re.compile(r"\bfrom_raw_parts(?:_mut)?\b"),
    re.compile(r"\bBox::from_raw\b"),
    re.compile(r"\bBox::into_raw\b"),
    re.compile(r"\bstd::ptr::"),
    re.compile(r"\bcore::ptr::"),
    re.compile(r"\bptr::(?:copy|read|write|null)"),
    re.compile(r"\bmem::(?:zeroed|transmute|forget)\b"),
    re.compile(r"\bstd::mem::zeroed\b"),
    re.compile(r"\bfrom_raw\b"),
    re.compile(r"\bread_unaligned\b"),
    re.compile(r"\bwrite_unaligned\b"),
    re.compile(r"\bcopy_nonoverlapping\b"),
    re.compile(r"\bGlobalLock\b"),
    re.compile(r"\bGlobalAlloc\b"),
    re.compile(r"\bGlobalUnlock\b"),
    re.compile(r"\bCoTaskMemFree\b"),
    re.compile(r"\bCoInitialize"),
    re.compile(r"\bCoUninitialize\b"),
    re.compile(r"\bCoCreateInstance\b"),
    re.compile(r"\blpVtbl\b"),
    re.compile(r"\bav_\w+\b"),
    re.compile(r"\bswr_\w+\b"),
    re.compile(r"\bBASS_"),
    re.compile(r"\blibloading\b"),
    re.compile(r"\bGetBuffer\b"),
    re.compile(r"\bReleaseBuffer\b"),
    re.compile(r"\bGetMixFormat\b"),
    re.compile(r"\bPropVariant"),
    re.compile(r"\bSHGet"),
    re.compile(r"\bSHStrDup"),
    re.compile(r"\bCLSIDFrom"),
    re.compile(r"\b\.Invoke\b"),
    re.compile(r"\bGetIDsOfNames\b"),
    re.compile(r"\bstd::slice::from_raw"),
    re.compile(r"\bcore::slice::from_raw"),
    re.compile(r"\bunsafe\b"),  # nested unsafe keeps outer
    # Common remaining raw Win32 not yet wrapped:
    re.compile(r"\bCreateWindowEx[AW]?\b"),
    re.compile(r"\bSendMessage[AW]?\b"),
    re.compile(r"\bPostMessage[AW]?\b"),
    re.compile(r"\bDefWindowProc[AW]?\b"),
    re.compile(r"\bCallWindowProc[AW]?\b"),
    re.compile(r"\bSetWindowLongPtr[AW]?\b"),
    re.compile(r"\bGetWindowLongPtr[AW]?\b"),
    re.compile(r"\bDestroyWindow\b"),
    re.compile(r"\bSetFocus\b"),
    re.compile(r"\bGetFocus\b"),
    re.compile(r"\bShowWindow\b"),
    re.compile(r"\bEnableWindow\b"),
    re.compile(r"\bSetWindowText[AW]?\b"),
    re.compile(r"\bGetWindowText[AW]?\b"),
    re.compile(r"\bMessageBox[AW]?\b"),
    re.compile(r"\bMoveWindow\b"),
    re.compile(r"\bSetTimer\b"),
    re.compile(r"\bKillTimer\b"),
    re.compile(r"\bSetWindowPos\b"),
    re.compile(r"\bInvalidateRect\b"),
    re.compile(r"\bBeginPaint\b"),
    re.compile(r"\bEndPaint\b"),
    re.compile(r"\bGetDC\b"),
    re.compile(r"\bReleaseDC\b"),
    re.compile(r"\bFillRect\b"),
    re.compile(r"\bSelectObject\b"),
    re.compile(r"\bDeleteObject\b"),
    re.compile(r"\bCreateFont"),
    re.compile(r"\bLoadCursor[AW]?\b"),
    re.compile(r"\bRegisterClass[AW]?\b"),
    re.compile(r"\bGetMessage[AW]?\b"),
    re.compile(r"\bPeekMessage[AW]?\b"),
    re.compile(r"\bTranslateMessage\b"),
    re.compile(r"\bDispatchMessage[AW]?\b"),
    re.compile(r"\bIsWindow\b"),
    re.compile(r"\bIsChild\b"),
    re.compile(r"\bIsDialogMessage[AW]?\b"),
    re.compile(r"\bGetDlgItem\b"),
    re.compile(r"\bGetParent\b"),
    re.compile(r"\bGetClientRect\b"),
    re.compile(r"\bGetWindowRect\b"),
    re.compile(r"\bGetCursorPos\b"),
    re.compile(r"\bGetKeyState\b"),
    re.compile(r"\bGetMenu\b"),
    re.compile(r"\bAppendMenu[AW]?\b"),
    re.compile(r"\bTrackPopupMenu\b"),
    re.compile(r"\bCreatePopupMenu\b"),
    re.compile(r"\bCreateMenu\b"),
    re.compile(r"\bDestroyMenu\b"),
    re.compile(r"\bOpenClipboard\b"),
    re.compile(r"\bCloseClipboard\b"),
    re.compile(r"\bEmptyClipboard\b"),
    re.compile(r"\bSetClipboardData\b"),
    re.compile(r"\bGetOpenFileName[AW]?\b"),
    re.compile(r"\bGetSaveFileName[AW]?\b"),
    re.compile(r"\bDragQueryFile[AW]?\b"),
    re.compile(r"\bShellExecute[AW]?\b"),
    re.compile(r"\bGetModuleHandle[AW]?\b"),
    re.compile(r"\bLoadLibrary[AW]?\b"),
    re.compile(r"\bGetProcAddress\b"),
    re.compile(r"\bCreateThread\b"),
    re.compile(r"\bWaitForSingleObject\b"),
    re.compile(r"\bSetEvent\b"),
    re.compile(r"\bResetEvent\b"),
    re.compile(r"\bCreateEvent"),
    re.compile(r"\bCloseHandle\b"),
    re.compile(r"\bMapViewOfFile\b"),
    re.compile(r"\bUnmapViewOfFile\b"),
    re.compile(r"\bCreateFileMapping"),
    re.compile(r"\bOpenProcess\b"),
    re.compile(r"\bTerminateProcess\b"),
    re.compile(r"\bAttachThreadInput\b"),
    re.compile(r"\bSystemParametersInfo"),
    re.compile(r"\bGetSystemMetrics\b"),
    re.compile(r"\bMonitorFrom"),
    re.compile(r"\bGetMonitorInfo"),
    re.compile(r"\bEnumWindows\b"),
    re.compile(r"\bEnumChildWindows\b"),
    re.compile(r"\bFindWindow[AW]?\b"),
    re.compile(r"\bGetForegroundWindow\b"),
    re.compile(r"\bSetForegroundWindow\b"),
    re.compile(r"\bBringWindowToTop\b"),
    re.compile(r"\bAllowSetForegroundWindow\b"),
    re.compile(r"\bGetWindowThreadProcessId\b"),
    re.compile(r"\bGetCurrentThreadId\b"),
    re.compile(r"\bSendInput\b"),
    re.compile(r"\bkeybd_event\b"),
    re.compile(r"\bMapVirtualKey"),
    re.compile(r"\bToUnicode\b"),
    re.compile(r"\bGetAsyncKeyState\b"),
    re.compile(r"\bRegisterHotKey\b"),
    re.compile(r"\bUnregisterHotKey\b"),
    re.compile(r"\bCreateAcceleratorTable"),
    re.compile(r"\bTranslateAccelerator"),
    re.compile(r"\bDrawText"),
    re.compile(r"\bTextOut"),
    re.compile(r"\bBitBlt\b"),
    re.compile(r"\bStretchBlt\b"),
    re.compile(r"\bGetStockObject\b"),
    re.compile(r"\bSetBkMode\b"),
    re.compile(r"\bSetTextColor\b"),
    re.compile(r"\bCreateSolidBrush\b"),
    re.compile(r"\bCreatePen\b"),
    re.compile(r"\bRectangle\b"),
    re.compile(r"\bEllipse\b"),
    re.compile(r"\bLineTo\b"),
    re.compile(r"\bMoveToEx\b"),
    re.compile(r"\bGetObjectW\b"),
    re.compile(r"\bGetDeviceCaps\b"),
    re.compile(r"\bStartDoc"),
    re.compile(r"\bEndDoc\b"),
    re.compile(r"\bStartPage\b"),
    re.compile(r"\bEndPage\b"),
    re.compile(r"\bAbortDoc\b"),
    re.compile(r"\bPrintDlg"),
    re.compile(r"\bPageSetupDlg"),
    re.compile(r"\bCommDlgExtendedError\b"),
    re.compile(r"\bChooseFont"),
    re.compile(r"\bChooseColor"),
    re.compile(r"\bGetOpenFileName"),
    re.compile(r"\bImmGet"),
    re.compile(r"\bImmSet"),
    re.compile(r"\bOleInitialize\b"),
    re.compile(r"\bOleUninitialize\b"),
    re.compile(r"\bVariantClear\b"),
    re.compile(r"\bSysAllocString\b"),
    re.compile(r"\bSysFreeString\b"),
    re.compile(r"\*\s*[a-zA-Z_(]"),  # pointer deref (may false-positive multiply - keep conservative)
]

SAFE_FN_RE = re.compile(
    r"\b(?:pub(?:\s*\([^)]*\))?\s+)?(?:async\s+)?(?:unsafe\s+)?fn\s+(\w+_safe)\s*[<(]"
)


def find_matching(s: str, open_idx: int, open_ch: str, close_ch: str) -> int | None:
    depth = 0
    i = open_idx
    n = len(s)
    in_str = None
    while i < n:
        ch = s[i]
        if in_str:
            if ch == "\\" and i + 1 < n:
                i += 2
                continue
            if ch == in_str:
                in_str = None
            i += 1
            continue
        if ch in ("'", '"'):
            in_str = ch
            i += 1
            continue
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return None


def safe_fn_ranges(text: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    for m in SAFE_FN_RE.finditer(text):
        brace = text.find("{", m.start())
        if brace < 0:
            continue
        semi = text.find(";", m.start())
        if 0 <= semi < brace:
            continue
        end = find_matching(text, brace, "{", "}")
        if end is None:
            continue
        ranges.append((brace, end + 1))
    return ranges


def in_ranges(pos: int, ranges: list[tuple[int, int]]) -> bool:
    for a, b in ranges:
        if a <= pos < b:
            return True
    return False


def use_import_ranges(text: str) -> list[tuple[int, int]]:
    ranges = []
    for m in re.finditer(r"\buse\s+", text):
        # find semicolon end
        semi = text.find(";", m.start())
        if semi < 0:
            continue
        ranges.append((m.start(), semi + 1))
    return ranges


def helper_name(api: str, in_main: bool) -> str:
    base = API_TO_HELPER[api]
    return base if in_main else f"crate::{base}"


def rewrite_calls(text: str, in_main: bool) -> tuple[str, int]:
    skip = safe_fn_ranges(text) if in_main else []
    skip += use_import_ranges(text)
    count = 0
    # Special: IsWindow(...).as_bool() -> is_window_handle_valid(...)
    # and IsChild(...).as_bool() -> is_child_safe(...)
    specials = [
        ("IsWindow", "is_window_handle_valid"),
        ("IsChild", "is_child_safe"),
    ]
    out = text
    for api, helper in specials:
        h = helper if in_main else f"crate::{helper}"
        pat = re.compile(rf"\b{api}\s*\(([\s\S]*?)\)\s*\.\s*as_bool\s*\(\s*\)")

        def repl_special(m: re.Match[str], h=h, api=api) -> str:
            nonlocal count
            if in_ranges(m.start(), skip):
                return m.group(0)
            # avoid rewriting inside helper names
            count += 1
            return f"{h}({m.group(1)})"

        out = pat.sub(repl_special, out)

    # General APIs: \bAPI(
    # Process longer names first
    for api in sorted(API_TO_HELPER.keys(), key=len, reverse=True):
        h = helper_name(api, in_main)
        i = 0
        parts: list[str] = []
        last = 0
        s = out
        while True:
            m = re.search(rf"\b{api}\s*\(", s[i:])
            if not m:
                break
            start = i + m.start()
            # skip if already a helper call like send_message_w_safe - API names don't overlap
            # skip if part of larger ident already handled by \b
            # skip use / safe fn bodies
            if in_ranges(start, skip):
                i = start + len(api)
                continue
            # Never rewrite qualified paths: windows::...::CreateWindowExW
            if start > 0 and s[start - 1] == ':':
                i = start + len(api)
                continue
            # skip `fn SendMessageW` - rare
            # skip if this is inside a string - basic check of quotes balance before
            # Avoid double: if already helper_name before
            before = s[max(0, start - 40) : start]
            if before.rstrip().endswith(h) or before.rstrip().endswith(API_TO_HELPER[api]):
                i = start + len(api)
                continue
            # if previous chars form ident with _safe suffix calling - fine
            paren = start + m.end() - i - 1 + i  # absolute index of '('
            # m.end() is relative to s[i:], '(' is at i+m.end()-1
            paren = i + m.end() - 1
            close = find_matching(s, paren, "(", ")")
            if close is None:
                i = start + len(api)
                continue
            # Don't rewrite type references in function pointers etc. - if followed by -> weird
            # Replace API( with helper(
            # Keep args as-is
            parts.append(s[last:start])
            parts.append(h + s[paren : close + 1])
            count += 1
            last = close + 1
            i = last
        parts.append(s[last:])
        out = "".join(parts)
        # refresh skip ranges for main after mutations? use statements positions shift.
        # Recompute skip on each API for correctness.
        if in_main:
            skip = safe_fn_ranges(out) + use_import_ranges(out)
        else:
            skip = use_import_ranges(out)
    return out, count


def demote_panic_guard_wrappers(text: str) -> tuple[str, int]:
    """Remove outer unsafe around panic_guard::guard when body is only that call."""
    count = 0
    i = 0
    parts: list[str] = []
    last = 0
    s = text
    n = len(s)
    while i < n:
        if s.startswith("unsafe", i) and (
            i == 0 or not (s[i - 1].isalnum() or s[i - 1] == "_")
        ):
            j = i + len("unsafe")
            while j < n and s[j] in " \t\r\n":
                j += 1
            if j < n and s[j] == "{":
                close = find_matching(s, j, "{", "}")
                if close is not None:
                    body = s[j + 1 : close].strip()
                    # body should be panic_guard::guard( ... ) optionally with crate::
                    if re.match(
                        r"^(?:crate::)?panic_guard::guard\s*\(",
                        body,
                        re.S,
                    ):
                        # ensure only one top-level call
                        # find matching paren of guard(
                        gp = body.find("(")
                        gclose = find_matching(body, gp, "(", ")")
                        if gclose is not None and body[gclose + 1 :].strip() == "":
                            # If DefWindowProcW remains inside, still demote only if rewritten
                            if "DefWindowProcW" not in body and not re.search(
                                r"\bunsafe\b", body
                            ):
                                parts.append(s[last:i])
                                parts.append(body)
                                count += 1
                                i = close + 1
                                last = i
                                continue
        i += 1
    parts.append(s[last:])
    return "".join(parts), count


def body_needs_unsafe(body: str) -> bool:
    # Strip line comments and strings roughly before checking
    cleaned = re.sub(r"//[^\n]*", "", body)
    cleaned = re.sub(r"/\*.*?\*/", "", cleaned, flags=re.S)
    cleaned = re.sub(r'"(?:\\.|[^"\\])*"', '""', cleaned)
    cleaned = re.sub(r"'(?:\\.|[^'\\])*'", "''", cleaned)
    for rx in NEEDS_UNSAFE_RES:
        if rx.search(cleaned):
            return True
    return False


def demote_safe_unsafe_blocks(text: str) -> tuple[str, int]:
    count = 0
    i = 0
    parts: list[str] = []
    last = 0
    s = text
    n = len(s)
    while i < n:
        if s.startswith("unsafe", i) and (
            i == 0 or not (s[i - 1].isalnum() or s[i - 1] == "_")
        ):
            # don't touch `unsafe fn` / `unsafe impl` / `unsafe trait` / `unsafe extern`
            j = i + len("unsafe")
            while j < n and s[j] in " \t\r\n":
                j += 1
            if j < n and s[j] == "{":
                close = find_matching(s, j, "{", "}")
                if close is not None:
                    body = s[j + 1 : close]
                    if not body_needs_unsafe(body):
                        # keep braces (block may have multiple stmts / scope)
                        parts.append(s[last:i])
                        parts.append("{" + body + "}")
                        count += 1
                        i = close + 1
                        last = i
                        continue
        i += 1
    parts.append(s[last:])
    return "".join(parts), count


def process_file(path: Path) -> dict[str, int]:
    raw = path.read_bytes()
    text = raw.decode("utf-8")
    in_main = path.resolve() == (ROOT / "main.rs").resolve()
    stats = {"calls": 0, "panic_guard": 0, "demote": 0}

    text, n = rewrite_calls(text, in_main)
    stats["calls"] = n
    text, n = demote_panic_guard_wrappers(text)
    stats["panic_guard"] = n
    # General demotion is intentionally disabled: too many false positives
    # (unsafe fn pointers / COM / multi-API blocks). Re-enable only with
    # a proven allowlist.
    stats["demote"] = 0

    if text.encode("utf-8") != raw:
        path.write_bytes(text.encode("utf-8"))
    return stats


def main() -> int:
    totals = {"calls": 0, "panic_guard": 0, "demote": 0}
    files = 0
    for path in sorted(ROOT.rglob("*.rs")):
        if path.name.endswith(".bac"):
            continue
        st = process_file(path)
        if any(st.values()):
            files += 1
            for k, v in st.items():
                totals[k] += v
            print(
                f"{st['calls']:4d} calls, {st['panic_guard']:3d} guard, {st['demote']:3d} demote  {path.relative_to(ROOT.parent)}"
            )
    print(
        f"---\nfiles={files} calls={totals['calls']} panic_guard={totals['panic_guard']} demote={totals['demote']}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
