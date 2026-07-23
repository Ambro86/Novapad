#!/usr/bin/env python3
"""Mechanically replace single-API unsafe { Win32(...) } call sites with *_safe helpers.

Only rewrites blocks whose body is exactly one of the known APIs (optional trailing
semicolon, optional .as_bool() / .ok() for a few cases). Does not touch multi-statement
unsafe blocks or FFmpeg/COM method calls.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "src"

# Map Win32 fn name -> (helper name, postprocess kind)
# postprocess:
#   None - plain call
#   "as_bool" - unsafe { Foo(args).as_bool() } -> helper(args)
#   "is_ok" - unsafe { Foo(args).is_ok() } -> helper(args)  (if helper returns bool differently, skip)
API_MAP = {
    "SendMessageW": ("send_message_w_safe", None),
    "PostMessageW": ("post_message_w_safe", None),
    "IsWindow": ("is_window_handle_valid", "as_bool"),
    "DestroyWindow": ("destroy_window_safe", None),
    "SetFocus": ("set_focus_safe", None),
    "GetFocus": ("get_focus_safe", None),
    "SetForegroundWindow": ("set_foreground_window_safe", None),
    "GetForegroundWindow": ("get_foreground_window_safe", None),
    "ShowWindow": ("show_window_safe", None),
    "EnableWindow": ("enable_window_safe", None),
    "SetWindowTextW": ("set_window_text_w_safe", None),
    "GetWindowTextW": ("get_window_text_w_safe", None),
    "GetWindowTextLengthW": ("get_window_text_length_w_safe", None),
    "GetWindowLongPtrW": ("get_window_long_ptr_w_safe", None),
    "SetWindowLongPtrW": ("set_window_long_ptr_w_safe", None),
    "DefWindowProcW": ("def_window_proc_w_safe", None),
    "CallWindowProcW": ("call_window_proc_w_safe", None),
    "GetParent": ("get_parent_safe", None),
    "GetDlgItem": ("get_dlg_item_safe", None),
    "IsChild": ("is_child_safe", "as_bool"),
    "KillTimer": ("kill_timer_safe", None),
    "MessageBoxW": ("message_box_w_safe", None),
    "GetClientRect": ("get_client_rect_safe", None),
    "GetWindowRect": ("get_window_rect_safe", None),
    "GetKeyState": ("get_key_state_safe", None),
    "GetMenu": ("get_menu_safe", None),
    "IsDialogMessageW": ("is_dialog_message_w_safe", None),
    "GetClassNameW": ("get_class_name_w_safe", None),
    "GetNextDlgTabItem": ("get_next_dlg_tab_item_safe", None),
    "FindWindowW": ("find_window_w_safe", None),
    "GetCursorPos": ("get_cursor_pos_safe", None),
    "CreateMenu": ("create_menu_safe", None),
    "DestroyMenu": ("destroy_menu_safe", None),
    # CreatePopupMenu omitted: helper returns HMENU, raw API returns Result<HMENU>
    "AppendMenuW": ("append_menu_w_safe", None),
    "CheckMenuItem": ("check_menu_item_safe", None),
    "TrackPopupMenu": ("track_popup_menu_safe", None),
    "OpenClipboard": ("open_clipboard_safe", None),
    "CloseClipboard": ("close_clipboard_safe", None),
    "EmptyClipboard": ("empty_clipboard_safe", None),
    "SetClipboardData": ("set_clipboard_data_safe", None),
    "GetSaveFileNameW": ("get_save_file_name_w_safe", None),
    "GetOpenFileNameW": ("get_open_file_name_w_safe", None),
    "RegisterClassW": ("register_class_w_safe", None),
    "GetLastError": ("get_last_error_safe", None),
    "CreateWindowExW": ("create_window_ex_w_safe", None),
    "MoveWindow": ("move_window_safe", None),
    "SetTimer": ("set_timer_safe", None),
    "SetWindowPos": ("set_window_pos_safe", None),
    "InvalidateRect": ("invalidate_rect_safe", None),
    "TranslateMessage": ("translate_message_safe", None),
    "DispatchMessageW": ("dispatch_message_w_safe", None),
    "LoadCursorW": ("load_cursor_w_safe", None),
}

# Box::from_raw is special (path with ::)
SPECIAL = {
    "Box::from_raw": ("box_from_raw_safe", None),
}

API_NAMES = sorted(API_MAP.keys(), key=len, reverse=True)


def find_matching_paren(s: str, open_idx: int) -> int | None:
    """open_idx points at '('; return index of matching ')'."""
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
        if ch in ('"', "'"):
            in_str = ch
            i += 1
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return None


def find_matching_brace(s: str, open_idx: int) -> int | None:
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
        if ch in ('"', "'"):
            in_str = ch
            i += 1
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return None


def helper_ref(helper: str, in_main: bool) -> str:
    if in_main:
        return helper
    return f"crate::{helper}"


def try_replace_block(body: str, in_main: bool) -> str | None:
    """If body is a single known API call (+ optional ; / .as_bool()), return replacement."""
    body = body.strip()
    if not body:
        return None

    # Drop trailing semicolon inside the block; caller decides if a statement
    # terminator is needed so we don't break `|x| unsafe { Foo(); }` closures.
    if body.endswith(";"):
        body = body[:-1].rstrip()

    # Special: Box::from_raw(...)
    if body.startswith("Box::from_raw"):
        rest = body[len("Box::from_raw") :].lstrip()
        if not rest.startswith("("):
            return None
        close = find_matching_paren(rest, 0)
        if close is None:
            return None
        args = rest[1:close]
        after = rest[close + 1 :].strip()
        if after:
            return None
        h = helper_ref("box_from_raw_safe", in_main)
        return f"{h}({args})"

    # Match API name at start
    for api in API_NAMES:
        if not body.startswith(api):
            continue
        # ensure not prefix of longer ident
        if len(body) > len(api) and (body[len(api)].isalnum() or body[len(api)] == "_"):
            continue
        rest = body[len(api) :].lstrip()
        if not rest.startswith("("):
            return None
        close = find_matching_paren(rest, 0)
        if close is None:
            return None
        args = rest[1:close]
        after = rest[close + 1 :].strip()

        helper, kind = API_MAP[api]
        h = helper_ref(helper, in_main)

        if kind == "as_bool":
            # allow .as_bool() only
            if after in (".as_bool()",):
                return f"{h}({args})"
            if after == "":
                # IsWindow without as_bool - helper already returns bool for IsWindow
                # but is_window_handle_valid returns bool; raw IsWindow returns BOOL
                # Don't replace bare IsWindow without as_bool to avoid type change
                if api in ("IsWindow", "IsChild"):
                    return None
                return f"{h}({args})"
            return None

        if after != "":
            # e.g. chained methods - skip (COM-like)
            return None

        return f"{h}({args})"

    return None


SAFE_FN_RE = re.compile(
    r"\b(?:pub(?:\s*\([^)]*\))?\s+)?(?:async\s+)?(?:unsafe\s+)?fn\s+(\w+_safe)\s*[<(]"
)


def safe_fn_body_ranges(text: str) -> list[tuple[int, int]]:
    """Return [start, end) ranges of function bodies for * _safe helpers (skip rewrites)."""
    ranges: list[tuple[int, int]] = []
    for m in SAFE_FN_RE.finditer(text):
        # find opening brace of body after signature
        k = m.end() - 1
        # walk to '{' before any nested, handling where generics/args end
        brace = text.find("{", m.start())
        if brace < 0:
            continue
        # ensure this brace belongs to the fn: no ';' between fn name and brace (trait methods)
        semi = text.find(";", m.start())
        if 0 <= semi < brace:
            continue
        end = find_matching_brace(text, brace)
        if end is None:
            continue
        ranges.append((brace, end + 1))
    return ranges


def in_ranges(pos: int, ranges: list[tuple[int, int]]) -> bool:
    for a, b in ranges:
        if a <= pos < b:
            return True
    return False


def process_file(path: Path) -> tuple[int, str]:
    raw = path.read_bytes()
    # preserve original newline style by operating on text decoded as-is
    text = raw.decode("utf-8")
    in_main = path.resolve() == (ROOT / "main.rs").resolve()
    skip_ranges = safe_fn_body_ranges(text) if in_main else []

    count = 0
    out_parts: list[str] = []
    i = 0
    n = len(text)
    last = 0

    while i < n:
        if text.startswith("unsafe", i) and (
            i == 0 or not (text[i - 1].isalnum() or text[i - 1] == "_")
        ):
            j = i + len("unsafe")
            while j < n and text[j] in " \t\r\n":
                j += 1
            if j < n and text[j] == "{":
                close = find_matching_brace(text, j)
                if close is not None:
                    # Never rewrite unsafe inside *_safe helper definitions in main.rs
                    if in_ranges(i, skip_ranges):
                        i = close + 1
                        continue
                    body = text[j + 1 : close]
                    repl = try_replace_block(body, in_main)
                    if repl is not None:
                        # If the original block was a bare statement and the next
                        # token starts another statement, keep a trailing ';'.
                        k = close + 1
                        while k < n and text[k] in " \t\r\n":
                            k += 1
                        needs_semi = False
                        if body.strip().endswith(";"):
                            # Keep ';' for statement form (incl. before '}').
                            # Skip only when clearly inside a larger expression.
                            if k >= n or text[k] not in ",)].?;":
                                needs_semi = True
                        if needs_semi and not repl.endswith(";"):
                            repl = repl + ";"
                        out_parts.append(text[last:i])
                        out_parts.append(repl)
                        count += 1
                        i = close + 1
                        last = i
                        continue
        i += 1

    out_parts.append(text[last:])
    return count, "".join(out_parts)


def main() -> int:
    total = 0
    files_changed = 0
    for path in sorted(ROOT.rglob("*.rs")):
        if path.name.endswith(".bac"):
            continue
        n, new_text = process_file(path)
        if n > 0:
            old = path.read_text(encoding="utf-8")
            if new_text != old:
                # write bytes preserving whatever newlines the joined string has
                path.write_bytes(new_text.encode("utf-8"))
                files_changed += 1
                total += n
                print(f"{n:4d}  {path.relative_to(ROOT.parent)}")
    print(f"---\nReplaced {total} unsafe blocks in {files_changed} files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
