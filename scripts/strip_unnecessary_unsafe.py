#!/usr/bin/env python3
"""Strip `unsafe` keywords that rustc reports as unnecessary blocks."""

from __future__ import annotations

import re
import subprocess
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


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
        if ch in ("'", '"'):
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


def main() -> int:
    r = subprocess.run(
        ["cargo", "check"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    text = r.stderr + "\n" + r.stdout
    # error: unnecessary `unsafe` block\n  --> path:line:col
    locs = re.findall(
        r"unnecessary `unsafe` block\n\s*--> ([^:\n]+):(\d+):",
        text,
    )
    print(f"found {len(locs)} unnecessary unsafe blocks")
    by_file: dict[Path, list[int]] = defaultdict(list)
    for f, line in locs:
        p = Path(f)
        if not p.is_absolute():
            p = ROOT / p
        by_file[p].append(int(line))

    total = 0
    for path, lines in by_file.items():
        content = path.read_text(encoding="utf-8")
        # Convert line numbers to character offsets of "unsafe" at start of that line
        line_starts = [0]
        for i, ch in enumerate(content):
            if ch == "\n":
                line_starts.append(i + 1)
        # Process from end so offsets stay valid
        for line_no in sorted(set(lines), reverse=True):
            if line_no < 1 or line_no > len(line_starts):
                continue
            start = line_starts[line_no - 1]
            # find "unsafe" near start of line (allow indent)
            m = re.match(r"([ \t]*)unsafe\b", content[start : start + 80])
            if not m:
                # sometimes points at the keyword mid-line
                window = content[start : start + 200]
                m2 = re.search(r"\bunsafe\b", window)
                if not m2:
                    print(f"skip no unsafe at {path}:{line_no}")
                    continue
                abs_u = start + m2.start()
            else:
                abs_u = start + m.start(0) + len(m.group(1))
            # ensure this is block form: unsafe {
            j = abs_u + len("unsafe")
            while j < len(content) and content[j] in " \t\r\n":
                j += 1
            if j >= len(content) or content[j] != "{":
                print(f"skip not block {path}:{line_no}")
                continue
            # Remove only the keyword "unsafe" and following whitespace before {
            # Keep the brace block: "unsafe { ... }" -> "{ ... }"
            k = abs_u + len("unsafe")
            while k < j:
                k += 1
            # remove unsafe and spaces/newlines between unsafe and {
            # But preserve a space if needed? No, `{` can follow directly or with newline.
            # Prefer keeping original whitespace after keyword removal except keyword itself.
            # Remove "unsafe" only; leave one space if there was space before {
            content = content[:abs_u] + content[abs_u + len("unsafe") :].lstrip(" \t")
            # If next is newline, fine; if `{` fine.
            total += 1
        path.write_text(content, encoding="utf-8")
        print(f"{path}: stripped some")
    print(f"total stripped={total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
