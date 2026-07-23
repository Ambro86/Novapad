#!/usr/bin/env python3
"""Remove unused imports reported by cargo check until clean or stuck."""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def cargo_errors() -> str:
    r = subprocess.run(
        ["cargo", "check"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return r.stderr + "\n" + r.stdout


def parse_unused(text: str) -> list[tuple[Path, int, list[str]]]:
    # Multi-line messages: names may continue after comma on next diagnostic lines
    items: list[tuple[Path, int, list[str]]] = []
    # Split on error: unused
    for m in re.finditer(
        r"error: unused imports?: (.*?)\n\s*--> ([^:\n]+):(\d+):",
        text,
        re.S,
    ):
        blob = m.group(1)
        path = Path(m.group(2).strip())
        line = int(m.group(3))
        names = re.findall(r"`([A-Za-z0-9_:]+)`", blob)
        if names:
            items.append((path, line, names))
    return items


def remove_from_use_block(content: str, line_1based: int, names: list[str]) -> str:
    lines = content.splitlines(True)
    idx = max(0, line_1based - 1)
    start = idx
    while start > 0 and "use " not in lines[start]:
        start -= 1
    end = start
    while end < len(lines) and ";" not in lines[end]:
        end += 1
    if end >= len(lines):
        return content
    block = "".join(lines[start : end + 1])
    orig = block
    for name in names:
        simple = name.split("::")[-1]
        block = re.sub(rf"use\s+{re.escape(name)}\s*;\s*\n?", "", block)
        block = re.sub(rf",\s*{re.escape(simple)}\b", "", block)
        block = re.sub(rf"\b{re.escape(simple)}\s*,\s*", "", block)
        block = re.sub(rf"\{{\s*{re.escape(simple)}\s*\}}", "{ }", block)
    block = re.sub(r"use\s+[A-Za-z0-9_:]+::\s*\{\s*\}\s*;\s*\n?", "", block)
    block = re.sub(r",\s*,", ",", block)
    block = re.sub(r",\s*\}", "}", block)
    block = re.sub(r"\{\s*,", "{", block)
    block = re.sub(r"use\s+[A-Za-z0-9_:]+::\s*\{\s*\}\s*;\s*\n?", "", block)
    if block == orig:
        return content
    return "".join(lines[:start] + [block] + lines[end + 1 :])


def main() -> int:
    for round_i in range(12):
        text = cargo_errors()
        if "Finished" in text and "error" not in text.split("Finished")[0]:
            # may still have errors after Finished? no
            pass
        items = parse_unused(text)
        e_count = len(re.findall(r"^error", text, re.M))
        print(f"round {round_i}: errors={e_count} unused_items={len(items)}")
        if not items:
            if e_count == 0 or "Finished" in text and e_count == 0:
                print("clean")
                return 0
            # print remaining
            for line in text.splitlines():
                if line.startswith("error"):
                    print(line)
            return 1 if e_count else 0
        changed = 0
        # group by file
        by_file: dict[Path, list[tuple[int, list[str]]]] = {}
        for path, line, names in items:
            if not path.is_absolute():
                path = ROOT / path
            by_file.setdefault(path, []).append((line, names))
        for path, lst in by_file.items():
            if not path.exists():
                print("missing", path)
                continue
            content = path.read_text(encoding="utf-8")
            # apply from bottom to top so line numbers stay valid-ish; re-read each time
            for line, names in sorted(lst, key=lambda x: -x[0]):
                new_content = remove_from_use_block(content, line, names)
                if new_content != content:
                    content = new_content
                    changed += 1
            path.write_text(content, encoding="utf-8")
        if changed == 0:
            print("stuck; remaining:")
            for line in text.splitlines():
                if line.startswith("error"):
                    print(line)
            return 1
    return 1


if __name__ == "__main__":
    sys.exit(main())
