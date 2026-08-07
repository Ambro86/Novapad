#!/usr/bin/env python3
"""Archive-level regression tests for Sonarpad's conservative EPUB editing model.

The script applies four representative text-only edits to every supplied EPUB,
then verifies that package structure, markup, links, notes, images, CSS, metadata,
and reading order are unchanged. Only one selected XHTML resource may differ.
"""

from __future__ import annotations

import argparse
import copy
import html
import json
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

BLOCK_TAGS = {
    "p", "div", "li", "tr", "hr", "ul", "ol", "table",
    "h1", "h2", "h3", "h4", "h5", "h6",
}
SKIP_TAGS = {"head", "style", "script", "title"}
XML_EXTENSIONS = {".xml", ".opf", ".ncx", ".xhtml", ".html", ".htm", ".svg"}
MARKUP_RE = re.compile(r"<[^>]+>", re.DOTALL)
ATTR_RE = re.compile(
    r"\b(?:id|xml:id|name|href|src|xlink:href|epub:type)\s*=\s*(['\"])(.*?)\1",
    re.IGNORECASE | re.DOTALL,
)
NOTE_RE = re.compile(
    rb"(?:footnote|endnote|noteref|doc-footnote|doc-noteref|nota)", re.IGNORECASE
)


@dataclass(frozen=True)
class NodeChar:
    node_index: int
    char_index: int


@dataclass(frozen=True)
class BreakTag:
    source_start: int
    source_end: int
    preserve_empty: bool = False


@dataclass(frozen=True)
class Structural:
    pass


STRUCTURAL = Structural()
Source = NodeChar | BreakTag | Structural


@dataclass
class TextNode:
    source_start: int
    source_end: int
    chars: list[str]


@dataclass
class Extraction:
    chars: list[str]
    sources: list[Source]
    nodes: list[TextNode]


@dataclass
class Resource:
    archive_path: str
    html: str
    filtered_chars: list[str]
    filtered_sources: list[Source]
    nodes: list[TextNode]
    full_start: int


@dataclass
class BookModel:
    title: str
    full_text: str
    resources: list[Resource]
    opf_path: str
    spine_paths: list[str]


def decode_entity(entity: str, include_semicolon: bool = True) -> list[str]:
    known = {
        "nbsp": " ",
        "lt": "<",
        "gt": ">",
        "amp": "&",
        "quot": '"',
        "apos": "'",
    }
    if entity in known:
        return [known[entity]]
    try:
        if entity.lower().startswith("#x"):
            return [chr(int(entity[2:], 16))]
        if entity.startswith("#"):
            return [chr(int(entity[1:], 10))]
    except (ValueError, OverflowError):
        pass
    suffix = ";" if include_semicolon else ""
    return list(f"&{entity}{suffix}")


def extract_html(html_text: str) -> Extraction:
    chars: list[str] = []
    sources: list[Source] = []
    nodes: list[TextNode] = []
    node_start: int | None = None
    node_chars: list[str] = []
    inside_tag = False
    tag = ""
    tag_start = 0
    last_newline = False
    skip_stack: list[str] = []
    in_comment = False
    entity = ""
    in_entity = False

    def append_node_char(ch: str, raw_start: int) -> None:
        nonlocal node_start, last_newline
        if node_start is None:
            node_start = raw_start
        node_index = len(nodes)
        char_index = len(node_chars)
        node_chars.append(ch)
        chars.append(ch)
        sources.append(NodeChar(node_index, char_index))
        last_newline = ch == "\n"

    def flush_node(source_end: int) -> None:
        nonlocal node_start, node_chars
        if node_start is not None:
            nodes.append(TextNode(node_start, source_end, node_chars))
            node_start = None
            node_chars = []

    for byte_index, ch in enumerate_utf8_offsets(html_text):
        if in_comment:
            tag += ch
            if tag.endswith("-->"):
                in_comment = False
                tag = ""
            continue

        if inside_tag:
            if ch == ">":
                inside_tag = False
                source_end = byte_index + len(ch.encode("utf-8"))
                trimmed = tag.strip()
                if trimmed.startswith("!--"):
                    if not trimmed.endswith("--"):
                        in_comment = True
                    tag = ""
                    continue
                stripped = trimmed.lstrip("/").strip()
                tag_name = stripped.split()[0].rstrip("/").lower() if stripped else ""
                is_closing = trimmed.startswith("/")
                if tag_name in SKIP_TAGS:
                    if is_closing:
                        try:
                            position = len(skip_stack) - 1 - skip_stack[::-1].index(tag_name)
                            skip_stack = skip_stack[:position]
                        except ValueError:
                            pass
                    else:
                        skip_stack.append(tag_name)
                    tag = ""
                    continue
                class_match = re.search(r'''\bclass\s*=\s*(['\"])(.*?)\1''', trimmed, re.IGNORECASE | re.DOTALL)
                preserve_empty = bool(
                    tag_name == "br"
                    and class_match
                    and "sonarpad-preserve-line-break" in class_match.group(2).split()
                )
                if tag_name == "br" and not skip_stack and (preserve_empty or not last_newline) and chars:
                    chars.append("\n")
                    sources.append(BreakTag(tag_start, source_end, preserve_empty))
                    last_newline = True
                elif tag_name in BLOCK_TAGS and not skip_stack and not last_newline and chars:
                    chars.append("\n")
                    sources.append(STRUCTURAL)
                    last_newline = True
                tag = ""
            else:
                tag += ch
            continue

        if ch == "<":
            if in_entity:
                for decoded in decode_entity(entity, True):
                    append_node_char(decoded, byte_index)
                entity = ""
                in_entity = False
            flush_node(byte_index)
            inside_tag = True
            tag_start = byte_index
            continue
        if skip_stack:
            continue
        if in_entity:
            if ch == ";":
                for decoded in decode_entity(entity, True):
                    append_node_char(decoded, byte_index)
                entity = ""
                in_entity = False
            elif len(entity) < 16 and not ch.isspace():
                entity += ch
            else:
                append_node_char("&", byte_index)
                for entity_ch in entity:
                    append_node_char(entity_ch, byte_index)
                append_node_char(ch, byte_index)
                entity = ""
                in_entity = False
            continue
        if ch == "&":
            if node_start is None:
                node_start = byte_index
            in_entity = True
            entity = ""
            continue
        append_node_char(ch, byte_index)

    if in_entity:
        append_node_char("&", len(html_text.encode("utf-8")))
        for entity_ch in entity:
            append_node_char(entity_ch, len(html_text.encode("utf-8")))
    flush_node(len(html_text.encode("utf-8")))
    return Extraction(chars, sources, nodes)


def enumerate_utf8_offsets(value: str) -> Iterable[tuple[int, str]]:
    offset = 0
    for ch in value:
        yield offset, ch
        offset += len(ch.encode("utf-8"))


def filter_extraction(extraction: Extraction) -> tuple[list[str], list[Source]]:
    chars: list[str] = []
    sources: list[Source] = []
    line_start = 0
    while line_start < len(extraction.chars):
        try:
            newline_position = extraction.chars.index("\n", line_start)
        except ValueError:
            newline_position = None
        line_end = newline_position if newline_position is not None else len(extraction.chars)
        trimmed_start = line_start
        while trimmed_start < line_end and extraction.chars[trimmed_start].isspace():
            trimmed_start += 1
        trimmed_end = line_end
        while trimmed_end > trimmed_start and extraction.chars[trimmed_end - 1].isspace():
            trimmed_end -= 1
        newline_source = (
            extraction.sources[newline_position]
            if newline_position is not None
            else STRUCTURAL
        )
        if trimmed_start < trimmed_end:
            line = "".join(extraction.chars[trimmed_start:trimmed_end])
            normalized = " ".join(line.split())
            if normalized.lower() not in {"epub r1.0", "epub base r2.1"} and not (
                line.startswith("part") and len(line) <= 12
            ):
                chars.extend(extraction.chars[trimmed_start:trimmed_end])
                sources.extend(extraction.sources[trimmed_start:trimmed_end])
                chars.append("\n")
                sources.append(newline_source)
        elif isinstance(newline_source, BreakTag) and newline_source.preserve_empty:
            chars.append("\n")
            sources.append(newline_source)
        if newline_position is None:
            break
        line_start = newline_position + 1
    return chars, sources


def normalize_path(path: str) -> str:
    parts: list[str] = []
    for part in path.replace("\\", "/").split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].rsplit(":", 1)[-1]


def load_model(path: Path) -> BookModel:
    with zipfile.ZipFile(path) as archive:
        container = ET.fromstring(archive.read("META-INF/container.xml"))
        rootfile = next(
            element for element in container.iter() if local_name(element.tag) == "rootfile"
        )
        opf_path = rootfile.attrib["full-path"]
        package = ET.fromstring(archive.read(opf_path))
        base = posixpath.dirname(opf_path)
        title = ""
        for element in package.iter():
            if local_name(element.tag) == "title":
                title = "".join(element.itertext())
                break
        manifest: dict[str, tuple[str, str]] = {}
        for element in package.iter():
            if local_name(element.tag) == "item" and "id" in element.attrib:
                href = element.attrib.get("href", "")
                manifest[element.attrib["id"]] = (
                    normalize_path(posixpath.join(base, href)),
                    element.attrib.get("media-type", ""),
                )
        spine_ids = [
            element.attrib["idref"]
            for element in package.iter()
            if local_name(element.tag) == "itemref" and "idref" in element.attrib
        ]
        full = title + ("\n\n" if title else "")
        resources: list[Resource] = []
        spine_paths: list[str] = []
        for spine_id in spine_ids:
            resource_info = manifest.get(spine_id)
            if not resource_info:
                continue
            resource_path, mime = resource_info
            spine_paths.append(resource_path)
            if not any(token in mime for token in ("xhtml", "html", "xml")):
                continue
            data = archive.read(resource_path)
            html_text = data.decode("utf-8")
            extraction = extract_html(html_text)
            filtered_chars, filtered_sources = filter_extraction(extraction)
            if not "".join(filtered_chars).strip():
                continue
            full_start = len(full)
            full += "".join(filtered_chars) + "\n"
            resources.append(
                Resource(
                    archive_path=resource_path,
                    html=html_text,
                    filtered_chars=filtered_chars,
                    filtered_sources=filtered_sources,
                    nodes=extraction.nodes,
                    full_start=full_start,
                )
            )
    return BookModel(title, full, resources, opf_path, spine_paths)


def choose_edit_location(model: BookModel) -> tuple[Resource, TextNode, int, int, int]:
    for resource in model.resources:
        positions_by_node: dict[int, list[tuple[int, int]]] = {}
        for local_index, source in enumerate(resource.filtered_sources):
            if isinstance(source, NodeChar):
                positions_by_node.setdefault(source.node_index, []).append(
                    (local_index, source.char_index)
                )
        for node_index, positions in positions_by_node.items():
            node = resource.nodes[node_index]
            node_text = "".join(node.chars)
            words = list(re.finditer(r"[A-Za-zÀ-ÖØ-öø-ÿ]{4,}", node_text))
            if len(node_text) < 80 or len(words) < 3:
                continue
            middle = words[len(words) // 2]
            char_start, char_end = middle.span()
            local_lookup = {char_index: local_index for local_index, char_index in positions}
            if char_start not in local_lookup or char_end - 1 not in local_lookup:
                continue
            global_start = resource.full_start + local_lookup[char_start]
            global_end = resource.full_start + local_lookup[char_end - 1] + 1
            return resource, node, char_start, global_start, global_end
    raise RuntimeError("No safe editable text node found")


def escape_xml_char(ch: str) -> str:
    return {"&": "&amp;", "<": "&lt;", ">": "&gt;"}.get(ch, ch)


def apply_node_edits(node: TextNode, operations: list[tuple[str, int, str | int]]) -> str:
    chars = list(node.chars)
    for operation, position, payload in sorted(operations, key=lambda item: item[1], reverse=True):
        if operation == "insert":
            assert isinstance(payload, str)
            chars[position:position] = list(payload)
        elif operation == "delete":
            assert isinstance(payload, int)
            del chars[position : position + payload]
        else:
            raise ValueError(operation)
    output: list[str] = []
    for ch in chars:
        if ch == "\n":
            output.append('<br class="sonarpad-preserve-line-break"/>')
        elif ch != "\r":
            output.append(escape_xml_char(ch))
    return "".join(output)


def normalize_editor_text(value: str, original_title: str) -> str:
    prefix = f"{original_title}\n\n" if original_title else ""
    if prefix and value.startswith(prefix):
        return prefix + "".join(
            line.rstrip("\n").strip() + ("\n" if line.endswith("\n") else "")
            for line in value[len(prefix):].splitlines(keepends=True)
        )
    return "".join(
        line.rstrip("\n").strip() + ("\n" if line.endswith("\n") else "")
        for line in value.splitlines(keepends=True)
    )

def edit_case(model: BookModel, case: str) -> tuple[str, str, str]:
    resource, node, node_word_start, global_start, global_end = choose_edit_location(model)
    word_len = global_end - global_start
    if case == "add_words":
        inserted = " parole aggiunte da Sonarpad"
        node_ops = [("insert", node_word_start + word_len, inserted)]
        expected = model.full_text[:global_end] + inserted + model.full_text[global_end:]
    elif case == "delete_word":
        node_ops = [("delete", node_word_start, word_len)]
        expected = model.full_text[:global_start] + model.full_text[global_end:]
    elif case == "new_line":
        inserted = "\n"
        node_ops = [("insert", node_word_start + word_len, inserted)]
        expected = model.full_text[:global_end] + inserted + model.full_text[global_end:]
    elif case == "blank_line":
        inserted = "\n\n"
        node_ops = [("insert", node_word_start + word_len, inserted)]
        expected = model.full_text[:global_end] + inserted + model.full_text[global_end:]
    elif case == "four_line_breaks":
        inserted = "\n\n\n\n"
        node_ops = [("insert", node_word_start + word_len, inserted)]
        expected = model.full_text[:global_end] + inserted + model.full_text[global_end:]
    elif case == "combined":
        prefix = "Test "
        suffix = "\nseconda riga"
        node_ops = [
            ("delete", node_word_start, word_len),
            ("insert", node_word_start, prefix + suffix),
        ]
        expected = model.full_text[:global_start] + prefix + suffix + model.full_text[global_end:]
    else:
        raise ValueError(case)
    expected = normalize_editor_text(expected, model.title)
    rendered = apply_node_edits(node, node_ops)
    original_bytes = resource.html.encode("utf-8")
    updated_bytes = (
        original_bytes[: node.source_start]
        + rendered.encode("utf-8")
        + original_bytes[node.source_end :]
    )
    return resource.archive_path, updated_bytes.decode("utf-8"), expected


def clone_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    cloned = copy.copy(info)
    cloned.CRC = 0
    cloned.compress_size = 0
    cloned.file_size = 0
    cloned.header_offset = 0
    return cloned


def write_modified_epub(source: Path, destination: Path, changed_path: str, changed_html: str) -> None:
    with zipfile.ZipFile(source) as input_zip, zipfile.ZipFile(destination, "w") as output_zip:
        output_zip.comment = input_zip.comment
        for info in input_zip.infolist():
            data = changed_html.encode("utf-8") if info.filename == changed_path else input_zip.read(info)
            cloned = clone_zipinfo(info)
            output_zip.writestr(cloned, data, compress_type=info.compress_type)


def xml_parse_status(data: bytes) -> tuple[bool, str]:
    try:
        ET.fromstring(data)
        return True, ""
    except ET.ParseError as error:
        return False, str(error)


def markup_signature(data: bytes) -> tuple[list[str], Counter[tuple[str, str]]]:
    text = data.decode("utf-8", errors="replace")
    tags = MARKUP_RE.findall(text)
    attrs = Counter((match.group(0).split("=", 1)[0].strip().lower(), html.unescape(match.group(2))) for match in ATTR_RE.finditer(text))
    return tags, attrs


def validate_case(
    source: Path,
    output: Path,
    model: BookModel,
    changed_path: str,
    expected_text: str,
    expected_added_breaks: int,
) -> dict[str, object]:
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output) as edited:
        if edited.testzip() is not None:
            raise AssertionError("ZIP CRC validation failed")
        original_infos = original.infolist()
        edited_infos = edited.infolist()
        if not edited_infos or edited_infos[0].filename != "mimetype":
            raise AssertionError("mimetype is not the first entry")
        if edited_infos[0].compress_type != zipfile.ZIP_STORED:
            raise AssertionError("mimetype is compressed")
        if edited.read("mimetype") != b"application/epub+zip":
            raise AssertionError("invalid mimetype value")
        if [item.filename for item in original_infos] != [item.filename for item in edited_infos]:
            raise AssertionError("archive entry order changed")
        if [item.compress_type for item in original_infos] != [item.compress_type for item in edited_infos]:
            raise AssertionError("compression methods changed")

        changed_entries: list[str] = []
        xml_status_changes: list[str] = []
        protected_count = 0
        for info in original_infos:
            if info.is_dir():
                continue
            before = original.read(info.filename)
            after = edited.read(info.filename)
            if before != after:
                changed_entries.append(info.filename)
            suffix = Path(info.filename).suffix.lower()
            if suffix in XML_EXTENSIONS:
                before_status = xml_parse_status(before)
                after_status = xml_parse_status(after)
                if before_status[0] != after_status[0]:
                    xml_status_changes.append(info.filename)
                if info.filename == changed_path and not after_status[0]:
                    raise AssertionError(f"modified XHTML is not well-formed: {after_status[1]}")

        if changed_entries != [changed_path]:
            raise AssertionError(f"unexpected changed entries: {changed_entries}")
        if xml_status_changes:
            raise AssertionError(f"XML parse status changed: {xml_status_changes}")

        before_markup = markup_signature(original.read(changed_path))
        after_markup = markup_signature(edited.read(changed_path))
        before_tags, before_attrs = before_markup
        after_tags, after_attrs = after_markup
        is_break = lambda tag: bool(re.fullmatch(r"<\s*br\b[^>]*>", tag, re.IGNORECASE))
        if [tag for tag in before_tags if not is_break(tag)] != [
            tag for tag in after_tags if not is_break(tag)
        ]:
            raise AssertionError("non-break markup changed")
        if sum(is_break(tag) for tag in after_tags) - sum(
            is_break(tag) for tag in before_tags
        ) != expected_added_breaks:
            raise AssertionError("unexpected number of inserted line-break tags")
        if before_attrs != after_attrs:
            raise AssertionError("IDs, links, or EPUB note semantics changed")

        protected_extensions = {".css", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".woff", ".woff2", ".ttf", ".otf"}
        for info in original_infos:
            if Path(info.filename).suffix.lower() in protected_extensions:
                protected_count += 1
                if original.read(info.filename) != edited.read(info.filename):
                    raise AssertionError(f"protected resource changed: {info.filename}")

    resource = next(item for item in model.resources if item.archive_path == changed_path)
    with zipfile.ZipFile(output) as edited:
        changed_html = edited.read(changed_path).decode("utf-8")
    changed_extraction = extract_html(changed_html)
    changed_chars, _ = filter_extraction(changed_extraction)
    original_resource_length = len(resource.filtered_chars) + 1
    actual_text = (
        model.full_text[: resource.full_start]
        + "".join(changed_chars)
        + "\n"
        + model.full_text[resource.full_start + original_resource_length :]
    )
    if actual_text.replace("\r\n", "\n").replace("\r", "\n") != expected_text.replace("\r\n", "\n").replace("\r", "\n"):
        raise AssertionError("edited text did not round-trip exactly")

    return {
        "changed_entry": changed_path,
        "archive_entries": len(original_infos),
        "protected_resources": protected_count,
        "spine_items": len(model.spine_paths),
        "text_chars": len(actual_text),
    }


def validate_no_change(source: Path, output: Path, model: BookModel) -> dict[str, object]:
    shutil.copy2(source, output)
    if source.read_bytes() != output.read_bytes():
        raise AssertionError("no-change Save As was not byte-identical")
    with zipfile.ZipFile(output) as archive:
        if archive.testzip() is not None:
            raise AssertionError("no-change copy failed ZIP CRC validation")
        infos = archive.infolist()
        if not infos or infos[0].filename != "mimetype":
            raise AssertionError("no-change copy lost the first mimetype entry")
        if infos[0].compress_type != zipfile.ZIP_STORED:
            raise AssertionError("no-change copy compressed the mimetype entry")
        protected_extensions = {
            ".css", ".jpg", ".jpeg", ".png", ".gif", ".svg",
            ".woff", ".woff2", ".ttf", ".otf",
        }
        protected_count = sum(
            Path(info.filename).suffix.lower() in protected_extensions for info in infos
        )
        link_attributes = 0
        note_markers = 0
        for info in infos:
            if info.is_dir():
                continue
            data = archive.read(info.filename)
            link_attributes += len(
                ATTR_RE.findall(data.decode("utf-8", errors="ignore"))
            )
            note_markers += len(NOTE_RE.findall(data))
    return {
        "changed_entry": None,
        "archive_entries": len(infos),
        "protected_resources": protected_count,
        "link_attributes": link_attributes,
        "note_markers": note_markers,
        "spine_items": len(model.spine_paths),
        "text_chars": len(model.full_text),
        "byte_identical": True,
    }

def run(inputs: list[Path], output_dir: Path) -> list[dict[str, object]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    cases = [
        "add_words",
        "delete_word",
        "new_line",
        "blank_line",
        "four_line_breaks",
        "combined",
    ]
    results: list[dict[str, object]] = []
    for source in inputs:
        model = load_model(source)
        no_change_destination = output_dir / f"{source.stem}__no_change.epub"
        no_change_details = validate_no_change(source, no_change_destination, model)
        results.append(
            {
                "book": source.name,
                "case": "no_change",
                "status": "PASS",
                **no_change_details,
            }
        )
        for case in cases:
            changed_path, changed_html, expected = edit_case(model, case)
            destination = output_dir / f"{source.stem}__{case}.epub"
            write_modified_epub(source, destination, changed_path, changed_html)
            details = validate_case(
                source,
                destination,
                model,
                changed_path,
                expected,
                {
                    "new_line": 1,
                    "blank_line": 2,
                    "four_line_breaks": 4,
                    "combined": 1,
                }.get(case, 0),
            )
            results.append(
                {
                    "book": source.name,
                    "case": case,
                    "status": "PASS",
                    "link_attributes": no_change_details["link_attributes"],
                    "note_markers": no_change_details["note_markers"],
                    **details,
                }
            )
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("epubs", nargs="+", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--json-report", type=Path)
    args = parser.parse_args()
    for path in args.epubs:
        if not path.is_file():
            parser.error(f"file not found: {path}")
    with tempfile.TemporaryDirectory(prefix="sonarpad-epub-test-") as temporary:
        output_dir = args.output_dir or Path(temporary)
        results = run(args.epubs, output_dir)
        payload = {"total": len(results), "passed": len(results), "failed": 0, "results": results}
        if args.json_report:
            args.json_report.parent.mkdir(parents=True, exist_ok=True)
            args.json_report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
