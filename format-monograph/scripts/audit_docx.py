#!/usr/bin/env python3
"""Audit content, protected objects, and automatically verifiable profile rules."""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

from docx.oxml.ns import qn

from _common import (
    FormatMonographError,
    _field_instruction_for_marker,
    _heading_prefix_pattern,
    content_fingerprint,
    equation_inventory,
    load_document,
    protected_object_manifest,
    style_name_for_selector,
    write_json,
)
from validate_profile import validate
from structure_map import load_structure_map, structure_content_fingerprint


def close(actual: Any, expected: Any, tolerance: float = 0.05) -> bool:
    if actual is None:
        return expected is None
    return math.isclose(float(actual), float(expected), abs_tol=tolerance)


def style_font_value(style: Any, attribute: str) -> Any:
    r_pr = style.element.rPr
    r_fonts = None if r_pr is None else r_pr.rFonts
    if r_fonts is None:
        return None
    return r_fonts.get(qn(f"w:{attribute}"))


def style_value(style: Any, key: str) -> Any:
    font, pf = style.font, style.paragraph_format
    if key == "font_name":
        return font.name
    if key == "font_name_ascii":
        return style_font_value(style, "ascii") or style_font_value(style, "hAnsi")
    if key == "font_name_east_asia":
        return style_font_value(style, "eastAsia")
    if key == "font_name_complex_script":
        return style_font_value(style, "cs")
    if key == "font_size_pt":
        return None if font.size is None else font.size.pt
    if key in {"bold", "italic"}:
        return getattr(font, key)
    if key == "color_hex":
        return None if font.color.rgb is None else str(font.color.rgb)
    if key == "a…11018 tokens truncated…
        raise FormatMonographError(
            "Structure map source fingerprint does not match the input DOCX."
        )


def _verified_paragraph(document: Any, entry: dict[str, Any]) -> Any:
    index = int(entry["paragraph"])
    if not 0 <= index < len(document.paragraphs):
        raise FormatMonographError(f"Structure-map paragraph is out of range: {index}")
    paragraph = document.paragraphs[index]
    if text_sha256(paragraph.text) != entry["text_sha256"]:
        raise FormatMonographError(
            f"Structure-map paragraph hash mismatch at paragraph {index}."
        )
    return paragraph


def _clear_paragraph(paragraph: Any) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def _text_run(value: str) -> Any:
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    if value[:1].isspace() or value[-1:].isspace():
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value
    run.append(text)
    return run


def _simple_field(instruction: str, placeholder: str) -> Any:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    field.set(qn("w:dirty"), "true")
    field.append(_text_run(placeholder))
    return field


def _apply_toc_ranges(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("toc_ranges", []):
        if not entry.get("approved"):
            continue
        start, end = int(entry["start_paragraph"]), int(entry["end_paragraph"])
        hashes = entry.get("paragraph_sha256", [])
        if start < 0 or end < start or end >= len(document.paragraphs):
            raise FormatMonographError("Approved TOC range is out of bounds.")
        if len(hashes) != end - start + 1:
            raise FormatMonographError("Approved TOC range hash count is invalid.")
        for offset, index in enumerate(range(start, end + 1)):
            if text_sha256(document.paragraphs[index].text) != hashes[offset]:
                raise FormatMonographError(
                    f"Approved TOC range hash mismatch at paragraph {index}."
                )
            _clear_paragraph(document.paragraphs[index])
        levels = int(entry.get("levels", 4))
        document.paragraphs[start]._p.append(
            _simple_field(
                f'TOC \\o "1-{levels}" \\h \\z \\u',
                "Update table of contents",
            )
        )
        changed += end - start + 1
    return changed


def _apply_headings(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("headings", []):
        if not entry.get("approved"):
            continue
        level = int(entry["level"])
        if level not in {1, 2, 3, 4}:
            raise FormatMonographError(f"Invalid approved heading level: {level}")
        _verified_paragraph(document, entry).style = f"Heading {level}"
        changed += 1
    return changed


def _remove_prefix_from_runs(paragraph: Any, length: int) -> None:
    remaining = length
    for run in paragraph.runs:
        if remaining <= 0:
            break
        take = min(len(run.text), remaining)
        run.text = run.text[take:]
        remaining -= take
    if remaining:
        raise FormatMonographError("Caption prefix spans an unsupported object.")


def _apply_captions(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("captions", []):
        if not entry.get("approved"):
            continue
        paragraph = _verified_paragraph(document, entry)
        match = CAPTION_PATTERN.match(paragraph.text)
        if not match:
            raise FormatMonographError(
                f"Approved caption no longer matches at paragraph {entry['paragraph']}."
            )
        label, hierarchy, sequence, _ = match.groups()
        if label != entry["label"]:
            raise FormatMonographError("Approved caption label does not match.")
        description_start = match.start(4)
        _remove_prefix_from_runs(paragraph, description_start)
        insertion = 1 if paragraph._p.pPr is not None else 0
        elements = [
            _text_run(f"{label} "),
            _simple_field(
                f"STYLEREF {int(entry['heading_level'])} \\s",
                entry.get("cached_hierarchy", hierarchy),
            ),
            _text_run("-"),
            _simple_field(
                f"SEQ {entry['sequence_name']} \\* ARABIC \\s {int(entry['heading_level'])}",
                entry.get("cached_sequence", sequence),
            ),
            _text_run(" "),
        ]
        for element in reversed(elements):
            paragraph._p.insert(insertion, element)
        paragraph.style = "Caption"
        changed += 1
    return changed


def _set_row_property(row: Any, name: str, enabled: bool) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    element = tr_pr.find(qn(f"w:{name}"))
    if enabled and element is None:
        element = OxmlElement(f"w:{name}")
        element.set(qn("w:val"), "true")
        tr_pr.append(element)
    elif not enabled and element is not None:
        tr_pr.remove(element)


def _apply_tables(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("tables", []):
        if not entry.get("approved"):
            continue
        index = int(entry["table"])
        if not 0 <= index < len(document.tables):
            raise FormatMonographError(f"Structure-map table is out of range: {index}")
        table = document.tables[index]
        first_row = "\u241f".join(cell.text for cell in table.rows[0].cells) if table.rows else ""
        if text_sha256(first_row) != entry["first_row_sha256"]:
            raise FormatMonographError(f"Structure-map table hash mismatch: {index}")
        if entry.get("repeat_header") and table.rows:
            _set_row_property(table.rows[0], "tblHeader", True)
        if entry.get("prevent_normal_row_split"):
            for row in table.rows:
                _set_row_property(row, "cantSplit", True)
        changed += 1
    return changed


def _final_section_is_empty(document: Any, boundary_index: int) -> bool:
    body = document.element.body
    boundary = document.paragraphs[boundary_index]._p
    children = list(body)
    start = children.index(boundary) + 1
    return not any(_paragraph_has_payload(element) for element in children[start:-1])


def _remove_final_empty_section(document: Any, entry: dict[str, Any]) -> None:
    boundary_index = int(entry["previous_boundary_paragraph"])
    if not 0 <= boundary_index < len(document.paragraphs):
        raise FormatMonographError("Trailing-section boundary is out of range.")
    boundary = document.paragraphs[boundary_index]
    if text_sha256(boundary.text) != entry["previous_boundary_sha256"]:
        raise FormatMonographError("Trailing-section boundary hash mismatch.")
    if not _final_section_is_empty(document, boundary_index):
        raise FormatMonographError("Approved trailing section is no longer empty.")

    p_pr = boundary._p.pPr
    previous = None if p_pr is None else p_pr.find(qn("w:sectPr"))
    final = document.element.body.sectPr
    if previous is None or final is None:
        raise FormatMonographError("Trailing-section boundary is missing section properties.")
    blocked = ("headerReference", "footerReference", "pgNumType", "titlePg")
    if any(final.find(qn(f"w:{name}")) is not None for name in blocked):
        raise FormatMonographError(
            "Trailing section has independent header, footer, or page-number settings."
        )

    body = document.element.body
    children = list(body)
    boundary_position = children.index(boundary._p)
    for child in list(children[boundary_position + 1 : -1]):
        body.remove(child)
    body.replace(final, copy.deepcopy(previous))
    p_pr.remove(previous)
    if not _paragraph_has_payload(boundary._p) and len(p_pr) == 0:
        body.remove(boundary._p)


def _apply_trailing_sections(document: Any, structure_map: dict[str, Any]) -> int:
    approved = [
        entry
        for entry in structure_map.get("trailing_empty_sections", [])
        if entry.get("approved_delete")
    ]
    approved.sort(key=lambda item: int(item["section"]), reverse=True)
    for entry in approved:
        if int(entry["section"]) != len(document.sections) - 1:
            raise FormatMonographError(
                "Trailing empty sections must be approved and removed from the end inward."
            )
        _remove_final_empty_section(document, entry)
    return len(approved)


def apply_structure_map(document: Any, structure_map: dict[str, Any]) -> list[dict[str, Any]]:
    changes = [
        {"kind": "structure_toc", "targets": _apply_toc_ranges(document, structure_map)},
        {"kind": "structure_headings", "targets": _apply_headings(document, structure_map)},
        {"kind": "structure_captions", "targets": _apply_captions(document, structure_map)},
        {"kind": "structure_tables", "targets": _apply_tables(document, structure_map)},
        {
            "kind": "structure_trailing_sections",
            "targets": _apply_trailing_sections(document, structure_map),
        },
    ]
    return [change for change in changes if change["targets"]]


def _approved_indexes(structure_map: dict[str, Any], key: str) -> dict[int, dict[str, Any]]:
    return {
        int(entry["paragraph"]): entry
        for entry in structure_map.get(key, [])
        if entry.get("approved")
    }


def structure_content_inventory(
    path: Path, structure_map: dict[str, Any]
) -> dict[str, list[str]]:
    toc_indexes = {
        index
        for entry in structure_map.get("toc_ranges", [])
        if entry.get("approved")
        for index in range(int(entry["start_paragraph"]), int(entry["end_paragraph"]) + 1)
    }
    headings = _approved_indexes(structure_map, "headings")
    captions = _approved_indexes(structure_map, "captions")
    tail_cutoffs = []
    for entry in structure_map.get("trailing_empty_sections", []):
        if not entry.get("approved_delete"):
            continue
        boundary = int(entry["previous_boundary_paragraph"])
        boundary_is_empty = entry.get("previous_boundary_sha256") == text_sha256("")
        tail_cutoffs.append(boundary if boundary_is_empty else boundary + 1)
    tail_cutoff = min(tail_cutoffs, default=None)

    result: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if not CONTENT_PART.match(name):
                continue
            root = etree.fromstring(package.read(name))
            values = []
            body_index = -1
            for paragraph in root.xpath(".//w:p", namespaces=NS):
                direct_body_paragraph = (
                    name == "word/document.xml"
                    and paragraph.getparent() is not None
                    and paragraph.getparent().tag == qn("w:body")
                )
                if direct_body_paragraph:
                    body_index += 1
                current_index = body_index if direct_body_paragraph else None
                if current_index is not None and tail_cutoff is not None and current_index >= tail_cutoff:
                    continue
                value = _paragraph_text_without_field_results(paragraph)
                value = FIELD_MARKER_PATTERN.sub("", value)
                if current_index in toc_indexes:
                    value = ""
                elif current_index in headings:
                    match = _heading_prefix_pattern(int(headings[current_index]["level"])).match(value)
                    if match:
                        value = value[match.end() :]
                elif current_index in captions:
                    match = CAPTION_PATTERN.match(value)
                    if match:
                        value = f"{match.group(1)} {match.group(4)}"
                    else:
                        value = re.sub(r"^(图|表)\s+[-－—–]?\s*", r"\1 ", value)
                values.append(value)
            result[name] = values
    return result


def structure_content_fingerprint(path: Path, structure_map: dict[str, Any]) -> str:
    result = structure_content_inventory(path, structure_map)
    encoded = json.dumps(result, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()
