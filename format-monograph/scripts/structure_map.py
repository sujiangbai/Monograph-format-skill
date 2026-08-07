"""Candidate generation, validation, application, and audit normalization."""

from __future__ import annotations

import copy
import hashlib
import json
import re
import zipfile
from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from _common import (
    CONTENT_PART,
    FIELD_MARKER_PATTERN,
    NS,
    FormatMonographError,
    _heading_prefix_pattern,
    _paragraph_text_without_field_results,
    content_fingerprint,
)


HEADING_PATTERNS = (
    (4, re.compile(r"^\s*\d+\.\d+\.\d+\.\d+\.?\s*\S")),
    (3, re.compile(r"^\s*\d+\.\d+\.\d+\.?\s*\S")),
    (2, re.compile(r"^\s*\d+\.\d+\.?\s*\S")),
    (1, re.compile(r"^\s*第\s*[0-9一二三四五六七八九十百]+\s*章\s*\S")),
)
CAPTION_PATTERN = re.compile(
    r"^\s*(图|表)\s*(\d+(?:\.\d+){1,3})\s*[-－—–]\s*(\d+)\s*(\S.*)$"
)


def text_sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _paragraph_has_payload(element: Any) -> bool:
    text = "".join(element.xpath(".//w:t/text()")).strip()
    if text:
        return True
    protected = (
        ".//w:tbl | .//w:drawing | .//w:object | .//w:pict | "
        ".//w:bookmarkStart | .//w:commentRangeStart | .//w:footnoteReference | "
        ".//w:endnoteReference | .//w:fldChar"
    )
    return bool(element.xpath(protected))


def _paragraph_style_signature(paragraph: Any) -> str:
    p_pr = paragraph._p.pPr
    r_pr = None if p_pr is None else p_pr.find(qn("w:rPr"))
    source = r_pr if r_pr is not None else (p_pr if p_pr is not None else paragraph._p)
    payload = etree.tostring(source)
    return hashlib.sha256(payload).hexdigest()


def _trailing_empty_sections(document: Any) -> list[dict[str, Any]]:
    body = document.element.body
    children = list(body)
    paragraph_index = {id(p._p): index for index, p in enumerate(document.paragraphs)}
    boundaries: list[tuple[int, int]] = []
    for child_index, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        p_pr = child.find(qn("w:pPr"))
        if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
            boundaries.append((child_index, paragraph_index[id(child)]))

    candidates = []
    starts = [0] + [child_index + 1 for child_index, _ in boundaries]
    ends = [child_index for child_index, _ in boundaries] + [len(children) - 2]
    for section_index in range(len(starts) - 1, 0, -1):
        start, end = starts[section_index], ends[section_index]
        span = children[start : end + 1] if end >= start else []
        if any(_paragraph_has_payload(element) for element in span):
            break
        previous_boundary = boundaries[section_index - 1][1]
        paragraph = document.paragraphs[previous_boundary]
        candidates.append(
            {
                "section": section_index,
                "previous_boundary_paragraph": previous_boundary,
                "previous_boundary_sha256": text_sha256(paragraph.text),
                "approved_delete": False,
                "confidence": "high",
            }
        )
    return candidates


def candidate_structure_map(path: Path) -> dict[str, Any]:
    from _common import load_document

    document = load_document(path)
    headings = []
    captions = []
    for index, paragraph in enumerate(document.paragraphs):
        value = paragraph.text
        for level, pattern in HEADING_PATTERNS:
            if pattern.match(value):
                headings.append(
                    {
                        "paragraph": index,
                        "text_sha256": text_sha256(value),
                        "level": level,
                        "source_style": paragraph.style.name if paragraph.style else None,
                        "direct_format_sha256": _paragraph_style_signature(paragraph),
                        "approved": False,
                    }
                )
                break
        match = CAPTION_PATTERN.match(value)
        if match:
            label, hierarchy, sequence, _ = match.groups()
            captions.append(
                {
                    "paragraph": index,
                    "text_sha256": text_sha256(value),
                    "label": label,
                    "sequence_name": "Figure" if label == "图" else "Table",
                    "heading_level": len(hierarchy.split(".")),
                    "cached_hierarchy": hierarchy,
                    "cached_sequence": sequence,
                    "approved": False,
                }
            )

    tables = []
    for index, table in enumerate(document.tables):
        first_row = "\u241f".join(cell.text for cell in table.rows[0].cells) if table.rows else ""
        tables.append(
            {
                "table": index,
                "first_row_sha256": text_sha256(first_row),
                "repeat_header": False,
                "prevent_normal_row_split": False,
                "approved": False,
            }
        )

    return {
        "schema_version": "1.0",
        "status": "candidate",
        "source_content_fingerprint_sha256": content_fingerprint(path),
        "toc_ranges": [],
        "headings": headings,
        "captions": captions,
        "tables": tables,
        "trailing_empty_sections": _trailing_empty_sections(document),
        "conflicts": [],
    }


def load_structure_map(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FormatMonographError(f"Invalid structure map: {path}: {exc}") from exc
    if value.get("schema_version") != "1.0":
        raise FormatMonographError("Structure map schema_version must be 1.0.")
    if value.get("status") != "approved":
        raise FormatMonographError("Structure map status must be approved.")
    if value.get("conflicts"):
        raise FormatMonographError("Structure map contains unresolved conflicts.")
    return value


def validate_structure_map_source(path: Path, structure_map: dict[str, Any]) -> None:
    actual = content_fingerprint(path)
    expected = structure_map.get("source_content_fingerprint_sha256")
    if actual != expected:
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
