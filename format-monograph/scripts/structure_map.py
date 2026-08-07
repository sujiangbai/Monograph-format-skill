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
from docx.text.paragraph import Paragraph
from lxml import etree

from _common import (
    CONTENT_PART,
    FIELD_MARKER_PATTERN,
    NS,
    FormatMonographError,
    _heading_prefix_pattern,
    _paragraph_text_without_field_results,
    content_fingerprint,
    ensure_paragraph_style,
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
LOOSE_CAPTION_PATTERN = re.compile(r"^\s*(图|表)\s*(\S.*)?$")

ROLE_ALIASES = {
    "body_text": "body",
    "chapter_title": "heading_1",
    "level_2_section": "heading_2",
    "level_3_section": "heading_3",
    "level_4_section": "heading_4",
    "heading1": "heading_1",
    "heading2": "heading_2",
    "heading3": "heading_3",
    "heading4": "heading_4",
    "caption": "all_captions",
    "all": "all_captions",
}

ROLE_STYLE_NAMES = {
    "body": "Normal",
    "title": "Title",
    "heading_1": "Heading 1",
    "heading_2": "Heading 2",
    "heading_3": "Heading 3",
    "heading_4": "Heading 4",
    "figure_caption": "Caption",
    "table_caption": "Caption",
    "equation_caption": "Caption",
    "long_quote": "Quote",
    "reference_entry": "Bibliography",
    "answer": "Normal",
    "teaching_callout": "Normal",
}


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


def _body_locator(index: int) -> dict[str, Any]:
    return {"kind": "body_paragraph", "paragraph": index}


def _cell_locator(table: int, row: int, cell: int, paragraph: int) -> dict[str, Any]:
    return {
        "kind": "table_cell_paragraph",
        "table": table,
        "row": row,
        "cell": cell,
        "paragraph": paragraph,
    }


def _locator_key(locator: dict[str, Any]) -> tuple[Any, ...]:
    if locator.get("kind") == "body_paragraph":
        return ("body", int(locator["paragraph"]))
    if locator.get("kind") == "table_cell_paragraph":
        return (
            "cell",
            int(locator["table"]),
            int(locator["row"]),
            int(locator["cell"]),
            int(locator["paragraph"]),
        )
    raise FormatMonographError(f"Unsupported paragraph locator: {locator}")


def resolve_paragraph_locator(document: Any, locator: dict[str, Any]) -> Any:
    cached = getattr(document, "_format_monograph_locator_cache", {}).get(
        _locator_key(locator)
    )
    if cached is not None:
        return cached
    kind = locator.get("kind")
    if kind == "body_paragraph":
        index = int(locator["paragraph"])
        if not 0 <= index < len(document.paragraphs):
            raise FormatMonographError(f"Body paragraph locator is out of range: {index}")
        return document.paragraphs[index]
    if kind == "table_cell_paragraph":
        table_index = int(locator["table"])
        row_index = int(locator["row"])
        cell_index = int(locator["cell"])
        paragraph_index = int(locator["paragraph"])
        if not 0 <= table_index < len(document.tables):
            raise FormatMonographError(f"Table locator is out of range: {table_index}")
        table = document.tables[table_index]
        if not 0 <= row_index < len(table.rows):
            raise FormatMonographError(f"Table row locator is out of range: {row_index}")
        row = table.rows[row_index]
        if not 0 <= cell_index < len(row.cells):
            raise FormatMonographError(f"Table cell locator is out of range: {cell_index}")
        paragraphs = row.cells[cell_index].paragraphs
        if not 0 <= paragraph_index < len(paragraphs):
            raise FormatMonographError(
                f"Table-cell paragraph locator is out of range: {paragraph_index}"
            )
        return paragraphs[paragraph_index]
    raise FormatMonographError(f"Unsupported paragraph locator kind: {kind}")


def _verified_locator_paragraph(document: Any, entry: dict[str, Any]) -> Any:
    key = _locator_key(entry["locator"])
    verified = getattr(document, "_format_monograph_verified_locators", set())
    if key in verified:
        return resolve_paragraph_locator(document, entry["locator"])
    paragraph = resolve_paragraph_locator(document, entry["locator"])
    expected = entry.get("text_sha256")
    if expected and text_sha256(paragraph.text) != expected:
        raise FormatMonographError(
            f"Structure-map paragraph hash mismatch at {key}."
        )
    return paragraph


def prime_structure_map_locators(document: Any, structure_map: dict[str, Any]) -> None:
    if structure_map.get("schema_version") != "1.1":
        return
    cache = getattr(document, "_format_monograph_locator_cache", {})
    verified = getattr(document, "_format_monograph_verified_locators", set())
    for entry in structure_map.get("paragraph_roles", []):
        if not entry.get("approved"):
            continue
        paragraph = resolve_paragraph_locator(document, entry["locator"])
        if text_sha256(paragraph.text) != entry["text_sha256"]:
            raise FormatMonographError(
                f"Structure-map paragraph hash mismatch at {_locator_key(entry['locator'])}."
            )
        key = _locator_key(entry["locator"])
        cache[key] = paragraph
        verified.add(key)
    setattr(document, "_format_monograph_locator_cache", cache)
    setattr(document, "_format_monograph_verified_locators", verified)


def normalized_role(value: str) -> str:
    return ROLE_ALIASES.get(value, value)


def approved_role_paragraphs(
    document: Any, structure_map: dict[str, Any], selector: dict[str, str]
) -> list[Any]:
    if structure_map.get("schema_version") != "1.1":
        return []
    wanted = normalized_role(selector["value"])
    caption_roles = {"figure_caption", "table_caption", "equation_caption"}
    result = []
    seen: set[int] = set()
    for entry in structure_map.get("paragraph_roles", []):
        if not entry.get("approved"):
            continue
        role = normalized_role(str(entry.get("role", "unknown")))
        matches = role == wanted or (wanted == "all_captions" and role in caption_roles)
        if not matches:
            continue
        try:
            paragraph = _verified_locator_paragraph(document, entry)
        except FormatMonographError:
            locator = entry.get("locator", {})
            expected_style = ROLE_STYLE_NAMES.get(role)
            if role.startswith("heading_") and locator.get("kind") == "body_paragraph":
                paragraph = resolve_paragraph_locator(document, locator)
                if not paragraph.style or paragraph.style.name != expected_style:
                    raise
            elif role in caption_roles and locator.get("kind") == "table_cell_paragraph":
                table_index = int(locator["table"])
                if not 0 <= table_index < len(document.tables):
                    raise
                previous = document.tables[table_index]._tbl.getprevious()
                if previous is None or previous.tag != qn("w:p"):
                    raise
                paragraph = Paragraph(previous, document.tables[table_index]._parent)
                if not paragraph.style or paragraph.style.name != "Caption":
                    raise
            else:
                raise
        if id(paragraph._p) not in seen:
            result.append(paragraph)
            seen.add(id(paragraph._p))
    return result


def approved_data_tables(
    document: Any, structure_map: dict[str, Any]
) -> list[tuple[Any, dict[str, Any]]]:
    if structure_map.get("schema_version") != "1.1":
        return []
    result = []
    for entry in structure_map.get("tables", []):
        if not entry.get("approved") or entry.get("kind") != "data":
            continue
        index = int(entry["table"])
        if not 0 <= index < len(document.tables):
            raise FormatMonographError(f"Structure-map table is out of range: {index}")
        normalized = copy.deepcopy(entry)
        deleted = getattr(document, "_format_monograph_deleted_table_rows", {}).get(index)
        if deleted is None:
            deleted = next(
                (
                    int(caption["locator"]["row"])
                    for caption in structure_map.get("captions", [])
                    if caption.get("approved")
                    and caption.get("migrate_outside_table")
                    and caption.get("locator", {}).get("kind")
                    == "table_cell_paragraph"
                    and int(caption["locator"]["table"]) == index
                ),
                None,
            )
        if deleted is not None:
            normalized["repeat_header_rows"] = [
                int(row) - 1 if int(row) > deleted else int(row)
                for row in normalized.get("repeat_header_rows", [])
                if int(row) != deleted
            ]
            normalized["header_rows"] = [
                int(row) - 1 if int(row) > deleted else int(row)
                for row in normalized.get("header_rows", [])
                if int(row) != deleted
            ]
            normalized["caption_row"] = None
        result.append((document.tables[index], normalized))
    return result


def _section_evidence(document: Any, sect_pr: Any) -> dict[str, Any]:
    header_ids = sect_pr.xpath("./w:headerReference/@r:id")
    footer_ids = sect_pr.xpath("./w:footerReference/@r:id")

    def related_has_payload(rel_id: str) -> bool:
        part = document.part.related_parts.get(rel_id)
        if part is None:
            return True
        try:
            root = etree.fromstring(part.blob)
        except (ValueError, etree.XMLSyntaxError):
            return True
        return bool(
            "".join(root.xpath(".//w:t/text()", namespaces=NS)).strip()
            or root.xpath(".//w:drawing | .//w:object | .//w:pict", namespaces=NS)
        )

    independent = {
        "header_reference_count": len(header_ids),
        "footer_reference_count": len(footer_ids),
        "header_footer_has_payload": any(
            related_has_payload(rel_id) for rel_id in header_ids + footer_ids
        ),
        "page_number_start": (
            sect_pr.xpath("string(./w:pgNumType/@w:start)") or None
        ),
        "different_first_page": bool(sect_pr.find(qn("w:titlePg")) is not None),
        "section_type": (
            sect_pr.xpath("string(./w:type/@w:val)") or None
        ),
        "has_page_geometry": bool(
            sect_pr.find(qn("w:pgSz")) is not None
            or sect_pr.find(qn("w:pgMar")) is not None
        ),
    }
    independent["empty_header_footer_references"] = bool(
        independent["header_reference_count"] or independent["footer_reference_count"]
    ) and not independent["header_footer_has_payload"]
    independent["safe_to_delete"] = not any(
        (
            independent["header_footer_has_payload"],
            independent["page_number_start"],
            independent["different_first_page"],
        )
    )
    return independent


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
        final_sect_pr = document.element.body.sectPr
        evidence = _section_evidence(document, final_sect_pr)
        candidates.append(
            {
                "section": section_index,
                "previous_boundary_paragraph": previous_boundary,
                "previous_boundary_sha256": text_sha256(paragraph.text),
                "approved_delete": False,
                "confidence": "high" if evidence["safe_to_delete"] else "low",
                "evidence": evidence,
            }
        )
    return candidates


def _chapter_number(value: str) -> int | None:
    match = re.match(r"^\s*第\s*([0-9一二三四五六七八九十百]+)\s*章", value)
    if not match:
        return None
    token = match.group(1)
    if token.isdigit():
        return int(token)
    digits = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    if token == "十":
        return 10
    if "十" in token:
        left, right = token.split("十", 1)
        return (digits.get(left, 1) * 10) + digits.get(right, 0)
    return digits.get(token)


def _heading_number(value: str, level: int) -> tuple[int, ...] | None:
    if level == 1:
        chapter = _chapter_number(value)
        return None if chapter is None else (chapter,)
    match = re.match(r"^\s*(\d+(?:[.-]\d+){%d})" % (level - 1), value)
    if not match:
        return None
    return tuple(int(part) for part in re.split(r"[.-]", match.group(1)))


def _numbering_anomalies(headings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    anomalies = []
    seen: set[tuple[int, ...]] = set()
    last_by_parent: dict[tuple[int, ...], int] = {}
    for entry in headings:
        number = tuple(entry.get("cached_number", []))
        if not number:
            continue
        locator = entry.get("locator", _body_locator(int(entry["paragraph"])))
        if number in seen:
            anomalies.append(
                {
                    "kind": "duplicate_number",
                    "locator": locator,
                    "observed": list(number),
                    "status": "open",
                }
            )
        parent = number[:-1]
        if len(number) > 1 and parent not in seen:
            anomalies.append(
                {
                    "kind": "missing_parent",
                    "locator": locator,
                    "observed": list(number),
                    "status": "open",
                }
            )
        previous = last_by_parent.get(parent)
        expected_last = previous + 1 if previous is not None else 1
        if len(number) > 1 and number[-1] != expected_last:
            anomalies.append(
                {
                    "kind": "sequence_gap",
                    "locator": locator,
                    "observed": list(number),
                    "expected_last": expected_last,
                    "status": "open",
                }
            )
        seen.add(number)
        last_by_parent[parent] = number[-1]
    return anomalies


def _role_for_paragraph(paragraph: Any, detected_level: int | None) -> str:
    if detected_level:
        return f"heading_{detected_level}"
    style_name = paragraph.style.name if paragraph.style else ""
    if style_name == "Title":
        return "title"
    match = re.fullmatch(r"Heading (\d+)", style_name)
    if match and 1 <= int(match.group(1)) <= 4:
        return f"heading_{match.group(1)}"
    if style_name == "Quote":
        return "long_quote"
    if style_name == "Bibliography":
        return "reference_entry"
    return "unknown"


def _caption_entry(value: str, locator: dict[str, Any]) -> dict[str, Any] | None:
    match = CAPTION_PATTERN.match(value)
    if not match:
        loose = LOOSE_CAPTION_PATTERN.match(value)
        if not loose:
            return None
        label = loose.group(1)
        return {
            "locator": locator,
            "text_sha256": text_sha256(value),
            "label": label,
            "sequence_name": "Figure" if label == "图" else "Table",
            "completeness": "incomplete",
            "hierarchy_status": "unresolved",
            "approved": False,
        }
    label, hierarchy, sequence, _ = match.groups()
    return {
        "locator": locator,
        "text_sha256": text_sha256(value),
        "label": label,
        "sequence_name": "Figure" if label == "图" else "Table",
        "heading_level": len(hierarchy.split(".")),
        "cached_hierarchy": hierarchy,
        "cached_sequence": sequence,
        "completeness": "complete",
        "hierarchy_status": "pending_qa",
        "approved": False,
    }


def _table_text_hash(table: Any) -> str:
    value = "\u241e".join(
        "\u241f".join(cell.text for cell in row.cells) for row in table.rows
    )
    return text_sha256(value)


def candidate_structure_map(path: Path) -> dict[str, Any]:
    from _common import load_document

    document = load_document(path)
    headings = []
    captions = []
    paragraph_roles = []
    chapter_starts = []
    current_chapter = None
    for index, paragraph in enumerate(document.paragraphs):
        value = paragraph.text
        detected_level = None
        for level, pattern in HEADING_PATTERNS:
            if pattern.match(value):
                detected_level = level
                entry = {
                    "paragraph": index,
                    "locator": _body_locator(index),
                    "text_sha256": text_sha256(value),
                    "level": level,
                    "cached_number": list(_heading_number(value, level) or ()),
                    "source_style": paragraph.style.name if paragraph.style else None,
                    "direct_format_sha256": _paragraph_style_signature(paragraph),
                    "approved": False,
                }
                headings.append(entry)
                if level == 1:
                    chapter = _chapter_number(value)
                    if chapter is not None:
                        chapter_starts.append(chapter)
                        current_chapter = chapter
                break
        caption = _caption_entry(value, _body_locator(index))
        if caption:
            if caption["completeness"] == "complete" and current_chapter is not None:
                observed = int(caption["cached_hierarchy"].split(".")[0])
                caption["hierarchy_status"] = (
                    "match" if observed == current_chapter else "mismatch"
                )
            caption["paragraph"] = index
            captions.append(caption)
            role = "figure_caption" if caption["label"] == "图" else "table_caption"
        else:
            role = _role_for_paragraph(paragraph, detected_level)
        if value.strip():
            paragraph_roles.append(
                {
                    "locator": _body_locator(index),
                    "text_sha256": text_sha256(value),
                    "role": role,
                    "source_style": paragraph.style.name if paragraph.style else None,
                    "direct_format_sha256": _paragraph_style_signature(paragraph),
                    "approved": False,
                }
            )

    tables = []
    for table_index, table in enumerate(document.tables):
        caption_row = None
        header_rows: list[int] = []
        seen_cells: set[int] = set()
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                if id(cell._tc) in seen_cells:
                    continue
                seen_cells.add(id(cell._tc))
                for paragraph_index, paragraph in enumerate(cell.paragraphs):
                    locator = _cell_locator(
                        table_index, row_index, cell_index, paragraph_index
                    )
                    caption = _caption_entry(paragraph.text, locator)
                    if not caption:
                        continue
                    if (
                        caption["completeness"] == "complete"
                        and len(set(chapter_starts)) == 1
                    ):
                        observed = int(caption["cached_hierarchy"].split(".")[0])
                        caption["hierarchy_status"] = (
                            "match" if observed == chapter_starts[0] else "mismatch"
                        )
                    caption["migrate_outside_table"] = False
                    captions.append(caption)
                    paragraph_roles.append(
                        {
                            "locator": locator,
                            "text_sha256": text_sha256(paragraph.text),
                            "role": (
                                "figure_caption"
                                if caption["label"] == "图"
                                else "table_caption"
                            ),
                            "source_style": paragraph.style.name if paragraph.style else None,
                            "direct_format_sha256": _paragraph_style_signature(paragraph),
                            "approved": False,
                        }
                    )
                    if row_index == 0 and len({id(item._tc) for item in row.cells}) == 1:
                        caption_row = 0
                        if len(table.rows) > 1:
                            header_rows = [1]

        first_row = "\u241f".join(cell.text for cell in table.rows[0].cells) if table.rows else ""
        tables.append(
            {
                "table": table_index,
                "table_text_sha256": _table_text_hash(table),
                "first_row_sha256": text_sha256(first_row),
                "kind": "unknown",
                "caption_row": caption_row,
                "header_rows": header_rows,
                "repeat_header_rows": [],
                "prevent_normal_row_split": False,
                "approved": False,
            }
        )

    return {
        "schema_version": "1.1",
        "status": "candidate",
        "source_content_fingerprint_sha256": content_fingerprint(path),
        "paragraph_roles": paragraph_roles,
        "numbering": {
            "mode": "single_chapter" if len(set(chapter_starts)) == 1 else "whole_book",
            "chapter_start": chapter_starts[0] if chapter_starts else None,
            "heading_levels": 4,
            "expected_progression": "strict",
            "approved": False,
            "anomalies": _numbering_anomalies(headings),
        },
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
    if value.get("schema_version") not in {"1.0", "1.1"}:
        raise FormatMonographError("Structure map schema_version must be 1.0 or 1.1.")
    if value.get("status") != "approved":
        raise FormatMonographError("Structure map status must be approved.")
    if value.get("conflicts"):
        raise FormatMonographError("Structure map contains unresolved conflicts.")
    if value.get("schema_version") == "1.1":
        for entry in value.get("paragraph_roles", []):
            if entry.get("approved") and entry.get("role") == "unknown":
                raise FormatMonographError("Approved paragraph role cannot be unknown.")
            if "locator" not in entry or "text_sha256" not in entry:
                raise FormatMonographError(
                    "Structure map 1.1 paragraph roles require locator and text_sha256."
                )
        for entry in value.get("captions", []):
            if not entry.get("approved"):
                continue
            if entry.get("completeness") != "complete":
                raise FormatMonographError("Incomplete captions cannot be approved.")
            if entry.get("hierarchy_status") not in {"match", "accepted"}:
                raise FormatMonographError(
                    "Caption hierarchy must match or be explicitly accepted."
                )
        numbering = value.get("numbering", {})
        if numbering.get("approved"):
            chapter_start = numbering.get("chapter_start")
            if not isinstance(chapter_start, int) or chapter_start < 1:
                raise FormatMonographError(
                    "Approved numbering requires a positive integer chapter_start."
                )
            unresolved = [
                item
                for item in numbering.get("anomalies", [])
                if item.get("status", "open") != "accepted"
            ]
            if unresolved:
                raise FormatMonographError(
                    "Approved numbering contains unresolved progression anomalies."
                )
    return value


def validate_structure_map_source(path: Path, structure_map: dict[str, Any]) -> None:
    actual = content_fingerprint(path)
    expected = structure_map.get("source_content_fingerprint_sha256")
    if actual != expected:
        raise FormatMonographError(
            "Structure map source fingerprint does not match the input DOCX."
        )
    from _common import load_document

    document = load_document(path)
    if structure_map.get("schema_version") == "1.1":
        prime_structure_map_locators(document, structure_map)
    for key in ("headings", "captions"):
        for entry in structure_map.get(key, []):
            if not entry.get("approved"):
                continue
            if entry.get("locator"):
                _verified_locator_paragraph(document, entry)
            else:
                _verified_paragraph(document, entry)
    for entry in structure_map.get("tables", []):
        if not entry.get("approved"):
            continue
        index = int(entry["table"])
        if not 0 <= index < len(document.tables):
            raise FormatMonographError(f"Structure-map table is out of range: {index}")
        if entry.get("table_text_sha256") and _table_text_hash(document.tables[index]) != entry["table_text_sha256"]:
            raise FormatMonographError(f"Structure-map table hash mismatch: {index}")
    numbering = structure_map.get("numbering", {})
    if numbering.get("approved"):
        chapter_start = int(numbering["chapter_start"])
        approved_chapters = [
            _verified_locator_paragraph(document, entry).text
            for entry in structure_map.get("headings", [])
            if entry.get("approved") and int(entry.get("level", 0)) == 1
        ]
        detected = [_chapter_number(value) for value in approved_chapters]
        if detected and detected[0] != chapter_start:
            raise FormatMonographError(
                "Approved chapter_start does not match the first approved chapter heading."
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
        paragraph = (
            _verified_locator_paragraph(document, entry)
            if entry.get("locator")
            else _verified_paragraph(document, entry)
        )
        paragraph.style = ensure_paragraph_style(document, f"Heading {level}")
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


def _move_caption_before_table(
    document: Any, paragraph: Any, locator: dict[str, Any]
) -> Any:
    if locator.get("kind") != "table_cell_paragraph":
        return paragraph
    table_index = int(locator["table"])
    row_index = int(locator["row"])
    table = document.tables[table_index]
    row = table.rows[row_index]
    unique_cells = {id(cell._tc): cell for cell in row.cells}
    if len(unique_cells) != 1:
        raise FormatMonographError("Approved caption row must contain one merged cell.")
    cell = next(iter(unique_cells.values()))
    other_text = [item.text for item in cell.paragraphs if item._p is not paragraph._p]
    if any(value.strip() for value in other_text):
        raise FormatMonographError("Approved caption row contains additional authored text.")

    key = _locator_key(locator)
    table._tbl.addprevious(paragraph._p)
    moved = Paragraph(paragraph._p, table._parent)
    table._tbl.remove(row._tr)
    cache = getattr(document, "_format_monograph_locator_cache", {})
    cache[key] = moved
    setattr(document, "_format_monograph_locator_cache", cache)
    deleted_rows = getattr(document, "_format_monograph_deleted_table_rows", {})
    deleted_rows[table_index] = row_index
    setattr(document, "_format_monograph_deleted_table_rows", deleted_rows)
    return moved


def _apply_captions(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("captions", []):
        if not entry.get("approved"):
            continue
        paragraph = (
            _verified_locator_paragraph(document, entry)
            if entry.get("locator")
            else _verified_paragraph(document, entry)
        )
        match = CAPTION_PATTERN.match(paragraph.text)
        if not match:
            raise FormatMonographError(
                f"Approved caption no longer matches at paragraph {entry['paragraph']}."
            )
        label, hierarchy, sequence, _ = match.groups()
        if label != entry["label"]:
            raise FormatMonographError("Approved caption label does not match.")
        if entry.get("migrate_outside_table"):
            paragraph = _move_caption_before_table(
                document, paragraph, entry["locator"]
            )
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
        paragraph.style = ensure_paragraph_style(document, "Caption")
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
        if structure_map.get("schema_version") == "1.1" and entry.get("kind") != "data":
            raise FormatMonographError(
                "Only tables approved with kind=data may receive data-table rules."
            )
        index = int(entry["table"])
        if not 0 <= index < len(document.tables):
            raise FormatMonographError(f"Structure-map table is out of range: {index}")
        table = document.tables[index]
        expected_table_hash = entry.get("table_text_sha256")
        if expected_table_hash and _table_text_hash(table) != expected_table_hash:
            raise FormatMonographError(f"Structure-map table hash mismatch: {index}")
        first_row = "\u241f".join(cell.text for cell in table.rows[0].cells) if table.rows else ""
        if text_sha256(first_row) != entry["first_row_sha256"]:
            raise FormatMonographError(f"Structure-map table hash mismatch: {index}")
        repeat_rows = entry.get("repeat_header_rows", [])
        if not repeat_rows and entry.get("repeat_header"):
            repeat_rows = [0]
        for row_index in repeat_rows:
            row_index = int(row_index)
            if not 0 <= row_index < len(table.rows):
                raise FormatMonographError(
                    f"Approved repeat-header row is out of range: {index}:{row_index}"
                )
            _set_row_property(table.rows[row_index], "tblHeader", True)
        if entry.get("prevent_normal_row_split"):
            caption_row = entry.get("caption_row")
            for row_index, row in enumerate(table.rows):
                if caption_row is not None and row_index == int(caption_row):
                    continue
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
    blocked = ("pgNumType", "titlePg")
    has_header_footer_refs = any(
        final.find(qn(f"w:{name}")) is not None
        for name in ("headerReference", "footerReference")
    )
    evidence = entry.get("evidence", {})
    if (
        any(final.find(qn(f"w:{name}")) is not None for name in blocked)
        or (has_header_footer_refs and evidence.get("header_footer_has_payload", True))
    ):
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
        if (
            structure_map.get("schema_version") == "1.1"
            and not entry.get("evidence", {}).get("safe_to_delete")
        ):
            raise FormatMonographError(
                "Approved trailing section lacks safe deletion evidence."
            )
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
        {"kind": "structure_tables", "targets": _apply_tables(document, structure_map)},
        {"kind": "structure_captions", "targets": _apply_captions(document, structure_map)},
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
        if entry.get("approved") and "paragraph" in entry
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
    migrated_caption_hashes = {
        entry["text_sha256"]
        for entry in structure_map.get("captions", [])
        if entry.get("approved")
        and entry.get("migrate_outside_table")
        and entry.get("locator", {}).get("kind") == "table_cell_paragraph"
    }
    has_migrated_captions = bool(migrated_caption_hashes)
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
                original_caption = CAPTION_PATTERN.match(value)
                if text_sha256(value) in migrated_caption_hashes and original_caption:
                    value = f"{original_caption.group(1)} {original_caption.group(4)}"
                elif has_migrated_captions and direct_body_paragraph:
                    styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
                    generated = re.match(r"^\s*(图|表)\s*[-－—–]\s*(.*)$", value)
                    if styles == ["Caption"] and generated:
                        value = f"{generated.group(1)} {generated.group(2)}"
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
