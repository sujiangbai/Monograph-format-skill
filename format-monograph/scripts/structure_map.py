"""Candidate generation, validation, application, and audit normalization."""

from __future__ import annotations

import copy
import hashlib
import json
import posixpath
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
from docx_pagination import apply_pagination_sections


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
CAPTION_IDENTIFIER_PATTERN = re.compile(
    r"^\s*(图|表)\s*([0-9]+(?:[.．-][0-9]+)*)"
    r"(?:\s+|\s*[-－—–]\s*(?=\D))(\S.*)$"
)
DRAWING_MARK_PATTERN = re.compile(
    r"(?:^|\s)\d+(?:-\d+)+\s*(?:剖面|断面|截面|节点|详图|大样|立面|平面)"
)
ARCHITECTURE_TERMS = ("建筑", "平面图", "立面图", "剖面图", "详图", "大样")
CIVIL_ENGINEERING_TERMS = (
    "土木",
    "结构",
    "钢结构",
    "混凝土",
    "梁",
    "柱",
    "桁架",
    "基础",
    "节点",
    "截面",
    "断面",
)
STRUCTURE_MAP_VERSIONS = {"1.0", "1.1", "1.2", "1.3", "1.4"}
SEMANTIC_STRUCTURE_MAP_VERSIONS = {"1.1", "1.2", "1.3", "1.4"}
CAPTION_ACTIONS = {
    "preserve",
    "style_only",
    "replace_identifier",
    "convert_to_seq",
    "move_caption",
}

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
    "table_of_contents_level_1": "toc_level_1",
    "table_of_contents_level_2": "toc_level_2",
    "table_of_contents_level_3": "toc_level_3",
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
    "toc_level_1": "TOC 1",
    "toc_level_2": "TOC 2",
    "toc_level_3": "TOC 3",
}


def text_sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _heading_authored_text_hash(value: str, level: int) -> str:
    match = _heading_prefix_pattern(level).match(value)
    return text_sha256(value[match.end() :] if match else value)


def has_semantic_structure_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in SEMANTIC_STRUCTURE_MAP_VERSIONS


def has_caption_actions_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in {"1.2", "1.3", "1.4"}


def _caption_domain(value: str, context: str = "") -> tuple[str, str]:
    evidence = f"{value}\n{context}"
    architecture = any(term in evidence for term in ARCHITECTURE_TERMS)
    civil = any(term in evidence for term in CIVIL_ENGINEERING_TERMS)
    if architecture and civil:
        return "mixed", "high"
    if architecture:
        return "architecture", "high"
    if civil:
        return "civil_engineering", "high"
    return "unknown", "low"


def _caption_identifier_semantics(
    identifier: str | None, title: str | None, domain_context: str
) -> str:
    if not identifier:
        return "unknown"
    drawing_evidence = bool(DRAWING_MARK_PATTERN.search(title or ""))
    if drawing_evidence:
        return "mixed" if "." in identifier or "．" in identifier else "drawing_mark"
    if domain_context != "unknown" and "-" in identifier and not any(
        marker in identifier for marker in (".", "．")
    ):
        return "drawing_mark"
    if any(marker in identifier for marker in (".", "．")):
        return "publication_number"
    return "unknown"


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


def _nearest_nonempty_hash(values: list[str], index: int, step: int) -> str | None:
    cursor = index + step
    while 0 <= cursor < len(values):
        if values[cursor].strip():
            return text_sha256(values[cursor])
        cursor += step
    return None


def _body_locator(
    index: int, values: list[str] | None = None
) -> dict[str, Any]:
    locator: dict[str, Any] = {"kind": "body_paragraph", "paragraph": index}
    if values is not None:
        locator.update(
            {
                "text_sha256": text_sha256(values[index]),
                "previous_nonempty_sha256": _nearest_nonempty_hash(values, index, -1),
                "next_nonempty_sha256": _nearest_nonempty_hash(values, index, 1),
            }
        )
    return locator


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
        expected = locator.get("text_sha256")
        if 0 <= index < len(document.paragraphs):
            paragraph = document.paragraphs[index]
            if expected is None or text_sha256(paragraph.text) == expected:
                return paragraph
        if expected is None:
            raise FormatMonographError(f"Body paragraph locator is out of range: {index}")

        values = [paragraph.text for paragraph in document.paragraphs]
        candidates = [
            candidate
            for candidate, value in enumerate(values)
            if text_sha256(value) == expected
        ]
        before = locator.get("previous_nonempty_sha256")
        after = locator.get("next_nonempty_sha256")

        def context_score(candidate: int) -> int:
            return int(
                before is not None
                and _nearest_nonempty_hash(values, candidate, -1) == before
            ) + int(
                after is not None
                and _nearest_nonempty_hash(values, candidate, 1) == after
            )

        if len(candidates) > 1:
            scored = [(context_score(candidate), candidate) for candidate in candidates]
            best = max(score for score, _ in scored)
            candidates = [candidate for score, candidate in scored if score == best]
        if len(candidates) == 1:
            return document.paragraphs[candidates[0]]

        # A caller-approved identifier replacement changes the paragraph hash. Its
        # unchanged neighboring paragraphs still provide a stable, text-free anchor.
        contextual = [
            candidate
            for candidate in range(len(values))
            if context_score(candidate) == int(before is not None) + int(after is not None)
        ]
        if len(contextual) == 1:
            return document.paragraphs[contextual[0]]
        raise FormatMonographError(
            f"Stable body paragraph locator is ambiguous or missing: {index}"
        )
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
    locator = copy.deepcopy(entry["locator"])
    if locator.get("kind") == "body_paragraph" and entry.get("text_sha256"):
        locator.setdefault("text_sha256", entry["text_sha256"])
    key = _locator_key(locator)
    verified = getattr(document, "_format_monograph_verified_locators", set())
    if key in verified:
        return resolve_paragraph_locator(document, locator)
    paragraph = resolve_paragraph_locator(document, locator)
    expected = entry.get("text_sha256")
    if expected and text_sha256(paragraph.text) != expected:
        raise FormatMonographError(
            f"Structure-map paragraph hash mismatch at {key}."
        )
    return paragraph


def _replacement_target_matches(value: str, entry: dict[str, Any]) -> bool:
    span = entry["identifier_span"]
    start, end = int(span["start"]), int(span["end"])
    replacement = str(entry["replacement_identifier"])
    replacement_end = start + len(replacement)
    title_start = int(entry["title_span_start"]) + replacement_end - end
    return bool(
        value[start:replacement_end] == replacement
        and text_sha256(value[:start]) == entry["identifier_prefix_sha256"]
        and text_sha256(value[replacement_end:]) == entry["identifier_suffix_sha256"]
        and text_sha256(value[title_start:]) == entry["title_text_sha256"]
    )


def _resolve_replaced_caption_target(document: Any, entry: dict[str, Any]) -> Any:
    candidates = [
        paragraph
        for paragraph in document.paragraphs
        if _replacement_target_matches(paragraph.text, entry)
    ]
    if len(candidates) != 1:
        raise FormatMonographError(
            "Approved caption replacement target is ambiguous or missing."
        )
    return candidates[0]


def prime_structure_map_locators(document: Any, structure_map: dict[str, Any]) -> None:
    if not has_semantic_structure_map(structure_map):
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
    if structure_map.get("schema_version") in {"1.3", "1.4"}:
        for group in structure_map.get("pagination_groups", []):
            if not group.get("approved"):
                continue
            for name in ("anchor", "caption"):
                locator = group.get(name)
                if not isinstance(locator, dict):
                    continue
                paragraph = resolve_paragraph_locator(document, locator)
                expected = locator.get("text_sha256")
                if expected and text_sha256(paragraph.text) != expected:
                    raise FormatMonographError(
                        f"Pagination locator hash mismatch at {_locator_key(locator)}."
                    )
                key = _locator_key(locator)
                cache[key] = paragraph
                verified.add(key)
    if structure_map.get("schema_version") == "1.4":
        pagination = structure_map.get("pagination_sections", {})
        if pagination.get("approved"):
            for name in ("toc_start", "body_start"):
                locator = pagination.get(name)
                if not isinstance(locator, dict):
                    continue
                paragraph = resolve_paragraph_locator(document, locator)
                expected = locator.get("text_sha256")
                if expected and text_sha256(paragraph.text) != expected:
                    raise FormatMonographError(
                        f"Pagination-section locator hash mismatch at {_locator_key(locator)}."
                    )
                key = _locator_key(locator)
                cache[key] = paragraph
                verified.add(key)
        for entry in structure_map.get("trailing_empty_sections", []):
            if not entry.get("approved_delete"):
                continue
            locator = entry.get("previous_boundary_locator")
            if not isinstance(locator, dict):
                continue
            paragraph = resolve_paragraph_locator(document, locator)
            key = _locator_key(locator)
            cache[key] = paragraph
            verified.add(key)
    setattr(document, "_format_monograph_locator_cache", cache)
    setattr(document, "_format_monograph_verified_locators", verified)


def normalized_role(value: str) -> str:
    return ROLE_ALIASES.get(value, value)


def approved_role_paragraphs(
    document: Any, structure_map: dict[str, Any], selector: dict[str, str]
) -> list[Any]:
    if not has_semantic_structure_map(structure_map):
        return []
    wanted = normalized_role(selector["value"])
    caption_roles = {"figure_caption", "table_caption", "equation_caption"}
    result = []
    seen: set[int] = set()
    for entry in structure_map.get("paragraph_roles", []):
        if not entry.get("approved"):
            continue
        locator = entry.get("locator", {})
        if locator.get("kind") == "body_paragraph" and any(
            toc.get("approved")
            and int(toc["start_paragraph"])
            <= int(locator.get("paragraph", -1))
            <= int(toc["end_paragraph"])
            for toc in structure_map.get("toc_ranges", [])
        ):
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
                normalized_hash = entry.get("normalized_text_sha256")
                candidates = [
                    paragraph
                    for paragraph in document.paragraphs
                    if paragraph.style
                    and paragraph.style.name == expected_style
                    and (
                        normalized_hash is None
                        or text_sha256(paragraph.text) == normalized_hash
                    )
                ]
                if len(candidates) != 1:
                    raise
                paragraph = candidates[0]
            elif role in caption_roles:
                caption_entry = next(
                    (
                        item
                        for item in structure_map.get("captions", [])
                        if item.get("approved")
                        and item.get("locator", {}).get("kind")
                        in {"body_paragraph", "table_cell_paragraph"}
                        and _locator_key(item["locator"])
                        == _locator_key(locator)
                    ),
                    None,
                )
                if caption_entry is None:
                    raise
                if (
                    has_caption_actions_map(structure_map)
                    and caption_entry.get("action")
                    not in {"replace_identifier", "convert_to_seq", "move_caption"}
                ):
                    raise
                if caption_entry.get("action") == "replace_identifier":
                    paragraph = _resolve_replaced_caption_target(
                        document, caption_entry
                    )
                elif (
                    locator.get("kind") == "table_cell_paragraph"
                    and caption_entry.get("migrate_outside_table")
                ):
                    table_index = int(locator["table"])
                    if not 0 <= table_index < len(document.tables):
                        raise
                    previous = document.tables[table_index]._tbl.getprevious()
                    if previous is None or previous.tag != qn("w:p"):
                        raise
                    paragraph = Paragraph(previous, document.tables[table_index]._parent)
                else:
                    paragraph = resolve_paragraph_locator(document, locator)
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
    if not has_semantic_structure_map(structure_map):
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


def _section_properties_sha256(sect_pr: Any) -> str:
    canonical = etree.tostring(sect_pr, method="c14n", exclusive=True)
    return hashlib.sha256(canonical).hexdigest()


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
        candidate_sect_pr = list(document.sections)[section_index]._sectPr
        evidence = _section_evidence(document, candidate_sect_pr)
        candidates.append(
            {
                "section": section_index,
                "previous_boundary_paragraph": previous_boundary,
                "previous_boundary_sha256": text_sha256(paragraph.text),
                "previous_boundary_locator": _body_locator(
                    previous_boundary, [item.text for item in document.paragraphs]
                ),
                "section_properties_sha256": _section_properties_sha256(candidate_sect_pr),
                "approved_delete": False,
                "confidence": "high" if evidence["safe_to_delete"] else "low",
                "evidence": evidence,
            }
        )
    return candidates


def _candidate_pagination_sections(
    document: Any, headings: list[dict[str, Any]]
) -> dict[str, Any]:
    body_start = next(
        (entry["locator"] for entry in headings if int(entry["level"]) == 1), None
    )
    toc_start = None
    for index, paragraph in enumerate(document.paragraphs):
        style_name = paragraph.style.name if paragraph.style else ""
        instructions = " ".join(
            paragraph._p.xpath(".//w:fldSimple/@w:instr | .//w:instrText/text()")
        ).upper()
        if "[[TOC]]" in paragraph.text or "TOC" in instructions or style_name.upper().startswith("TOC"):
            toc_start = _body_locator(
                index, [item.text for item in document.paragraphs]
            )
            break
    return {
        "approved": False,
        "toc_start": toc_start,
        "body_start": body_start,
        "number_format": "decimal",
        "start_at": {"toc": 1, "body": 1},
        "continue_after_body_start": True,
        "odd_position": "outer_right",
        "even_position": "outer_left",
        "show_on_first_page": True,
    }


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


def _caption_entry(
    value: str, locator: dict[str, Any], context: str = ""
) -> dict[str, Any] | None:
    loose = LOOSE_CAPTION_PATTERN.match(value)
    if not loose:
        return None

    label = loose.group(1)
    domain_context, domain_confidence = _caption_domain(value, context)
    identifier = CAPTION_IDENTIFIER_PATTERN.match(value)
    entry: dict[str, Any] = {
        "locator": locator,
        "text_sha256": text_sha256(value),
        "label": label,
        "sequence_name": "Figure" if label == "图" else "Table",
        "numbering_mode": "manual_text",
        "identifier_semantics": "unknown",
        "domain_context": domain_context,
        "domain_confidence": domain_confidence,
        "action": "preserve",
        "completeness": "candidate",
        "hierarchy_status": "not_evaluated",
        "approved": False,
    }
    if identifier:
        raw_identifier = identifier.group(2)
        start, end = identifier.span(2)
        title = identifier.group(3)
        entry.update(
            {
                "identifier_span": {"start": start, "end": end},
                "identifier_sha256": text_sha256(raw_identifier),
                "identifier_prefix_sha256": text_sha256(value[:start]),
                "identifier_suffix_sha256": text_sha256(value[end:]),
                "title_text_sha256": text_sha256(title),
                "title_span_start": identifier.start(3),
                "identifier_semantics": _caption_identifier_semantics(
                    raw_identifier, title, domain_context
                ),
            }
        )

    legacy = CAPTION_PATTERN.match(value)
    if legacy:
        _, hierarchy, sequence, _ = legacy.groups()
        entry.update(
            {
                "heading_level": len(hierarchy.split(".")),
                "cached_hierarchy": hierarchy,
                "cached_sequence": sequence,
            }
        )
    return entry


def _table_text_hash(table: Any) -> str:
    value = "\u241e".join(
        "\u241f".join(cell.text for cell in row.cells) for row in table.rows
    )
    return text_sha256(value)


def _candidate_pagination_groups(
    document: Any, captions: list[dict[str, Any]], body_values: list[str]
) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    table_indexes = {id(table._tbl): index for index, table in enumerate(document.tables)}
    for caption in captions:
        locator = caption.get("locator", {})
        if locator.get("kind") != "body_paragraph":
            continue
        index = int(locator["paragraph"])
        if caption.get("sequence_name") == "Figure" and index > 0:
            previous = document.paragraphs[index - 1]
            if previous._p.xpath(".//w:drawing | .//w:pict"):
                groups.append(
                    {
                        "kind": "figure_with_caption",
                        "anchor": _body_locator(index - 1, body_values),
                        "caption": copy.deepcopy(locator),
                        "approved": False,
                    }
                )
        elif caption.get("sequence_name") == "Table":
            sibling = document.paragraphs[index]._p.getnext()
            while sibling is not None and sibling.tag == qn("w:p") and not _paragraph_has_payload(sibling):
                sibling = sibling.getnext()
            if sibling is not None and sibling.tag == qn("w:tbl"):
                table_index = table_indexes.get(id(sibling))
                if table_index is not None:
                    groups.append(
                        {
                            "kind": "table_caption_with_table",
                            "anchor": copy.deepcopy(locator),
                            "table": table_index,
                            "table_text_sha256": _table_text_hash(
                                document.tables[table_index]
                            ),
                            "approved": False,
                        }
                    )
    return groups


def candidate_structure_map(path: Path) -> dict[str, Any]:
    from _common import load_document

    document = load_document(path)
    headings = []
    captions = []
    paragraph_roles = []
    chapter_starts = []
    body_values = [paragraph.text for paragraph in document.paragraphs]
    for index, paragraph in enumerate(document.paragraphs):
        value = paragraph.text
        detected_level = None
        for level, pattern in HEADING_PATTERNS:
            if pattern.match(value):
                detected_level = level
                entry = {
                    "paragraph": index,
                    "locator": _body_locator(index, body_values),
                    "text_sha256": text_sha256(value),
                    "level": level,
                    "cached_number": list(_heading_number(value, level) or ()),
                    "source_style": paragraph.style.name if paragraph.style else None,
                    "direct_format_sha256": _paragraph_style_signature(paragraph),
                    "normalized_text_sha256": _heading_authored_text_hash(
                        value, level
                    ),
                    "approved": False,
                }
                headings.append(entry)
                if level == 1:
                    chapter = _chapter_number(value)
                    if chapter is not None:
                        chapter_starts.append(chapter)
                break
        nearby = "\n".join(body_values[max(0, index - 2) : index + 3])
        caption = _caption_entry(value, _body_locator(index, body_values), nearby)
        if caption:
            caption["paragraph"] = index
            captions.append(caption)
            role = "figure_caption" if caption["label"] == "图" else "table_caption"
        else:
            role = _role_for_paragraph(paragraph, detected_level)
        if value.strip():
            role_entry = {
                    "locator": _body_locator(index, body_values),
                    "text_sha256": text_sha256(value),
                    "role": role,
                    "source_style": paragraph.style.name if paragraph.style else None,
                    "direct_format_sha256": _paragraph_style_signature(paragraph),
                    "approved": False,
                }
            if detected_level is not None:
                role_entry["normalized_text_sha256"] = _heading_authored_text_hash(
                    value, detected_level
                )
            paragraph_roles.append(role_entry)

    tables = []
    for table_index, table in enumerate(document.tables):
        caption_row = None
        header_rows: list[int] = []
        table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
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
                    caption = _caption_entry(paragraph.text, locator, table_text)
                    if not caption:
                        continue
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
        unique_cells = {
            id(cell._tc) for row in table.rows for cell in row.cells
        }
        visible_controls = sum(
            1
            for character in table_text
            if 0x2400 <= ord(character) <= 0x2426
            or (ord(character) < 32 and character not in "\t\n\r")
        )
        has_floating_objects = bool(
            table._tbl.xpath(".//wp:anchor | .//w:object | .//w:pict")
        )
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
                "complex_merge": len(unique_cells)
                < sum(len(row.cells) for row in table.rows),
                "has_floating_objects": has_floating_objects,
                "visible_control_mark_candidates": visible_controls,
                "visual": {
                    "approved": False,
                    "available_width_percent": 100,
                    "allow_autofit": True,
                    "cell_margins_mm": {
                        "top": 1.0,
                        "right": 1.5,
                        "bottom": 1.0,
                        "left": 1.5,
                    },
                    "vertical_alignment": "center",
                    "border_preset": "preserve",
                    "column_roles": ["unknown"] * len(table.columns),
                    "orientation": "portrait",
                    "landscape_approved": False,
                },
                "approved": False,
            }
        )

    return {
        "schema_version": "1.4",
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
        "pagination_groups": _candidate_pagination_groups(document, captions, body_values),
        "pagination_sections": _candidate_pagination_sections(document, headings),
        "trailing_empty_sections": _trailing_empty_sections(document),
        "conflicts": [],
    }


def load_structure_map(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FormatMonographError(f"Invalid structure map: {path}: {exc}") from exc
    if value.get("schema_version") not in STRUCTURE_MAP_VERSIONS:
        raise FormatMonographError(
            "Structure map schema_version must be 1.0, 1.1, 1.2, 1.3, or 1.4."
        )
    if value.get("status") != "approved":
        raise FormatMonographError("Structure map status must be approved.")
    if value.get("conflicts"):
        raise FormatMonographError("Structure map contains unresolved conflicts.")
    if has_semantic_structure_map(value):
        for entry in value.get("paragraph_roles", []):
            if entry.get("approved") and entry.get("role") == "unknown":
                raise FormatMonographError("Approved paragraph role cannot be unknown.")
            if "locator" not in entry or "text_sha256" not in entry:
                raise FormatMonographError(
                    "Semantic structure-map paragraph roles require locator and text_sha256."
                )
        for entry in value.get("captions", []):
            if not entry.get("approved"):
                continue
            if value.get("schema_version") == "1.1":
                if entry.get("completeness") != "complete":
                    raise FormatMonographError("Incomplete captions cannot be approved.")
                if entry.get("hierarchy_status") not in {"match", "accepted"}:
                    raise FormatMonographError(
                        "Caption hierarchy must match or be explicitly accepted."
                    )
                continue

            action = entry.get("action")
            if action not in CAPTION_ACTIONS:
                raise FormatMonographError(
                    f"Structure map caption action is invalid: {action}"
                )
            if entry.get("numbering_mode") not in {"manual_text", "seq_field"}:
                raise FormatMonographError(
                    "Structure map 1.2+ captions require numbering_mode."
                )
            if entry.get("identifier_semantics") not in {
                "publication_number",
                "drawing_mark",
                "mixed",
                "unknown",
            }:
                raise FormatMonographError(
                    "Structure map caption identifier_semantics is invalid."
                )
            if entry.get("domain_context") not in {
                "general",
                "architecture",
                "civil_engineering",
                "mixed",
                "unknown",
            }:
                raise FormatMonographError(
                    "Structure map caption domain_context is invalid."
                )
            if entry.get("domain_confidence") not in {"high", "medium", "low"}:
                raise FormatMonographError(
                    "Structure map caption domain_confidence is invalid."
                )
            if action != "move_caption" and entry.get("migrate_outside_table"):
                raise FormatMonographError(
                    "Caption relocation requires the separate move_caption action."
                )
            if action == "replace_identifier":
                required = {
                    "identifier_span",
                    "identifier_sha256",
                    "identifier_prefix_sha256",
                    "identifier_suffix_sha256",
                    "title_text_sha256",
                    "title_span_start",
                    "replacement_identifier",
                }
                missing = sorted(required - set(entry))
                if missing or entry.get("replacement_confirmed") is not True:
                    raise FormatMonographError(
                        "Manual caption identifier replacement requires an exact boundary, "
                        "hashes, replacement_identifier, and replacement_confirmed=true."
                    )
                if entry.get("numbering_mode") != "manual_text":
                    raise FormatMonographError(
                        "Manual caption identifier replacement requires numbering_mode=manual_text."
                    )
            elif action == "convert_to_seq":
                required = {
                    "heading_level",
                    "cached_hierarchy",
                    "cached_sequence",
                    "sequence_name",
                }
                if required - set(entry):
                    raise FormatMonographError(
                        "SEQ conversion requires an unambiguous legacy caption boundary."
                    )
                if entry.get("numbering_mode") != "seq_field":
                    raise FormatMonographError(
                        "SEQ conversion requires numbering_mode=seq_field."
                    )
                if entry.get("hierarchy_status") not in {"match", "accepted"}:
                    raise FormatMonographError(
                        "SEQ conversion hierarchy must match or be explicitly accepted."
                    )
            elif action == "move_caption":
                if not entry.get("migrate_outside_table"):
                    raise FormatMonographError(
                        "move_caption requires migrate_outside_table=true."
                    )
                if entry.get("numbering_mode") != "manual_text":
                    raise FormatMonographError(
                        "move_caption preserves manual text; field conversion is separate."
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
        if value.get("schema_version") in {"1.3", "1.4"}:
            for entry in value.get("paragraph_roles", []):
                if not entry.get("approved"):
                    continue
                locator = entry.get("locator", {})
                if locator.get("kind") == "body_paragraph" and not locator.get(
                    "text_sha256"
                ):
                    raise FormatMonographError(
                        "Structure map 1.3 body locators require a stable text hash."
                    )
            for table in value.get("tables", []):
                if not table.get("approved"):
                    continue
                if table.get("kind") == "layout" and not table.get(
                    "pagination_only"
                ):
                    raise FormatMonographError(
                        "Approved layout tables must set pagination_only=true."
                    )
                if table.get("repeat_caption_with_header") and not (
                    table.get("caption_row") == 0 and 1 in table.get("header_rows", [])
                ):
                    raise FormatMonographError(
                        "repeat_caption_with_header requires caption row 0 and header row 1."
                    )
            for group in value.get("pagination_groups", []):
                if not group.get("approved"):
                    continue
                if group.get("kind") not in {
                    "figure_with_caption",
                    "table_caption_with_table",
                    "keep_rows_together",
                }:
                    raise FormatMonographError("Invalid approved pagination group kind.")
                anchor = group.get("anchor")
                if group.get("kind") != "keep_rows_together" and not isinstance(
                    anchor, dict
                ):
                    raise FormatMonographError(
                        "Approved pagination group requires an anchor locator."
                    )
        if value.get("schema_version") == "1.4":
            pagination = value.get("pagination_sections", {})
            if pagination.get("approved"):
                required = {
                    "toc_start",
                    "body_start",
                    "number_format",
                    "start_at",
                    "continue_after_body_start",
                    "odd_position",
                    "even_position",
                    "show_on_first_page",
                }
                if required - set(pagination):
                    raise FormatMonographError(
                        "Approved pagination_sections is missing required settings."
                    )
                if pagination.get("number_format") != "decimal":
                    raise FormatMonographError(
                        "V0.2.5 technical-textbook pagination requires decimal numbering."
                    )
                starts = pagination.get("start_at", {})
                if starts != {"toc": 1, "body": 1}:
                    raise FormatMonographError(
                        "TOC and body pagination must each start at 1."
                    )
                if pagination.get("odd_position") != "outer_right" or pagination.get(
                    "even_position"
                ) != "outer_left":
                    raise FormatMonographError(
                        "Mirrored pagination requires odd outer-right and even outer-left."
                    )
                if pagination.get("show_on_first_page") is not True:
                    raise FormatMonographError(
                        "Approved TOC and body first pages must show page numbers."
                    )
                for name in ("toc_start", "body_start"):
                    if not isinstance(pagination.get(name), dict):
                        raise FormatMonographError(
                            f"Approved pagination_sections requires {name} locator."
                        )
            for table in value.get("tables", []):
                visual = table.get("visual", {})
                if not table.get("approved") or not visual.get("approved"):
                    continue
                if table.get("kind") != "data":
                    raise FormatMonographError(
                        "Only approved data tables may receive visual formatting."
                    )
                roles = visual.get("column_roles", [])
                if not roles or any(
                    role not in {"numeric", "unit", "short_code", "narrative"}
                    for role in roles
                ):
                    raise FormatMonographError(
                        "Approved table visuals require a known role for every column."
                    )
                if visual.get("border_preset") not in {
                    "preserve",
                    "three_line",
                    "full_grid",
                }:
                    raise FormatMonographError("Invalid approved table border preset.")
                if visual.get("orientation") == "landscape" and visual.get(
                    "landscape_approved"
                ) is not True:
                    raise FormatMonographError(
                        "Landscape table layout requires landscape_approved=true."
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
    if has_semantic_structure_map(structure_map):
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
    pagination = structure_map.get("pagination_sections", {})
    if pagination.get("approved"):
        starts = [
            resolve_paragraph_locator(document, pagination[name])
            for name in ("toc_start", "body_start")
        ]
        if starts[0]._p.getparent().index(starts[0]._p) >= starts[1]._p.getparent().index(
            starts[1]._p
        ):
            raise FormatMonographError(
                "Approved TOC pagination start must precede the body start."
            )
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


def _complex_field_runs(instruction: str, placeholder: str) -> list[Any]:
    begin_run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    begin_run.append(begin)

    instruction_run = OxmlElement("w:r")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set(
        "{http://www.w3.org/XML/1998/namespace}space", "preserve"
    )
    instruction_text.text = f" {instruction} "
    instruction_run.append(instruction_text)

    separate_run = OxmlElement("w:r")
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run.append(separate)

    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)
    return [
        begin_run,
        instruction_run,
        separate_run,
        _text_run(placeholder),
        end_run,
    ]


def _apply_toc_ranges(document: Any, structure_map: dict[str, Any]) -> int:
    approved = [
        entry for entry in structure_map.get("toc_ranges", []) if entry.get("approved")
    ]
    for entry in approved:
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

    changed = 0
    for entry in sorted(
        approved, key=lambda item: int(item["start_paragraph"]), reverse=True
    ):
        start, end = int(entry["start_paragraph"]), int(entry["end_paragraph"])
        anchor = document.paragraphs[start]
        _clear_paragraph(anchor)
        p_pr = anchor._p.get_or_add_pPr()
        for name in ("pStyle", "outlineLvl", "keepNext", "pageBreakBefore"):
            element = p_pr.find(qn(f"w:{name}"))
            if element is not None:
                p_pr.remove(element)
        levels = int(entry.get("levels", 4))
        anchor._p.extend(
            _complex_field_runs(
                f'TOC \\o "1-{levels}" \\h \\z \\u',
                "Update table of contents",
            )
        )
        for index in range(end, start, -1):
            paragraph = document.paragraphs[index]
            parent = paragraph._p.getparent()
            if parent is None or parent.tag != qn("w:body"):
                raise FormatMonographError(
                    "Approved static TOC range must contain direct body paragraphs."
                )
            parent.remove(paragraph._p)
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


def _replace_caption_identifier(paragraph: Any, entry: dict[str, Any]) -> None:
    span = entry["identifier_span"]
    start, end = int(span["start"]), int(span["end"])
    value = paragraph.text
    if not 0 <= start < end <= len(value):
        raise FormatMonographError("Approved caption identifier span is out of range.")
    if text_sha256(value[start:end]) != entry["identifier_sha256"]:
        raise FormatMonographError("Approved caption identifier hash does not match.")
    if text_sha256(value[:start]) != entry["identifier_prefix_sha256"]:
        raise FormatMonographError("Caption text before the identifier changed.")
    if text_sha256(value[end:]) != entry["identifier_suffix_sha256"]:
        raise FormatMonographError("Caption title or separator changed before replacement.")

    runs = list(paragraph.runs)
    if "".join(run.text for run in runs) != value:
        raise FormatMonographError(
            "Caption identifier crosses an unsupported inline object or hyperlink."
        )
    replacement = str(entry["replacement_identifier"])
    cursor = 0
    inserted = False
    for run in runs:
        run_start, run_end = cursor, cursor + len(run.text)
        cursor = run_end
        overlap_start, overlap_end = max(start, run_start), min(end, run_end)
        if overlap_start >= overlap_end:
            continue
        unsupported = [
            child
            for child in run._r
            if child.tag not in {qn("w:rPr"), qn("w:t")}
        ]
        if unsupported:
            raise FormatMonographError(
                "Caption identifier shares a run with an unsupported inline object."
            )
        local_start, local_end = overlap_start - run_start, overlap_end - run_start
        before, after = run.text[:local_start], run.text[local_end:]
        run.text = before + (replacement if not inserted else "") + after
        inserted = True
    if not inserted:
        raise FormatMonographError("Approved caption identifier could not be replaced.")


def _convert_caption_to_seq(document: Any, paragraph: Any, entry: dict[str, Any]) -> None:
    match = CAPTION_PATTERN.match(paragraph.text)
    if not match:
        raise FormatMonographError(
            f"Approved SEQ caption no longer matches at {entry.get('locator', entry.get('paragraph'))}."
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
    paragraph.style = ensure_paragraph_style(document, "Caption")


def _apply_captions(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    version = structure_map.get("schema_version")
    for entry in structure_map.get("captions", []):
        if not entry.get("approved"):
            continue
        paragraph = (
            _verified_locator_paragraph(document, entry)
            if entry.get("locator")
            else _verified_paragraph(document, entry)
        )
        if version in {"1.2", "1.3", "1.4"}:
            action = entry["action"]
            if action in {"preserve", "style_only"}:
                continue
            if action == "replace_identifier":
                _replace_caption_identifier(paragraph, entry)
                changed += 1
                continue
            if action == "move_caption":
                paragraph = _move_caption_before_table(
                    document, paragraph, entry["locator"]
                )
                paragraph.style = ensure_paragraph_style(document, "Caption")
                changed += 1
                continue
            if action != "convert_to_seq":
                raise FormatMonographError(f"Unsupported caption action: {action}")

        if entry.get("migrate_outside_table"):
            paragraph = _move_caption_before_table(
                document, paragraph, entry["locator"]
            )
        _convert_caption_to_seq(document, paragraph, entry)
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


def _set_paragraph_property(paragraph: Any, name: str, enabled: bool) -> bool:
    p_pr = paragraph._p.get_or_add_pPr()
    element = p_pr.find(qn(f"w:{name}"))
    if enabled and element is None:
        element = OxmlElement(f"w:{name}")
        element.set(qn("w:val"), "true")
        p_pr.append(element)
        return True
    if not enabled and element is not None:
        p_pr.remove(element)
        return True
    return False


def _apply_outline_cleanup(document: Any, structure_map: dict[str, Any]) -> int:
    if structure_map.get("schema_version") not in {"1.3", "1.4"}:
        return 0
    changed = 0
    for entry in structure_map.get("paragraph_roles", []):
        if not entry.get("approved"):
            continue
        role = normalized_role(str(entry.get("role", "unknown")))
        if role.startswith("heading_") or role in {"title", "subtitle"}:
            continue
        paragraph = _verified_locator_paragraph(document, entry)
        touched = False
        p_pr = paragraph._p.get_or_add_pPr()
        outline = p_pr.find(qn("w:outlineLvl"))
        if outline is not None:
            p_pr.remove(outline)
            touched = True
        if paragraph.style and paragraph.style.name in {
            "Heading 1",
            "Heading 2",
            "Heading 3",
            "Heading 4",
        }:
            style_name = ROLE_STYLE_NAMES.get(role, "Normal")
            paragraph.style = ensure_paragraph_style(document, style_name)
            touched = True
        changed += int(touched)
    return changed


def _apply_pagination_groups(document: Any, structure_map: dict[str, Any]) -> int:
    if structure_map.get("schema_version") not in {"1.3", "1.4"}:
        return 0
    changed = 0
    for group in structure_map.get("pagination_groups", []):
        if not group.get("approved"):
            continue
        kind = group["kind"]
        if kind in {"figure_with_caption", "table_caption_with_table"}:
            paragraph = resolve_paragraph_locator(document, group["anchor"])
            changed += int(_set_paragraph_property(paragraph, "keepNext", True))
            continue
        if kind == "keep_rows_together":
            table_index = int(group["table"])
            if not 0 <= table_index < len(document.tables):
                raise FormatMonographError("Pagination table locator is out of range.")
            table = document.tables[table_index]
            if group.get("table_text_sha256") and _table_text_hash(table) != group[
                "table_text_sha256"
            ]:
                raise FormatMonographError("Pagination table hash mismatch.")
            for row in table.rows:
                before = row._tr.get_or_add_trPr().find(qn("w:cantSplit"))
                _set_row_property(row, "cantSplit", True)
                changed += int(before is None)
    return changed


def _section_properties_for_body_child(document: Any, child: Any) -> Any:
    children = list(document.element.body)
    position = children.index(child)
    for candidate in children[position:]:
        if candidate.tag == qn("w:p"):
            p_pr = candidate.find(qn("w:pPr"))
            sect_pr = None if p_pr is None else p_pr.find(qn("w:sectPr"))
            if sect_pr is not None:
                return sect_pr
        elif candidate.tag == qn("w:sectPr"):
            return candidate
    raise FormatMonographError("Table section properties could not be resolved.")


def _section_boundary(properties: Any) -> Any:
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    sect_pr = copy.deepcopy(properties)
    section_type = sect_pr.find(qn("w:type"))
    if section_type is None:
        section_type = OxmlElement("w:type")
        sect_pr.insert(0, section_type)
    section_type.set(qn("w:val"), "nextPage")
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is not None:
        pg_num.attrib.pop(qn("w:start"), None)
    p_pr.append(sect_pr)
    paragraph.append(p_pr)
    return paragraph


def _wrap_table_landscape(document: Any, table: Any) -> None:
    body = document.element.body
    if table._tbl.getparent() is not body:
        raise FormatMonographError(
            "A landscape table must be a top-level body table, not a nested layout table."
        )
    source = _section_properties_for_body_child(document, table._tbl)
    portrait_boundary = _section_boundary(source)
    landscape_boundary = _section_boundary(source)
    landscape = landscape_boundary.find("./w:pPr/w:sectPr", namespaces=NS)
    page_size = landscape.find(qn("w:pgSz"))
    if page_size is None:
        raise FormatMonographError("Landscape table requires explicit page size properties.")
    width, height = page_size.get(qn("w:w")), page_size.get(qn("w:h"))
    if width and height:
        page_size.set(qn("w:w"), height)
        page_size.set(qn("w:h"), width)
    page_size.set(qn("w:orient"), "landscape")
    position = list(body).index(table._tbl)
    body.insert(position, portrait_boundary)
    body.insert(position + 2, landscape_boundary)


def _apply_tables(document: Any, structure_map: dict[str, Any]) -> int:
    changed = 0
    for entry in structure_map.get("tables", []):
        if not entry.get("approved"):
            continue
        if (
            has_semantic_structure_map(structure_map)
            and entry.get("kind") != "data"
            and not (
                structure_map.get("schema_version") in {"1.3", "1.4"}
                and entry.get("kind") == "layout"
                and entry.get("pagination_only")
            )
        ):
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
        if entry.get("repeat_caption_with_header"):
            repeat_rows = sorted({0, *[int(row) for row in repeat_rows], 1})
        for row_index in repeat_rows:
            row_index = int(row_index)
            if not 0 <= row_index < len(table.rows):
                raise FormatMonographError(
                    f"Approved repeat-header row is out of range: {index}:{row_index}"
                )
            _set_row_property(table.rows[row_index], "tblHeader", True)
        if entry.get("prevent_normal_row_split") or entry.get("keep_rows_together"):
            caption_row = entry.get("caption_row")
            for row_index, row in enumerate(table.rows):
                if caption_row is not None and row_index == int(caption_row):
                    continue
                _set_row_property(row, "cantSplit", True)
        visual = entry.get("visual", {})
        if visual.get("approved") and visual.get("orientation") == "landscape":
            _wrap_table_landscape(document, table)
        changed += 1
    return changed


def _final_section_is_empty(document: Any, boundary_index: int) -> bool:
    body = document.element.body
    boundary = document.paragraphs[boundary_index]._p
    children = list(body)
    start = children.index(boundary) + 1
    return not any(_paragraph_has_payload(element) for element in children[start:-1])


def _remove_final_empty_section(document: Any, entry: dict[str, Any]) -> None:
    locator = entry.get("previous_boundary_locator")
    if isinstance(locator, dict):
        boundary = resolve_paragraph_locator(document, locator)
        boundary_index = next(
            (
                index
                for index, paragraph in enumerate(document.paragraphs)
                if paragraph._p is boundary._p
            ),
            -1,
        )
        if boundary_index < 0:
            raise FormatMonographError(
                "Trailing-section boundary no longer belongs to the document."
            )
    else:
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
    expected_section_hash = entry.get("section_properties_sha256")
    if expected_section_hash and _section_properties_sha256(final) != expected_section_hash:
        raise FormatMonographError("Trailing-section properties hash mismatch.")
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
            has_semantic_structure_map(structure_map)
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


def _cleanup_orphan_header_footer_relationships(document: Any) -> int:
    referenced = {
        rel_id
        for sect_pr in document.element.xpath(".//w:sectPr")
        for rel_id in sect_pr.xpath(
            "./w:headerReference/@r:id | ./w:footerReference/@r:id"
        )
    }
    removable = [
        rel_id
        for rel_id, relationship in document.part.rels.items()
        if relationship.reltype.rsplit("/", 1)[-1] in {"header", "footer"}
        and rel_id not in referenced
    ]
    for rel_id in removable:
        document.part.drop_rel(rel_id)
    return len(removable)


def apply_structure_map(document: Any, structure_map: dict[str, Any]) -> list[dict[str, Any]]:
    trailing_targets = _apply_trailing_sections(document, structure_map)
    orphan_parts = _cleanup_orphan_header_footer_relationships(document)
    toc_targets = _apply_toc_ranges(document, structure_map)
    heading_targets = _apply_headings(document, structure_map)
    outline_targets = _apply_outline_cleanup(document, structure_map)
    table_targets = _apply_tables(document, structure_map)
    caption_targets = _apply_captions(document, structure_map)
    pagination_targets = _apply_pagination_groups(document, structure_map)
    pagination_sections = apply_pagination_sections(
        document,
        structure_map.get("pagination_sections", {}),
        resolve_paragraph_locator,
        replace_static_page_text=any(
            entry.get("approved_delete")
            and entry.get("evidence", {}).get("approved_derived_footer_only")
            for entry in structure_map.get("trailing_empty_sections", [])
        ),
    )
    caption_actions: dict[str, int] = {}
    if has_caption_actions_map(structure_map):
        for entry in structure_map.get("captions", []):
            if not entry.get("approved") or entry.get("action") not in {
                "replace_identifier",
                "convert_to_seq",
                "move_caption",
            }:
                continue
            action = str(entry["action"])
            caption_actions[action] = caption_actions.get(action, 0) + 1
    changes = [
        {"kind": "structure_toc", "targets": toc_targets},
        {"kind": "structure_headings", "targets": heading_targets},
        {"kind": "structure_outline_cleanup", "targets": outline_targets},
        {"kind": "structure_tables", "targets": table_targets},
        {
            "kind": "structure_captions",
            "targets": caption_targets,
            "actions": caption_actions,
        },
        {"kind": "structure_pagination", "targets": pagination_targets},
        {
            "kind": "structure_trailing_sections",
            "targets": trailing_targets,
        },
        {
            "kind": "structure_orphan_header_footer_cleanup",
            "targets": orphan_parts,
        },
        {
            "kind": "structure_pagination_sections",
            "targets": 1 if pagination_sections else 0,
            "details": pagination_sections,
        },
    ]
    return [change for change in changes if change["targets"]]


def _approved_indexes(structure_map: dict[str, Any], key: str) -> dict[int, dict[str, Any]]:
    return {
        int(entry["paragraph"]): entry
        for entry in structure_map.get(key, [])
        if entry.get("approved") and "paragraph" in entry
    }


def _normalize_manual_identifier(
    value: str, entries: list[dict[str, Any]]
) -> str:
    for entry in entries:
        span = entry["identifier_span"]
        start, end = int(span["start"]), int(span["end"])
        if text_sha256(value) == entry["text_sha256"]:
            if (
                text_sha256(value[:start]) == entry["identifier_prefix_sha256"]
                and text_sha256(value[start:end]) == entry["identifier_sha256"]
                and text_sha256(value[end:]) == entry["identifier_suffix_sha256"]
            ):
                return value[:start] + "[[CAPTION_IDENTIFIER]]" + value[end:]
        replacement = str(entry["replacement_identifier"])
        replacement_end = start + len(replacement)
        if (
            value[start:replacement_end] == replacement
            and text_sha256(value[:start]) == entry["identifier_prefix_sha256"]
            and text_sha256(value[replacement_end:])
            == entry["identifier_suffix_sha256"]
        ):
            return value[:start] + "[[CAPTION_IDENTIFIER]]" + value[replacement_end:]
    return value


def audit_caption_identifier_replacements(
    original_path: Path, formatted_path: Path, structure_map: dict[str, Any]
) -> list[dict[str, Any]]:
    if not has_caption_actions_map(structure_map):
        return []
    from _common import load_document

    original = load_document(original_path)
    formatted = load_document(formatted_path)
    results = []
    for entry in structure_map.get("captions", []):
        if not entry.get("approved") or entry.get("action") != "replace_identifier":
            continue
        source_locator = copy.deepcopy(entry["locator"])
        if source_locator.get("kind") == "body_paragraph" and entry.get("text_sha256"):
            source_locator.setdefault("text_sha256", entry["text_sha256"])
        source = resolve_paragraph_locator(original, source_locator)
        target = _resolve_replaced_caption_target(formatted, entry)
        span = entry["identifier_span"]
        start, end = int(span["start"]), int(span["end"])
        title_start = int(entry["title_span_start"])
        source_ok = (
            text_sha256(source.text) == entry["text_sha256"]
            and text_sha256(source.text[start:end]) == entry["identifier_sha256"]
            and text_sha256(source.text[title_start:]) == entry["title_text_sha256"]
        )
        expected = (
            source.text[:start]
            + str(entry["replacement_identifier"])
            + source.text[end:]
        )
        replacement_end = start + len(str(entry["replacement_identifier"]))
        title_shift = replacement_end - end
        title_preserved = text_sha256(target.text[title_start + title_shift :]) == entry[
            "title_text_sha256"
        ]
        passed = source_ok and target.text == expected and title_preserved
        results.append(
            {
                "locator": entry["locator"],
                "status": "pass" if passed else "fail",
                "identifier_changed_as_approved": target.text == expected,
                "title_preserved": title_preserved,
            }
        )
    return results


def _toc_field_paragraphs(root: etree._Element) -> set[Any]:
    bodies = root.xpath("/w:document/w:body", namespaces=NS)
    if not bodies:
        return set()
    result: set[Any] = set()
    stack: list[dict[str, Any]] = []
    for paragraph in bodies[0].xpath("./w:p", namespaces=NS):
        if paragraph.xpath(
            './/w:fldSimple[starts-with(translate(normalize-space(@w:instr), "toc", "TOC"), "TOC ")]',
            namespaces=NS,
        ):
            result.add(paragraph)
        if any(item.get("is_toc") and item.get("result") for item in stack):
            result.add(paragraph)
        for element in paragraph.iter():
            if element.tag == qn("w:fldChar"):
                kind = element.get(qn("w:fldCharType"))
                if kind == "begin":
                    stack.append({"parts": [], "is_toc": False, "result": False})
                elif kind == "separate" and stack:
                    instruction = "".join(stack[-1]["parts"]).strip().upper()
                    stack[-1]["is_toc"] = instruction.startswith("TOC ")
                    stack[-1]["result"] = True
                    if stack[-1]["is_toc"]:
                        result.add(paragraph)
                elif kind == "end" and stack:
                    if any(item.get("is_toc") and item.get("result") for item in stack):
                        result.add(paragraph)
                    stack.pop()
            elif element.tag == qn("w:instrText") and stack and not stack[-1]["result"]:
                stack[-1]["parts"].append(element.text or "")
            if any(item.get("is_toc") and item.get("result") for item in stack):
                result.add(paragraph)
    return result


def _legacy_toc_empty_anchors(
    root: etree._Element,
    toc_field_paragraphs: set[Any],
    structure_map: dict[str, Any],
) -> set[Any]:
    """Recognize empty anchors left by pre-1.3 static-TOC migration."""
    if structure_map.get("schema_version") in {"1.3", "1.4"}:
        return set()
    if not toc_field_paragraphs:
        return set()
    approved_lengths = [
        int(entry["end_paragraph"]) - int(entry["start_paragraph"])
        for entry in structure_map.get("toc_ranges", [])
        if entry.get("approved")
    ]
    remaining = max(approved_lengths, default=0)
    if remaining <= 0:
        return set()
    body = root.find(qn("w:body"))
    if body is None:
        return set()
    children = list(body)
    field_positions = [
        index for index, child in enumerate(children) if child in toc_field_paragraphs
    ]
    if not field_positions:
        return set()
    result = set()
    for child in children[max(field_positions) + 1 :]:
        if remaining <= 0 or child.tag != qn("w:p"):
            break
        value = FIELD_MARKER_PATTERN.sub(
            "", _paragraph_text_without_field_results(child)
        )
        if value:
            break
        result.add(child)
        remaining -= 1
    return result


def _pagination_boundary_paragraphs(
    root: etree._Element, structure_map: dict[str, Any]
) -> set[Any]:
    pagination = structure_map.get("pagination_sections", {})
    if structure_map.get("schema_version") != "1.4":
        return set()
    result = set()
    has_landscape_table = any(
        entry.get("approved")
        and entry.get("visual", {}).get("approved")
        and entry.get("visual", {}).get("orientation") == "landscape"
        for entry in structure_map.get("tables", [])
    )
    if has_landscape_table:
        result.update(
            paragraph
            for paragraph in root.xpath("/w:document/w:body/w:p", namespaces=NS)
            if paragraph.find("./w:pPr/w:sectPr", namespaces=NS) is not None
            and not _paragraph_text_without_field_results(paragraph)
        )
    if not pagination.get("approved"):
        return result
    body_locator = pagination.get("body_start", {})
    expected = body_locator.get("text_sha256")
    if not expected:
        return result
    expected_hashes = {expected}
    heading = next(
        (
            entry
            for entry in structure_map.get("headings", [])
            if entry.get("approved")
            and entry.get("locator", {}).get("text_sha256") == expected
        ),
        None,
    )
    if heading and heading.get("normalized_text_sha256"):
        expected_hashes.add(heading["normalized_text_sha256"])
    paragraphs = root.xpath("/w:document/w:body/w:p", namespaces=NS)
    matches = [
        paragraph
        for paragraph in paragraphs
        if text_sha256(_paragraph_text_without_field_results(paragraph))
        in expected_hashes
    ]
    if len(matches) != 1:
        return result
    previous = matches[0].getprevious()
    if (
        previous is not None
        and previous.tag == qn("w:p")
        and previous.find("./w:pPr/w:sectPr", namespaces=NS) is not None
        and not _paragraph_text_without_field_results(previous)
    ):
        result.add(previous)
    return result


def _approved_tail_paragraphs(
    root: etree._Element, structure_map: dict[str, Any]
) -> set[Any]:
    approved = [
        entry
        for entry in structure_map.get("trailing_empty_sections", [])
        if entry.get("approved_delete")
    ]
    if not approved:
        return set()
    body = root.find(qn("w:body"))
    if body is None:
        return set()
    direct_paragraphs = [child for child in body if child.tag == qn("w:p")]
    first_removed_positions = []
    children = list(body)
    for entry in approved:
        expected_hash = entry.get("previous_boundary_sha256")
        boundary_is_empty = expected_hash == text_sha256("")
        matching = [
            paragraph
            for paragraph in direct_paragraphs
            if expected_hash and text_sha256(_paragraph_text_without_field_results(paragraph))
            == expected_hash
        ]
        if len(matching) == 1:
            boundary_paragraph = matching[0]
        else:
            boundary = int(entry["previous_boundary_paragraph"])
            if not 0 <= boundary < len(direct_paragraphs):
                continue
            candidate = direct_paragraphs[boundary]
            if expected_hash and text_sha256(
                _paragraph_text_without_field_results(candidate)
            ) != expected_hash:
                continue
            boundary_paragraph = candidate
        if boundary_is_empty and any(
            item.get("previous_boundary_sha256") != text_sha256("")
            for item in approved
        ):
            continue
        boundary_position = children.index(boundary_paragraph)
        first_removed_positions.append(
            boundary_position if boundary_is_empty else boundary_position + 1
        )
    if not first_removed_positions:
        return set()
    first_removed = min(first_removed_positions)
    return {
        paragraph
        for child in children[first_removed:]
        for paragraph in child.xpath(".//w:p | self::w:p", namespaces=NS)
    }


def _approved_deleted_header_footer_parts(
    package: zipfile.ZipFile,
    document_root: etree._Element,
    structure_map: dict[str, Any],
) -> set[str]:
    deleted_sections = {
        int(entry["section"])
        for entry in structure_map.get("trailing_empty_sections", [])
        if entry.get("approved_delete")
    }
    if not deleted_sections:
        return set()
    sections = document_root.xpath(".//w:sectPr", namespaces=NS)
    deleted_ids: set[str] = set()
    retained_ids: set[str] = set()
    relationship_attribute = qn("r:id")
    for index, section in enumerate(sections):
        target = deleted_ids if index in deleted_sections else retained_ids
        for reference in section.xpath(
            "./w:headerReference | ./w:footerReference", namespaces=NS
        ):
            relationship_id = reference.get(relationship_attribute)
            if relationship_id:
                target.add(relationship_id)
    removable_ids = deleted_ids - retained_ids
    if not removable_ids:
        return set()
    relationships_name = "word/_rels/document.xml.rels"
    if relationships_name not in package.namelist():
        return set()
    relationships = etree.fromstring(package.read(relationships_name))
    result = set()
    for relationship in relationships:
        if relationship.get("Id") not in removable_ids:
            continue
        target = relationship.get("Target", "")
        normalized = posixpath.normpath(posixpath.join("word", target))
        if re.fullmatch(r"word/(?:header|footer)\d+\.xml", normalized):
            result.add(normalized)
    return result


def structure_content_inventory(
    path: Path, structure_map: dict[str, Any]
) -> dict[str, list[str]]:
    toc_indexes = {
        index
        for entry in structure_map.get("toc_ranges", [])
        if entry.get("approved")
        for index in range(int(entry["start_paragraph"]), int(entry["end_paragraph"]) + 1)
    }
    heading_entries = [
        entry for entry in structure_map.get("headings", []) if entry.get("approved")
    ]
    caption_entries = [
        entry
        for entry in structure_map.get("captions", [])
        if entry.get("approved")
        and (
            not has_caption_actions_map(structure_map)
            or entry.get("action") == "convert_to_seq"
        )
    ]
    identifier_replacements = [
        entry
        for entry in structure_map.get("captions", [])
        if entry.get("approved") and entry.get("action") == "replace_identifier"
    ]
    migrated_caption_hashes = {
        entry["text_sha256"]
        for entry in structure_map.get("captions", [])
        if entry.get("approved")
        and entry.get("migrate_outside_table")
        and entry.get("locator", {}).get("kind") == "table_cell_paragraph"
    }
    manual_migrated_caption_hashes = {
        entry["text_sha256"]
        for entry in structure_map.get("captions", [])
        if entry.get("approved")
        and entry.get("action") == "move_caption"
        and entry.get("migrate_outside_table")
    }
    has_migrated_captions = bool(
        migrated_caption_hashes - manual_migrated_caption_hashes
    )
    result: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        document_root = etree.fromstring(package.read("word/document.xml"))
        ignored_parts = _approved_deleted_header_footer_parts(
            package, document_root, structure_map
        )
        for name in sorted(package.namelist()):
            if not CONTENT_PART.match(name) or name in ignored_parts:
                continue
            root = document_root if name == "word/document.xml" else etree.fromstring(
                package.read(name)
            )
            toc_field_paragraphs = (
                _toc_field_paragraphs(root) if name == "word/document.xml" else set()
            )
            legacy_toc_anchors = (
                _legacy_toc_empty_anchors(root, toc_field_paragraphs, structure_map)
                if name == "word/document.xml"
                else set()
            )
            approved_tail_paragraphs = (
                _approved_tail_paragraphs(root, structure_map)
                if name == "word/document.xml"
                else set()
            )
            pagination_boundaries = (
                _pagination_boundary_paragraphs(root, structure_map)
                if name == "word/document.xml"
                else set()
            )
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
                if paragraph in approved_tail_paragraphs or paragraph in pagination_boundaries:
                    continue
                if paragraph in toc_field_paragraphs or paragraph in legacy_toc_anchors:
                    continue
                if (
                    current_index is not None
                    and not toc_field_paragraphs
                    and current_index in toc_indexes
                ):
                    continue
                value = _paragraph_text_without_field_results(paragraph)
                value = FIELD_MARKER_PATTERN.sub("", value)
                value = _normalize_manual_identifier(value, identifier_replacements)
                original_caption = CAPTION_PATTERN.match(value)
                if text_sha256(value) in manual_migrated_caption_hashes:
                    value = "[[MOVED_MANUAL_CAPTION]]"
                elif text_sha256(value) in migrated_caption_hashes and original_caption:
                    value = f"{original_caption.group(1)} {original_caption.group(4)}"
                elif (has_migrated_captions or caption_entries) and direct_body_paragraph:
                    styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
                    generated = re.match(r"^\s*(图|表)\s*[-－—–]\s*(.*)$", value)
                    if styles == ["Caption"] and generated:
                        value = f"{generated.group(1)} {generated.group(2)}"
                heading_entry = next(
                    (
                        entry
                        for entry in heading_entries
                        if text_sha256(value) == entry.get("text_sha256")
                    ),
                    None,
                )
                if heading_entry is not None:
                    match = _heading_prefix_pattern(int(heading_entry["level"])).match(value)
                    if match:
                        value = value[match.end() :]
                else:
                    styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
                    style_match = (
                        re.fullmatch(r"Heading([1-4])", styles[0]) if styles else None
                    )
                    if style_match:
                        match = _heading_prefix_pattern(int(style_match.group(1))).match(
                            value
                        )
                        if match:
                            value = value[match.end() :]
                if any(
                    text_sha256(value) == entry.get("text_sha256")
                    for entry in caption_entries
                ):
                    match = CAPTION_PATTERN.match(value)
                    if match:
                        value = f"{match.group(1)} {match.group(4)}"
                    else:
                        value = re.sub(r"^(图|表)\s+[-－—–]?\s*", r"\1 ", value)
                values.append(value)
            if name == "word/document.xml" or any(values):
                result[name] = values
    approved_derived_footer_only = bool(
        structure_map.get("pagination_sections", {}).get("approved")
        and any(
            entry.get("approved_delete")
            and entry.get("evidence", {}).get("approved_derived_footer_only")
            for entry in structure_map.get("trailing_empty_sections", [])
        )
    )
    canonical: dict[str, list[str]] = {}
    headers: list[str] = []
    footers: list[str] = []
    for name, values in result.items():
        if re.fullmatch(r"word/header\d+\.xml", name):
            headers.extend(values)
        elif re.fullmatch(r"word/footer\d+\.xml", name):
            if not approved_derived_footer_only:
                footers.extend(values)
        else:
            canonical[name] = values
    canonical["word/_headers_by_content"] = sorted(headers)
    canonical["word/_footers_by_content"] = sorted(footers)
    return canonical


def structure_content_fingerprint(path: Path, structure_map: dict[str, Any]) -> str:
    result = structure_content_inventory(path, structure_map)
    encoded = json.dumps(result, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()
