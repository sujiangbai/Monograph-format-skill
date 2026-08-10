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
    (1, re.compile(r"^\s*绗琝s*[0-9涓€浜屼笁鍥涗簲鍏竷鍏節鍗佺櫨]+\s*绔燶s*\S")),
)
CAPTION_PATTERN = re.compile(
    r"^\s*(鍥緗琛?\s*(\d+(?:\.\d+){1,3})\s*[-锛嶁€斺€揮\s*(\d+)\s*(\S.*)$"
)
LOOSE_CAPTION_PATTERN = re.compile(r"^\s*(鍥緗琛?\s*(\S.*)?$")
CAPTION_IDENTIFIER_PATTERN = re.compile(
    r"^\s*(鍥緗琛?\s*([0-9]+(?:[.锛?][0-9]+)*)"
    r"(?:\s+|\s*[-锛嶁€斺€揮\s*(?=\D))(\S.*)$"
)
DRAWING_MARK_PATTERN = re.compile(
    r"(?:^|\s)\d+(?:-\d+)+\s*(?:鍓栭潰|鏂潰|鎴潰|鑺傜偣|璇﹀浘|澶ф牱|绔嬮潰|骞抽潰)"
)
ARCHITECTURE_TERMS = ("寤虹瓚", "骞抽潰鍥?, "绔嬮潰鍥?, "鍓栭潰鍥?, "璇﹀浘", "澶ф牱")
CIVIL_ENGINEERING_TERMS = (
    "鍦熸湪",
    "缁撴瀯",
    "閽㈢粨鏋?,
    "娣峰嚌鍦?,
    "姊?,
    "鏌?,
    "妗佹灦",
    "鍩虹",
    "鑺傜偣",
    "鎴潰",
    "鏂潰",
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
        return "mixed" if "." in identifier or "锛? in identifier else "drawing_mark"
    if domain_context != "unknown" and "-" in identifier and not any(
        marker in identifier for marker in (".", "锛?)
    ):
        return "drawing_mark"
    if any(marker in identifier for marker in (".", "锛?)):
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
                    if caption_entry.get("action") == "move_caption":
                        if text_sha256(paragraph.text) != caption_entry.get(
                            "text_sha256"
                        ):
                            raise FormatMonographError(
                                "Moved caption target content does not match approval."
                            )
                    elif caption_entry.get("action") == "convert_to_seq":
                        instructions = " ".join(
                            paragraph._p.xpath(
                                ".//w:fldSimple/@w:instr | .//w:instrText/text()"
         …15983 tokens truncated…y
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
                    generated = re.match(r"^\s*(鍥緗琛?\s*[-锛嶁€斺€揮\s*(.*)$", value)
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
                        value = re.sub(r"^(鍥緗琛?\s+[-锛嶁€斺€揮?\s*", r"\1 ", value)
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

