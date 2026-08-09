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


HEADING_PATTERNS = (
    (4, re.compile(r"^\s*\d+\.\d+\.\d+\.\d+\.?\s*\S")),
    (3, re.compile(r"^\s*\d+\.\d+\.\d+\.?\s*\S")),
    (2, re.compile(r"^\s*\d+\.\d+\.?\s*\S")),
    (1, re.compile(r"^\s*ç¬¬\s*[0-9ä¸€äºŒä¸‰å››äº”å…­ä¸ƒå…«ä¹åç™¾]+\s*ç« \s*\S")),
)
CAPTION_PATTERN = re.compile(
    r"^\s*(å›¾|è¡¨)\s*(\d+(?:\.\d+){1,3})\s*[-ï¼â€”â€“]\s*(\d+)\s*(\S.*)$"
)
LOOSE_CAPTION_PATTERN = re.compile(r"^\s*(å›¾|è¡¨)\s*(\S.*)?$")
CAPTION_IDENTIFIER_PATTERN = re.compile(
    r"^\s*(å›¾|è¡¨)\s*([0-9]+(?:[.ï¼Ž-][0-9]+)*)"
    r"(?:\s+|\s*[-ï¼â€”â€“]\s*(?=\D))(\S.*)$"
)
DRAWING_MARK_PATTERN = re.compile(
    r"(?:^|\s)\d+(?:-\d+)+\s*(?:å‰–é¢|æ–­é¢|æˆªé¢|èŠ‚ç‚¹|è¯¦å›¾|å¤§æ ·|ç«‹é¢|å¹³é¢)"
)
ARCHITECTURE_TERMS = ("å»ºç­‘", "å¹³é¢å›¾", "ç«‹é¢å›¾", "å‰–é¢å›¾", "è¯¦å›¾", "å¤§æ ·")
CIVIL_ENGINEERING_TERMS = (
    "åœŸæœ¨",
    "ç»“æž„",
    "é’¢ç»“æž„",
    "æ··å‡åœŸ",
    "æ¢",
    "æŸ±",
    "æ¡æž¶",
    "åŸºç¡€",
    "èŠ‚ç‚¹",
    "æˆªé¢",
    "æ–­é¢",
)
STRUCTURE_MAP_VERSIONS = {"1.0", "1.1", "1.2", "1.3"}
SEMANTIC_STRUCTURE_MAP_VERSIONS = {"1.1", "1.2", "1.3"}
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


def _heading_authored_text_hash(value: str, level: int) -> str:
    match = _heading_prefix_pattern(level).match(value)
    return text_sha256(value[match.end() :] if match else value)


def has_semantic_structure_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in SEMANTIC_STRUCTURE_MAP_VERSIONS


def has_caption_actions_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in {"1.2", "1.3"}


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
        return "mixed" if "." in identifier or "ï¼Ž" in identifier else "drawing_mark"
    if domain_context != "unknown" and "-" in identifier and not any(
        marker in identifier for marker in (".", "ï¼Ž")
    ):
        return "drawing_mark"
    if any(marker in identifier for marker in (".", "ï¼Ž")):
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
    if structure_map.get("schema_version") == "1.3":
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
  ×Í}îÚ$z{-®éÜj×÷&–v–æÅ÷Fƒ¢F‚Âf÷&ÖGFVE÷Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•Ð¢’ÓâÆ—7E¶F–7E·7G"Âç•ÕÓ ¢–bæ÷B†5ö6F–öåö7F–öç5öÖ‡7G'V7GW&UöÖ“ ¢&WGW&âµÐ¢g&öÒö6öÖÖöâ–×÷'BÆöEöFö7VÖVç@ ¢÷&–v–æÂÒÆöEöFö7VÖVçB†÷&–v–æÅ÷F‚¢f÷&ÖGFVBÒÆöEöFö7VÖVçB†f÷&ÖGFVE÷F‚¢&W7VÇG2ÒµÐ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒ“ ¢–bæ÷BVçG'’ævWB‚&&÷fVB"’÷"VçG'’ævWB‚&7F–öâ"’Ò'&WÆ6Uö–FVçF–f–W"# ¢6öçF–çVP¢6÷W&6UöÆö6F÷"Ò6÷’æFVW6÷’†VçG'•²&Æö6F÷"%Ò¢–b6÷W&6UöÆö6F÷"ævWB‚&¶–æB"’ÓÒ&&öG•÷&w&‚"æBVçG'’ævWB‚'FW‡E÷6†#Sb"“ ¢6÷W&6UöÆö6F÷"ç6WFFVfVÇB‚'FW‡E÷6†#Sb"ÂVçG'•²'FW‡E÷6†#Sb%Ò¢6÷W&6RÒ&W6öÇfU÷&w&…öÆö6F÷"†÷&–v–æÂÂ6÷W&6UöÆö6F÷"¢F&vWBÒ÷&W6öÇfU÷&WÆ6VEö6F–öå÷F&vWB†f÷&ÖGFVBÂVçG'’¢7âÒVçG'•²&–FVçF–f–W%÷7â%Ð¢7F'BÂVæBÒ–çB‡7å²'7F'B%Ò’Â–çB‡7å²&VæB%Ò¢F—FÆU÷7F'BÒ–çB†VçG'•²'F—FÆU÷7å÷7F'B%Ò¢6÷W&6Uöö²Ò€¢FW‡E÷6†#Sb‡6÷W&6RçFW‡B’ÓÒVçG'•²'FW‡E÷6†#Sb%Ð¢æBFW‡E÷6†#Sb‡6÷W&6RçFW‡E·7F'C¦VæEÒ’ÓÒVçG'•²&–FVçF–f–W%÷6†#Sb%Ð¢æBFW‡E÷6†#Sb‡6÷W&6RçFW‡E·F—FÆU÷7F'C¥Ò’ÓÒVçG'•²'F—FÆU÷FW‡E÷6†#Sb%Ð¢¢W‡V7FVBÒ€¢6÷W&6RçFW‡E³§7F'EÐ¢²7G"†VçG'•²'&WÆ6VÖVçEö–FVçF–f–W"%Ò¢²6÷W&6RçFW‡E¶VæC¥Ð¢¢&WÆ6VÖVçEöVæBÒ7F'B²ÆVâ‡7G"†VçG'•²'&WÆ6VÖVçEö–FVçF–f–W"%Ò’¢F—FÆU÷6†–gBÒ&WÆ6VÖVçEöVæBÒVæ@¢F—FÆU÷&W6W'fVBÒFW‡E÷6†#Sb‡F&vWBçFW‡E·F—FÆU÷7F'B²F—FÆU÷6†–gB¥Ò’ÓÒVçG'•°¢'F—FÆU÷FW‡E÷6†#Sb ¢Ð¢76VBÒ6÷W&6Uöö²æBF&vWBçFW‡BÓÒW‡V7FVBæBF—FÆU÷&W6W'fV@¢&W7VÇG2æVæB€¢°¢&Æö6F÷"#¢VçG'•²&Æö6F÷"%ÒÀ¢'7FGW2#¢'72"–b76VBVÇ6R&f–Â"À¢&–FVçF–f–W%ö6†ævVEö5ö&÷fVB#¢F&vWBçFW‡BÓÒW‡V7FVBÀ¢'F—FÆU÷&W6W'fVB#¢F—FÆU÷&W6W'fVBÀ¢Ð¢¢&WGW&â&W7VÇG0  ¦FVb÷Fö5öf–VÆE÷&w&‡2‡&ö÷C¢WG&VRåôVÆVÖVçB’Óâ6WE´ç•Ó ¢&öF–W2Ò&ö÷Bç‡F‚‚"÷s¦Fö7VÖVçB÷s¦&öG’"ÂæÖW76W3Ôå2¢–bæ÷B&öF–W3 ¢&WGW&â6WB‚¢&W7VÇC¢6WE´ç•ÒÒ6WB‚¢7F6³¢Æ—7E¶F–7E·7G"Âç•ÕÒÒµÐ¢f÷"&w&‚–â&öF–W5³Òç‡F‚‚"â÷s§"ÂæÖW76W3Ôå2“ ¢–b&w&‚ç‡F‚€¢râò÷s¦fÆE6–×ÆU·7F'G2×v—F‚‡G&ç6ÆFR†æ÷&ÖÆ—¦R×76R„s¦–ç7G"’Â'Fö2"Â%Dô2"’Â%Dô2"•ÒrÀ¢æÖW76W3Ôå2À¢“ ¢&W7VÇBæFB‡&w&‚¢–bç’†—FVÒævWB‚&—5÷Fö2"’æB—FVÒævWB‚'&W7VÇB"’f÷"—FVÒ–â7F6²“ ¢&W7VÇBæFB‡&w&‚¢f÷"VÆVÖVçB–â&w&‚æ—FW"‚“ ¢–bVÆVÖVçBçFrÓÒâ‚'s¦fÆD6†""“ ¢¶–æBÒVÆVÖVçBævWB‡â‚'s¦fÆD6†%G—R"’¢–b¶–æBÓÒ&&Vv–â# ¢7F6²æVæB‡²''G2#¢µÒÂ&—5÷Fö2#¢fÇ6RÂ'&W7VÇB#¢fÇ6WÒ¢VÆ–b¶–æBÓÒ'6W&FR"æB7F6³ ¢–ç7G'V7F–öâÒ""æ¦ö–â‡7F6µ²ÓÕ²''G2%Ò’ç7G&—‚’çWW"‚¢7F6µ²ÓÕ²&—5÷Fö2%ÒÒ–ç7G'V7F–öâç7F'G7v—F‚‚%Dô2"¢7F6µ²ÓÕ²'&W7VÇB%ÒÒG'VP¢–b7F6µ²ÓÕ²&—5÷Fö2%Ó ¢&W7VÇBæFB‡&w&‚¢VÆ–b¶–æBÓÒ&VæB"æB7F6³ ¢–bç’†—FVÒævWB‚&—5÷Fö2"’æB—FVÒævWB‚'&W7VÇB"’f÷"—FVÒ–â7F6²“ ¢&W7VÇBæFB‡&w&‚¢7F6²ç÷‚¢VÆ–bVÆVÖVçBçFrÓÒâ‚'s¦–ç7G%FW‡B"’æB7F6²æBæ÷B7F6µ²ÓÕ²'&W7VÇB%Ó ¢7F6µ²ÓÕ²''G2%ÒæVæB†VÆVÖVçBçFW‡B÷"""¢–bç’†—FVÒævWB‚&—5÷Fö2"’æB—FVÒævWB‚'&W7VÇB"’f÷"—FVÒ–â7F6²“ ¢&W7VÇBæFB‡&w&‚¢&WGW&â&W7VÇ@  ¦FVböÆVv7•÷Fö5öV×G•öæ6†÷'2€¢&ö÷C¢WG&VRåôVÆVÖVçBÀ¢Fö5öf–VÆE÷&w&‡3¢6WE´ç•ÒÀ¢7G'V7GW&UöÖ¢F–7E·7G"Âç•ÒÀ¢’Óâ6WE´ç•Ó ¢""%&V6övæ—¦RV×G’æ6†÷'2ÆVgB'’&RÓã27FF–2ÕDô2Ö–w&F–öââ"" ¢–bæ÷BFö5öf–VÆE÷&w&‡3 ¢&WGW&â6WB‚¢&÷fVEöÆVæwF‡2Ò°¢–çB†VçG'•²&VæE÷&w&‚%Ò’Ò–çB†VçG'•²'7F'E÷&w&‚%Ò¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'Fö5÷&ævW2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVB"¢Ð¢&VÖ–æ–ærÒÖ‚†&÷fVEöÆVæwF‡2ÂFVfVÇCÓ¢–b&VÖ–æ–ærÃÒ ¢&WGW&â6WB‚¢&öG’Ò&ö÷Bæf–æB‡â‚'s¦&öG’"’¢–b&öG’—2æöæS ¢&WGW&â6WB‚¢6†–ÆG&VâÒÆ—7B†&öG’¢f–VÆE÷÷6—F–öç2Ò°¢–æFW‚f÷"–æFW‚Â6†–ÆB–âVçVÖW&FR†6†–ÆG&Vâ’–b6†–ÆB–âFö5öf–VÆE÷&w&‡0¢Ð¢–bæ÷Bf–VÆE÷÷6—F–öç3 ¢&WGW&â6WB‚¢&W7VÇBÒ6WB‚¢f÷"6†–ÆB–â6†–ÆG&Vå¶Ö‚†f–VÆE÷÷6—F–öç2’²¥Ó ¢–b&VÖ–æ–ærÃÒ÷"6†–ÆBçFrÒâ‚'s§"“ ¢'&V°¢fÇVRÒd”TÄEôÔ$´U%õEDU$âç7V"€¢""Â÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2†6†–ÆB¢¢–bfÇVS ¢'&V°¢&W7VÇBæFB†6†–ÆB¢&VÖ–æ–ærÓÒ¢&WGW&â&W7VÇ@  ¦FVbö&÷fVE÷F–Å÷&w&‡2€¢&ö÷C¢WG&VRåôVÆVÖVçBÂ7G'V7GW&UöÖ¢F–7E·7G"Âç•Ð¢’Óâ6WE´ç•Ó ¢&÷fVBÒ°¢VçG'¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"¢Ð¢–bæ÷B&÷fVC ¢&WGW&â6WB‚¢&öG’Ò&ö÷Bæf–æB‡â‚'s¦&öG’"’¢–b&öG’—2æöæS ¢&WGW&â6WB‚¢F—&V7E÷&w&‡2Ò¶6†–ÆBf÷"6†–ÆB–â&öG’–b6†–ÆBçFrÓÒâ‚'s§"•Ð¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2ÒµÐ¢6†–ÆG&VâÒÆ—7B†&öG’¢f÷"VçG'’–â&÷fVC ¢W‡V7FVEö†6‚ÒVçG'’ævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"¢&÷VæF'•ö—5öV×G’ÒW‡V7FVEö†6‚ÓÒFW‡E÷6†#Sb‚""¢ÖF6†–ærÒ°¢&w&€¢f÷"&w&‚–âF—&V7E÷&w&‡0¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb…÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚’¢ÓÒW‡V7FVEö†6€¢Ð¢–bÆVâ†ÖF6†–ær’ÓÒ ¢&÷VæF'•÷&w&‚ÒÖF6†–æu³Ð¢VÇ6S ¢&÷VæF'’Ò–çB†VçG'•²'&Wf–÷W5ö&÷VæF'•÷&w&‚%Ò¢–bæ÷BÃÒ&÷VæF'’ÂÆVâ†F—&V7E÷&w&‡2“ ¢6öçF–çVP¢6æF–FFRÒF—&V7E÷&w&‡5¶&÷VæF'•Ð¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb€¢÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2†6æF–FFR¢’ÒW‡V7FVEö†6ƒ ¢6öçF–çVP¢&÷VæF'•÷&w&‚Ò6æF–FFP¢–b&÷VæF'•ö—5öV×G’æBç’€¢—FVÒævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"’ÒFW‡E÷6†#Sb‚""¢f÷"—FVÒ–â&÷fV@¢“ ¢6öçF–çVP¢&÷VæF'•÷÷6—F–öâÒ6†–ÆG&Vâæ–æFW‚†&÷VæF'•÷&w&‚¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2æVæB€¢&÷VæF'•÷÷6—F–öâ–b&÷VæF'•ö—5öV×G’VÇ6R&÷VæF'•÷÷6—F–öâ²¢¢–bæ÷Bf—'7E÷&VÖ÷fVE÷÷6—F–öç3 ¢&WGW&â6WB‚¢f—'7E÷&VÖ÷fVBÒÖ–â†f—'7E÷&VÖ÷fVE÷÷6—F–öç2¢&WGW&â°¢&w&€¢f÷"6†–ÆB–â6†–ÆG&Vå¶f—'7E÷&VÖ÷fVC¥Ð¢f÷"&w&‚–â6†–ÆBç‡F‚‚"âò÷s§Â6VÆc£§s§"ÂæÖW76W3Ôå2¢Ð  ¦FVbö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€¢6¶vS¢¦—f–ÆRå¦—f–ÆRÀ¢Fö7VÖVçE÷&ö÷C¢WG&VRåôVÆVÖVçBÀ¢7G'V7GW&UöÖ¢F–7E·7G"Âç•ÒÀ¢’Óâ6WE·7G%Ó ¢FVÆWFVE÷6V7F–öç2Ò°¢–çB†VçG'•²'6V7F–öâ%Ò¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"¢Ð¢–bæ÷BFVÆWFVE÷6V7F–öç3 ¢&WGW&â6WB‚¢6V7F–öç2ÒFö7VÖVçE÷&ö÷Bç‡F‚‚"âò÷s§6V7E""ÂæÖW76W3Ôå2¢FVÆWFVEö–G3¢6WE·7G%ÒÒ6WB‚¢&WF–æVEö–G3¢6WE·7G%ÒÒ6WB‚¢&VÆF–öç6†—öGG&–'WFRÒâ‚'#¦–B"¢f÷"–æFW‚Â6V7F–öâ–âVçVÖW&FR‡6V7F–öç2“ ¢F&vWBÒFVÆWFVEö–G2–b–æFW‚–âFVÆWFVE÷6V7F–öç2VÇ6R&WF–æVEö–G0¢f÷"&VfW&Væ6R–â6V7F–öâç‡F‚€¢"â÷s¦†VFW%&VfW&Væ6RÂâ÷s¦fö÷FW%&VfW&Væ6R"ÂæÖW76W3Ôå0¢“ ¢&VÆF–öç6†—ö–BÒ&VfW&Væ6RævWB‡&VÆF–öç6†—öGG&–'WFR¢–b&VÆF–öç6†—ö–C ¢F&vWBæFB‡&VÆF–öç6†—ö–B¢&VÖ÷f&ÆUö–G2ÒFVÆWFVEö–G2Ò&WF–æVEö–G0¢–bæ÷B&VÖ÷f&ÆUö–G3 ¢&WGW&â6WB‚¢&VÆF–öç6†—5öæÖRÒ'v÷&Bõ÷&VÇ2öFö7VÖVçBç†ÖÂç&VÇ2 ¢–b&VÆF–öç6†—5öæÖRæ÷B–â6¶vRææÖVÆ—7B‚“ ¢&WGW&â6WB‚¢&VÆF–öç6†—2ÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‡&VÆF–öç6†—5öæÖR’¢&W7VÇBÒ6WB‚¢f÷"&VÆF–öç6†—–â&VÆF–öç6†—3 ¢–b&VÆF–öç6†—ævWB‚$–B"’æ÷B–â&VÖ÷f&ÆUö–G3 ¢6öçF–çVP¢F&vWBÒ&VÆF–öç6†—ævWB‚%F&vWB"Â""¢æ÷&ÖÆ—¦VBÒ÷6—‡F‚ææ÷&×F‚‡÷6—‡F‚æ¦ö–â‚'v÷&B"ÂF&vWB’¢–b&RægVÆÆÖF6‚‡"'v÷&Bòƒó¦†VFW'Æfö÷FW"•ÆBµÂç†ÖÂ"Âæ÷&ÖÆ—¦VB“ ¢&W7VÇBæFB†æ÷&ÖÆ—¦VB¢&WGW&â&W7VÇ@  ¦FVb7G'V7GW&Uö6öçFVçEö–çfVçF÷'’€¢Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•ÐÐ¢’ÓâF–7E·7G"ÂÆ—7E·7G%ÕÓ Ð¢Fö5ö–æFW†W2Ò°Ð¢–æFW€Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'Fö5÷&ævW2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢f÷"–æFW‚–â&ævR†–çB†VçG'•²'7F'E÷&w&‚%Ò’Â–çB†VçG'•²&VæE÷&w&‚%Ò’²Ð¢ÐÐ¢†VF–æuöVçG&–W2Ò°¢VçG'’f÷"VçG'’–â7G'V7GW&UöÖævWB‚&†VF–æw2"ÂµÒ’–bVçG'’ævWB‚&&÷fVB"¢Ð¢6F–öåöVçG&–W2Ò°¢VçG'¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVB"¢æB€¢æ÷B†5ö6F–öåö7F–öç5öÖ‡7G'V7GW&UöÖ¢÷"VçG'’ævWB‚&7F–öâ"’ÓÒ&6öçfW'E÷Fõ÷6W ¢¢Ð¢–FVçF–f–W%÷&WÆ6VÖVçG2Ò°¢VçG'¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVB"’æBVçG'’ævWB‚&7F–öâ"’ÓÒ'&WÆ6Uö–FVçF–f–W" ¢Ð¢Ö–w&FVEö6F–öåö†6†W2Ò°Ð¢VçG'•²'FW‡E÷6†#Sb%ÐÐ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"Ð¢æBVçG'’ævWB‚&Æö6F÷""Â·Ò’ævWB‚&¶–æB"’ÓÒ'F&ÆUö6VÆÅ÷&w&‚ Ð¢ÐÐ¢ÖçVÅöÖ–w&FVEö6F–öåö†6†W2Ò°¢VçG'•²'FW‡E÷6†#Sb%Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVB"¢æBVçG'’ævWB‚&7F–öâ"’ÓÒ&Ö÷fUö6F–öâ ¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"¢Ð¢†5öÖ–w&FVEö6F–öç2Ò&ööÂ€¢Ö–w&FVEö6F–öåö†6†W2ÒÖçVÅöÖ–w&FVEö6F–öåö†6†W0¢¢&W7VÇC¢F–7E·7G"ÂÆ—7E·7G%ÕÒÒ·Ð¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS ¢Fö7VÖVçE÷&ö÷BÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‚'v÷&BöFö7VÖVçBç†ÖÂ"’¢–væ÷&VE÷'G2Òö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€¢6¶vRÂFö7VÖVçE÷&ö÷BÂ7G'V7GW&UöÖ ¢¢f÷"æÖR–â6÷'FVB‡6¶vRææÖVÆ—7B‚’“ ¢–bæ÷B4ôåDTåEõ%BæÖF6‚†æÖR’÷"æÖR–â–væ÷&VE÷'G3 ¢6öçF–çVP¢&ö÷BÒFö7VÖVçE÷&ö÷B–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6RWG&VRæg&ö×7G&–ær€¢6¶vRç&VB†æÖR¢¢Fö5öf–VÆE÷&w&‡2Ò€¢÷Fö5öf–VÆE÷&w&‡2‡&ö÷B’–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6R6WB‚¢¢ÆVv7•÷Fö5öæ6†÷'2Ò€¢öÆVv7•÷Fö5öV×G•öæ6†÷'2‡&ö÷BÂFö5öf–VÆE÷&w&‡2Â7G'V7GW&UöÖ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ ¢VÇ6R6WB‚¢¢&÷fVE÷F–Å÷&w&‡2Ò€¢ö&÷fVE÷F–Å÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ ¢VÇ6R6WB‚¢¢fÇVW2ÒµÐ¢&öG•ö–æFW‚ÒÓ¢f÷"&w&‚–â&ö÷Bç‡F‚‚"âò÷s§"ÂæÖW76W3Ôå2“ ¢F—&V7Eö&öG•÷&w&‚Ò€Ð¢æÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢æB&w&‚ævWG&VçB‚’—2æ÷BæöæPÐ¢æB&w&‚ævWG&VçB‚’çFrÓÒâ‚'s¦&öG’"Ð¢Ð¢–bF—&V7Eö&öG•÷&w&ƒ Ð¢&öG•ö–æFW‚³ÒÐ¢7W'&VçEö–æFW‚Ò&öG•ö–æFW‚–bF—&V7Eö&öG•÷&w&‚VÇ6RæöæPÐ¢–b&w&‚–â&÷fVE÷F–Å÷&w&‡3 ¢6öçF–çVP¢–b&w&‚–âFö5öf–VÆE÷&w&‡2÷"&w&‚–âÆVv7•÷Fö5öæ6†÷'3 ¢6öçF–çVP¢–b€¢7W'&VçEö–æFW‚—2æ÷BæöæP¢æBæ÷BFö5öf–VÆE÷&w&‡0¢æB7W'&VçEö–æFW‚–âFö5ö–æFW†W0¢“ ¢6öçF–çVP¢fÇVRÒ÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚¢fÇVRÒd”TÄEôÔ$´U%õEDU$âç7V"‚""ÂfÇVR¢fÇVRÒöæ÷&ÖÆ—¦UöÖçVÅö–FVçF–f–W"‡fÇVRÂ–FVçF–f–W%÷&WÆ6VÖVçG2¢÷&–v–æÅö6F–öâÒ4D”ôåõEDU$âæÖF6‚‡fÇVR¢–bFW‡E÷6†#Sb‡fÇVR’–âÖçVÅöÖ–w&FVEö6F–öåö†6†W3 ¢fÇVRÒ%µ´ÔõdTEôÔåTÅô4D”ôåÕÒ ¢VÆ–bFW‡E÷6†#Sb‡fÇVR’–âÖ–w&FVEö6F–öåö†6†W2æB÷&–v–æÅö6F–öã ¢fÇVRÒb'¶÷&–v–æÅö6F–öâæw&÷Wƒ—Ò¶÷&–v–æÅö6F–öâæw&÷WƒB—Ò Ð¢VÆ–b††5öÖ–w&FVEö6F–öç2÷"6F–öåöVçG&–W2’æBF—&V7Eö&öG•÷&w&ƒ ¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ð¢vVæW&FVBÒ&RæÖF6‚‡"%åÇ2¢ŽY»çÎŠ‚•Ç2¥²ÞûÈÞ(	N(	5ÕÇ2¢‚â¢’B"ÂfÇVRÐ¢–b7G–ÆW2ÓÒ²$6F–öâ%ÒæBvVæW&FVC Ð¢fÇVRÒb'¶vVæW&FVBæw&÷Wƒ—Ò¶vVæW&FVBæw&÷Wƒ"—Ò Ð¢†VF–æuöVçG'’ÒæW‡B€¢€¢VçG'¢f÷"VçG'’–â†VF–æuöVçG&–W0¢–bFW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"¢’À¢æöæRÀ¢¢–b†VF–æuöVçG'’—2æ÷BæöæS ¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB††VF–æuöVçG'•²&ÆWfVÂ%Ò’’æÖF6‚‡fÇVR¢–bÖF6ƒ ¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥Ð¢VÇ6S ¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2¢7G–ÆUöÖF6‚Ò€¢&RægVÆÆÖF6‚‡"$†VF–ær…³ÓEÒ’"Â7G–ÆW5³Ò’–b7G–ÆW2VÇ6RæöæP¢¢–b7G–ÆUöÖF6ƒ ¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB‡7G–ÆUöÖF6‚æw&÷Wƒ’’’æÖF6‚€¢fÇVP¢¢–bÖF6ƒ ¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥Ð¢–bç’€¢FW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"¢f÷"VçG'’–â6F–öåöVçG&–W0¢“ ¢ÖF6‚Ò4D”ôåõEDU$âæÖF6‚‡fÇVR¢–bÖF6ƒ Ð¢fÇVRÒb'¶ÖF6‚æw&÷Wƒ—Ò¶ÖF6‚æw&÷WƒB—Ò Ð¢VÇ6S Ð¢fÇVRÒ&Rç7V"‡"%âŽY»çÎŠ‚•Ç2µ²ÞûÈÞ(	N(	5ÓõÇ2¢"Â"%Ã"ÂfÇVRÐ¢fÇVW2æVæB‡fÇVRÐ¢&W7VÇE¶æÖUÒÒfÇVW0Ð¢&WGW&â&W7VÇ@Ð Ð Ð¦FVb7G'V7GW&Uö6öçFVçEöf–ævW'&–çB‡Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•Ò’Óâ7G# Ð¢&W7VÇBÒ7G'V7GW&Uö6öçFVçEö–çfVçF÷'’‡F‚Â7G'V7GW&UöÖÐ¢Væ6öFVBÒ§6öâæGV×2‡&W7VÇBÂVç7W&Uö66–“ÔfÇ6RÂ6W&F÷'3Ò‚"Â"Â#¢"’Â6÷'Eö¶W—3ÕG'VRÐ¢&WGW&â†6†Æ–"ç6†#Sb†Væ6öFVBæVæ6öFR‚'WFbÓ‚"’’æ†W†F–vW7B‚Ð