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
    (1, re.compile(r"^\s*ç¬¬\s*[0-9ä¸€äºŒä¸‰å››äº”å…­ä¸ƒå…«ä¹åç™¾]+\s*ç« \s*\S")),
)
CAPTION_PATTERN = re.compile(
    r"^\s*(å›¾|è¡¨)\s*(\d+(?:\.\d+){1,3})\s*[-ï¼â€”â€“]\s*(\d+)\s*(\S.*)$"
)
LOOSE_CAPTION_PATTERN = re.compile(r"^\s*(å›¾|è¡¨)\s*(\S.*)?$")
CAPTION_IDENTIFIER_PATTERN = re.compile(
    r"^\s*(å›¾|è¡¨)\s*([0-9]+(?:[.ï¼-][0-9]+)*)"
    r"(?:\s+|\s*[-ï¼â€”â€“]\s*(?=\D))(\S.*)$"
)
DRAWING_MARK_PATTERN = re.compile(
    r"(?:^|\s)\d+(?:-\d+)+\s*(?:å‰–é¢|æ–­é¢|æˆªé¢|èŠ‚ç‚¹|è¯¦å›¾|å¤§æ ·|ç«‹é¢|å¹³é¢)"
)
ARCHITECTURE_TERMS = ("å»ºç­‘", "å¹³é¢å›¾", "ç«‹é¢å›¾", "å‰–é¢å›¾", "è¯¦å›¾", "å¤§æ ·")
CIVIL_ENGINEERING_TERMS = (
    "åœŸæœ¨",
    "ç»“æ„",
    "é’¢ç»“æ„",
    "æ··å‡åœŸ",
    "æ¢",
    "æŸ±",
    "æ¡æ¶",
    "åŸºç¡€",
    "èŠ‚ç‚¹",
    "æˆªé¢",
    "æ–­é¢",
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
        return "mixed" if "." in identifier or "ï¼" in identifier else "drawing_mark"
    if domain_context != "unknown" and "-" in identifier and not any(
        marker in identifier for marker in (".", "ï¼")
    ):
        return "drawing_mark"
    if any(marker in identifier for marker in (".", "ï¼")):
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
            locatoÛ»êÚ$z{-®éÜj×ÓÕ²''G2%ÒæVæB†VÆVÖVçBçFW‡B÷"""Ğ¢–bç’†—FVÒævWB‚&—5÷Fö2"’æB—FVÒævWB‚'&W7VÇB"’f÷"—FVÒ–â7F6²“ Ğ¢&W7VÇBæFB‡&w&‚Ğ¢&WGW&â&W7VÇ@Ğ Ğ Ğ¦FVböÆVv7•÷Fö5öV×G•öæ6†÷'2€Ğ¢&ö÷C¢WG&VRåôVÆVÖVçBÀĞ¢Fö5öf–VÆE÷&w&‡3¢6WE´ç•ÒÀĞ¢7G'V7GW&UöÖ¢F–7E·7G"Âç•ÒÀĞ¢’Óâ6WE´ç•Ó Ğ¢""%&V6övæ—¦RV×G’æ6†÷'2ÆVgB'’&RÓã27FF–2ÕDô2Ö–w&F–öââ"" Ğ¢–b7G'V7GW&UöÖævWB‚'66†VÖ÷fW'6–öâ"’–â²#ã2"Â#ãB'Ó Ğ¢&WGW&â6WB‚Ğ¢–bæ÷BFö5öf–VÆE÷&w&‡3 Ğ¢&WGW&â6WB‚Ğ¢&÷fVEöÆVæwF‡2Ò°Ğ¢–çB†VçG'•²&VæE÷&w&‚%Ò’Ò–çB†VçG'•²'7F'E÷&w&‚%ÒĞ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'Fö5÷&ævW2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"Ğ¢ĞĞ¢&VÖ–æ–ærÒÖ‚†&÷fVEöÆVæwF‡2ÂFVfVÇCÓĞ¢–b&VÖ–æ–ærÃÒ Ğ¢&WGW&â6WB‚Ğ¢&öG’Ò&ö÷Bæf–æB‡â‚'s¦&öG’"’Ğ¢–b&öG’—2æöæS Ğ¢&WGW&â6WB‚Ğ¢6†–ÆG&VâÒÆ—7B†&öG’Ğ¢f–VÆE÷÷6—F–öç2Ò°Ğ¢–æFW‚f÷"–æFW‚Â6†–ÆB–âVçVÖW&FR†6†–ÆG&Vâ’–b6†–ÆB–âFö5öf–VÆE÷&w&‡0Ğ¢ĞĞ¢–bæ÷Bf–VÆE÷÷6—F–öç3 Ğ¢&WGW&â6WB‚Ğ¢&W7VÇBÒ6WB‚Ğ¢f÷"6†–ÆB–â6†–ÆG&Vå¶Ö‚†f–VÆE÷÷6—F–öç2’²¥Ó Ğ¢–b&VÖ–æ–ærÃÒ÷"6†–ÆBçFrÒâ‚'s§"“ Ğ¢'&V°Ğ¢fÇVRÒd”TÄEôÔ$´U%õEDU$âç7V"€Ğ¢""Â÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2†6†–ÆBĞ¢Ğ¢–bfÇVS Ğ¢'&V°Ğ¢&W7VÇBæFB†6†–ÆBĞ¢&VÖ–æ–ærÓÒĞ¢&WGW&â&W7VÇ@Ğ Ğ Ğ¦FVb÷v–æF–öåö&÷VæF'•÷&w&‡2€Ğ¢&ö÷C¢WG&VRåôVÆVÖVçBÂ7G'V7GW&UöÖ¢F–7E·7G"Âç•ĞĞ¢’Óâ6WE´ç•Ó Ğ¢v–æF–öâÒ7G'V7GW&UöÖævWB‚'v–æF–öå÷6V7F–öç2"Â·ÒĞ¢–b7G'V7GW&UöÖævWB‚'66†VÖ÷fW'6–öâ"’Ò#ãB# Ğ¢&WGW&â6WB‚Ğ¢&W7VÇBÒ6WB‚Ğ¢†5öÆæG66U÷F&ÆRÒç’€Ğ¢VçG'’ævWB‚&&÷fVB"Ğ¢æBVçG'’ævWB‚'f—7VÂ"Â·Ò’ævWB‚&&÷fVB"Ğ¢æBVçG'’ævWB‚'f—7VÂ"Â·Ò’ævWB‚&÷&–VçFF–öâ"’ÓÒ&ÆæG66R Ğ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'F&ÆW2"ÂµÒĞ¢Ğ¢–b†5öÆæG66U÷F&ÆS Ğ¢&W7VÇBçWFFR€Ğ¢&w&€Ğ¢f÷"&w&‚–â&ö÷Bç‡F‚‚"÷s¦Fö7VÖVçB÷s¦&öG’÷s§"ÂæÖW76W3Ôå2Ğ¢–b&w&‚æf–æB‚"â÷s§"÷s§6V7E""ÂæÖW76W3Ôå2’—2æ÷BæöæPĞ¢æBæ÷B÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚Ğ¢Ğ¢–bæ÷Bv–æF–öâævWB‚&&÷fVB"“ Ğ¢&WGW&â&W7VÇ@Ğ¢&öG•öÆö6F÷"Òv–æF–öâævWB‚&&öG•÷7F'B"Â·Ò¢W‡V7FVBÒ&öG•öÆö6F÷"ævWB‚'FW‡E÷6†#Sb"¢–bæ÷BW‡V7FVC ¢&WGW&â&W7VÇ@¢W‡V7FVEö†6†W2Ò¶W‡V7FVGĞ¢†VF–ærÒæW‡B€¢€¢VçG'¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&†VF–æw2"ÂµÒ¢–bVçG'’ævWB‚&&÷fVB"¢æBVçG'’ævWB‚&Æö6F÷""Â·Ò’ævWB‚'FW‡E÷6†#Sb"’ÓÒW‡V7FV@¢’À¢æöæRÀ¢¢–b†VF–æræB†VF–ærævWB‚&æ÷&ÖÆ—¦VE÷FW‡E÷6†#Sb"“ ¢W‡V7FVEö†6†W2æFB††VF–æu²&æ÷&ÖÆ—¦VE÷FW‡E÷6†#Sb%Ò¢&w&‡2Ò&ö÷Bç‡F‚‚"÷s¦Fö7VÖVçB÷s¦&öG’÷s§"ÂæÖW76W3Ôå2¢ÖF6†W2Ò°¢&w&€¢f÷"&w&‚–â&w&‡0¢–bFW‡E÷6†#Sb…÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚’¢–âW‡V7FVEö†6†W0¢ĞĞ¢–bÆVâ†ÖF6†W2’Ò Ğ¢&WGW&â&W7VÇ@Ğ¢&Wf–÷W2ÒÖF6†W5³ÒævWG&Wf–÷W2‚Ğ¢–b€¢&Wf–÷W2—2æ÷BæöæP¢æB&Wf–÷W2çFrÓÒâ‚'s§"Ğ¢æB&Wf–÷W2æf–æB‚"â÷s§"÷s§6V7E""ÂæÖW76W3Ôå2’—2æ÷BæöæPĞ¢æBæ÷B÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&Wf–÷W2Ğ¢“ ¢&W7VÇBæFB‡&Wf–÷W2¢&WGW&â&W7VÇ@ Ğ Ğ¦FVbö&÷fVE÷F–Å÷&w&‡2€Ğ¢&ö÷C¢WG&VRåôVÆVÖVçBÂ7G'V7GW&UöÖ¢F–7E·7G"Âç•ĞĞ¢’Óâ6WE´ç•Ó Ğ¢&÷fVBÒ°Ğ¢VçG'Ğ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"Ğ¢ĞĞ¢–bæ÷B&÷fVC Ğ¢&WGW&â6WB‚Ğ¢&öG’Ò&ö÷Bæf–æB‡â‚'s¦&öG’"’Ğ¢–b&öG’—2æöæS Ğ¢&WGW&â6WB‚Ğ¢F—&V7E÷&w&‡2Ò¶6†–ÆBf÷"6†–ÆB–â&öG’–b6†–ÆBçFrÓÒâ‚'s§"•ĞĞ¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2ÒµĞĞ¢6†–ÆG&VâÒÆ—7B†&öG’Ğ¢f÷"VçG'’–â&÷fVC Ğ¢W‡V7FVEö†6‚ÒVçG'’ævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"Ğ¢&÷VæF'•ö—5öV×G’ÒW‡V7FVEö†6‚ÓÒFW‡E÷6†#Sb‚""Ğ¢ÖF6†–ærÒ°Ğ¢&w&€Ğ¢f÷"&w&‚–âF—&V7E÷&w&‡0Ğ¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb…÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚’Ğ¢ÓÒW‡V7FVEö†6€Ğ¢ĞĞ¢–bÆVâ†ÖF6†–ær’ÓÒ Ğ¢&÷VæF'•÷&w&‚ÒÖF6†–æu³ĞĞ¢VÇ6S Ğ¢&÷VæF'’Ò–çB†VçG'•²'&Wf–÷W5ö&÷VæF'•÷&w&‚%ÒĞ¢–bæ÷BÃÒ&÷VæF'’ÂÆVâ†F—&V7E÷&w&‡2“ Ğ¢6öçF–çVPĞ¢6æF–FFRÒF—&V7E÷&w&‡5¶&÷VæF'•ĞĞ¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb€Ğ¢÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2†6æF–FFRĞ¢’ÒW‡V7FVEö†6ƒ Ğ¢6öçF–çVPĞ¢&÷VæF'•÷&w&‚Ò6æF–FFPĞ¢–b&÷VæF'•ö—5öV×G’æBç’€Ğ¢—FVÒævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"’ÒFW‡E÷6†#Sb‚""Ğ¢f÷"—FVÒ–â&÷fV@Ğ¢“ Ğ¢6öçF–çVPĞ¢&÷VæF'•÷÷6—F–öâÒ6†–ÆG&Vâæ–æFW‚†&÷VæF'•÷&w&‚Ğ¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2æVæB€Ğ¢&÷VæF'•÷÷6—F–öâ–b&÷VæF'•ö—5öV×G’VÇ6R&÷VæF'•÷÷6—F–öâ²Ğ¢Ğ¢–bæ÷Bf—'7E÷&VÖ÷fVE÷÷6—F–öç3 Ğ¢&WGW&â6WB‚Ğ¢f—'7E÷&VÖ÷fVBÒÖ–â†f—'7E÷&VÖ÷fVE÷÷6—F–öç2Ğ¢&WGW&â°Ğ¢&w&€Ğ¢f÷"6†–ÆB–â6†–ÆG&Vå¶f—'7E÷&VÖ÷fVC¥ĞĞ¢f÷"&w&‚–â6†–ÆBç‡F‚‚"âò÷s§Â6VÆc£§s§"ÂæÖW76W3Ôå2Ğ¢ĞĞ Ğ Ğ¦FVbö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€Ğ¢6¶vS¢¦—f–ÆRå¦—f–ÆRÀĞ¢Fö7VÖVçE÷&ö÷C¢WG&VRåôVÆVÖVçBÀĞ¢7G'V7GW&UöÖ¢F–7E·7G"Âç•ÒÀĞ¢’Óâ6WE·7G%Ó Ğ¢FVÆWFVE÷6V7F–öç2Ò°Ğ¢–çB†VçG'•²'6V7F–öâ%ÒĞ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"Ğ¢ĞĞ¢–bæ÷BFVÆWFVE÷6V7F–öç3 Ğ¢&WGW&â6WB‚Ğ¢6V7F–öç2ÒFö7VÖVçE÷&ö÷Bç‡F‚‚"âò÷s§6V7E""ÂæÖW76W3Ôå2Ğ¢FVÆWFVEö–G3¢6WE·7G%ÒÒ6WB‚Ğ¢&WF–æVEö–G3¢6WE·7G%ÒÒ6WB‚Ğ¢&VÆF–öç6†—öGG&–'WFRÒâ‚'#¦–B"Ğ¢f÷"–æFW‚Â6V7F–öâ–âVçVÖW&FR‡6V7F–öç2“ Ğ¢F&vWBÒFVÆWFVEö–G2–b–æFW‚–âFVÆWFVE÷6V7F–öç2VÇ6R&WF–æVEö–G0Ğ¢f÷"&VfW&Væ6R–â6V7F–öâç‡F‚€Ğ¢"â÷s¦†VFW%&VfW&Væ6RÂâ÷s¦fö÷FW%&VfW&Væ6R"ÂæÖW76W3Ôå0Ğ¢“ Ğ¢&VÆF–öç6†—ö–BÒ&VfW&Væ6RævWB‡&VÆF–öç6†—öGG&–'WFRĞ¢–b&VÆF–öç6†—ö–C Ğ¢F&vWBæFB‡&VÆF–öç6†—ö–BĞ¢&VÖ÷f&ÆUö–G2ÒFVÆWFVEö–G2Ò&WF–æVEö–G0Ğ¢–bæ÷B&VÖ÷f&ÆUö–G3 Ğ¢&WGW&â6WB‚Ğ¢&VÆF–öç6†—5öæÖRÒ'v÷&Bõ÷&VÇ2öFö7VÖVçBç†ÖÂç&VÇ2 Ğ¢–b&VÆF–öç6†—5öæÖRæ÷B–â6¶vRææÖVÆ—7B‚“ Ğ¢&WGW&â6WB‚Ğ¢&VÆF–öç6†—2ÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‡&VÆF–öç6†—5öæÖR’Ğ¢&W7VÇBÒ6WB‚Ğ¢f÷"&VÆF–öç6†—–â&VÆF–öç6†—3 Ğ¢–b&VÆF–öç6†—ævWB‚$–B"’æ÷B–â&VÖ÷f&ÆUö–G3 Ğ¢6öçF–çVPĞ¢F&vWBÒ&VÆF–öç6†—ævWB‚%F&vWB"Â""Ğ¢æ÷&ÖÆ—¦VBÒ÷6—‡F‚ææ÷&×F‚‡÷6—‡F‚æ¦ö–â‚'v÷&B"ÂF&vWB’¢–b&RægVÆÆÖF6‚‡"'v÷&Bòƒó¦†VFW'Æfö÷FW"•ÆBµÂç†ÖÂ"Âæ÷&ÖÆ—¦VB“ ¢&W7VÇBæFB†æ÷&ÖÆ—¦VB¢&WGW&â&W7VÇ@ Ğ Ğ¦FVb7G'V7GW&Uö6öçFVçEö–çfVçF÷'’€Ğ¢Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•ĞĞ¢’ÓâF–7E·7G"ÂÆ—7E·7G%ÕÓ Ğ¢Fö5ö–æFW†W2Ò°Ğ¢–æFW€Ğ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'Fö5÷&ævW2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"Ğ¢f÷"–æFW‚–â&ævR†–çB†VçG'•²'7F'E÷&w&‚%Ò’Â–çB†VçG'•²&VæE÷&w&‚%Ò’²Ğ¢ĞĞ¢†VF–æuöVçG&–W2Ò°Ğ¢VçG'’f÷"VçG'’–â7G'V7GW&UöÖævWB‚&†VF–æw2"ÂµÒ’–bVçG'’ævWB‚&&÷fVB"Ğ¢ĞĞ¢6F–öåöVçG&–W2Ò°Ğ¢VçG'Ğ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"Ğ¢æB€Ğ¢æ÷B†5ö6F–öåö7F–öç5öÖ‡7G'V7GW&UöÖĞ¢÷"VçG'’ævWB‚&7F–öâ"’ÓÒ&6öçfW'E÷Fõ÷6W Ğ¢Ğ¢ĞĞ¢–FVçF–f–W%÷&WÆ6VÖVçG2Ò°Ğ¢VçG'Ğ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"’æBVçG'’ævWB‚&7F–öâ"’ÓÒ'&WÆ6Uö–FVçF–f–W" Ğ¢ĞĞ¢Ö–w&FVEö6F–öåö†6†W2Ò°Ğ¢VçG'•²'FW‡E÷6†#Sb%ĞĞ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"Ğ¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"Ğ¢æBVçG'’ævWB‚&Æö6F÷""Â·Ò’ævWB‚&¶–æB"’ÓÒ'F&ÆUö6VÆÅ÷&w&‚ Ğ¢ĞĞ¢ÖçVÅöÖ–w&FVEö6F–öåö†6†W2Ò°Ğ¢VçG'•²'FW‡E÷6†#Sb%ĞĞ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒĞ¢–bVçG'’ævWB‚&&÷fVB"Ğ¢æBVçG'’ævWB‚&7F–öâ"’ÓÒ&Ö÷fUö6F–öâ Ğ¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"Ğ¢ĞĞ¢†5öÖ–w&FVEö6F–öç2Ò&ööÂ€Ğ¢Ö–w&FVEö6F–öåö†6†W2ÒÖçVÅöÖ–w&FVEö6F–öåö†6†W0Ğ¢Ğ¢&W7VÇC¢F–7E·7G"ÂÆ—7E·7G%ÕÒÒ·ĞĞ¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ğ¢Fö7VÖVçE÷&ö÷BÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‚'v÷&BöFö7VÖVçBç†ÖÂ"’Ğ¢–væ÷&VE÷'G2Òö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€Ğ¢6¶vRÂFö7VÖVçE÷&ö÷BÂ7G'V7GW&UöÖ Ğ¢Ğ¢f÷"æÖR–â6÷'FVB‡6¶vRææÖVÆ—7B‚’“ Ğ¢–bæ÷B4ôåDTåEõ%BæÖF6‚†æÖR’÷"æÖR–â–væ÷&VE÷'G3 Ğ¢6öçF–çVPĞ¢&ö÷BÒFö7VÖVçE÷&ö÷B–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6RWG&VRæg&ö×7G&–ær€Ğ¢6¶vRç&VB†æÖRĞ¢Ğ¢Fö5öf–VÆE÷&w&‡2Ò€Ğ¢÷Fö5öf–VÆE÷&w&‡2‡&ö÷B’–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6R6WB‚Ğ¢Ğ¢ÆVv7•÷Fö5öæ6†÷'2Ò€Ğ¢öÆVv7•÷Fö5öV×G•öæ6†÷'2‡&ö÷BÂFö5öf–VÆE÷&w&‡2Â7G'V7GW&UöÖĞ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ğ¢VÇ6R6WB‚Ğ¢Ğ¢&÷fVE÷F–Å÷&w&‡2Ò€Ğ¢ö&÷fVE÷F–Å÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖĞ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ğ¢VÇ6R6WB‚Ğ¢Ğ¢v–æF–öåö&÷VæF&–W2Ò€Ğ¢÷v–æF–öåö&÷VæF'•÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖĞ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ğ¢VÇ6R6WB‚Ğ¢Ğ¢fÇVW2ÒµĞĞ¢&öG•ö–æFW‚ÒÓĞ¢f÷"&w&‚–â&ö÷Bç‡F‚‚"âò÷s§"ÂæÖW76W3Ôå2“ Ğ¢F—&V7Eö&öG•÷&w&‚Ò€Ğ¢æÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ğ¢æB&w&‚ævWG&VçB‚’—2æ÷BæöæPĞ¢æB&w&‚ævWG&VçB‚’çFrÓÒâ‚'s¦&öG’"Ğ¢Ğ¢–bF—&V7Eö&öG•÷&w&ƒ Ğ¢&öG•ö–æFW‚³ÒĞ¢7W'&VçEö–æFW‚Ò&öG•ö–æFW‚–bF—&V7Eö&öG•÷&w&‚VÇ6RæöæPĞ¢–b&w&‚–â&÷fVE÷F–Å÷&w&‡2÷"&w&‚–âv–æF–öåö&÷VæF&–W3 Ğ¢6öçF–çVPĞ¢–b&w&‚–âFö5öf–VÆE÷&w&‡2÷"&w&‚–âÆVv7•÷Fö5öæ6†÷'3 Ğ¢6öçF–çVPĞ¢–b€Ğ¢7W'&VçEö–æFW‚—2æ÷BæöæPĞ¢æBæ÷BFö5öf–VÆE÷&w&‡0Ğ¢æB7W'&VçEö–æFW‚–âFö5ö–æFW†W0Ğ¢“ Ğ¢6öçF–çVPĞ¢fÇVRÒ÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚Ğ¢fÇVRÒd”TÄEôÔ$´U%õEDU$âç7V"‚""ÂfÇVRĞ¢fÇVRÒöæ÷&ÖÆ—¦UöÖçVÅö–FVçF–f–W"‡fÇVRÂ–FVçF–f–W%÷&WÆ6VÖVçG2Ğ¢÷&–v–æÅö6F–öâÒ4D”ôåõEDU$âæÖF6‚‡fÇVRĞ¢–bFW‡E÷6†#Sb‡fÇVR’–âÖçVÅöÖ–w&FVEö6F–öåö†6†W3 Ğ¢fÇVRÒ%µ´ÔõdTEôÔåTÅô4D”ôåÕÒ Ğ¢VÆ–bFW‡E÷6†#Sb‡fÇVR’–âÖ–w&FVEö6F–öåö†6†W2æB÷&–v–æÅö6F–öã Ğ¢fÇVRÒb'¶÷&–v–æÅö6F–öâæw&÷Wƒ—Ò¶÷&–v–æÅö6F–öâæw&÷WƒB—Ò Ğ¢VÆ–b††5öÖ–w&FVEö6F–öç2÷"6F–öåöVçG&–W2’æBF—&V7Eö&öG•÷&w&ƒ Ğ¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ğ¢vVæW&FVBÒ&RæÖF6‚‡"%åÇ2¢Y»çÎŠ‚•Ç2¥²ŞûÈŞ(	N(	5ÕÇ2¢‚â¢’B"ÂfÇVRĞ¢–b7G–ÆW2ÓÒ²$6F–öâ%ÒæBvVæW&FVC Ğ¢fÇVRÒb'¶vVæW&FVBæw&÷Wƒ—Ò¶vVæW&FVBæw&÷Wƒ"—Ò Ğ¢†VF–æuöVçG'’ÒæW‡B€Ğ¢€Ğ¢VçG'Ğ¢f÷"VçG'’–â†VF–æuöVçG&–W0Ğ¢–bFW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"Ğ¢’ÀĞ¢æöæRÀĞ¢Ğ¢–b†VF–æuöVçG'’—2æ÷BæöæS Ğ¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB††VF–æuöVçG'•²&ÆWfVÂ%Ò’’æÖF6‚‡fÇVRĞ¢–bÖF6ƒ Ğ¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥ĞĞ¢VÇ6S Ğ¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ğ¢7G–ÆUöÖF6‚Ò€Ğ¢&RægVÆÆÖF6‚‡"$†VF–ær…³ÓEÒ’"Â7G–ÆW5³Ò’–b7G–ÆW2VÇ6RæöæPĞ¢Ğ¢–b7G–ÆUöÖF6ƒ Ğ¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB‡7G–ÆUöÖF6‚æw&÷Wƒ’’’æÖF6‚€Ğ¢fÇVPĞ¢Ğ¢–bÖF6ƒ Ğ¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥ĞĞ¢–bç’€Ğ¢FW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"Ğ¢f÷"VçG'’–â6F–öåöVçG&–W0Ğ¢“ Ğ¢ÖF6‚Ò4D”ôåõEDU$âæÖF6‚‡fÇVRĞ¢–bÖF6ƒ Ğ¢fÇVRÒb'¶ÖF6‚æw&÷Wƒ—Ò¶ÖF6‚æw&÷WƒB—Ò Ğ¢VÇ6S Ğ¢fÇVRÒ&Rç7V"‡"%âY»çÎŠ‚•Ç2µ²ŞûÈŞ(	N(	5ÓõÇ2¢"Â"%Ã"ÂfÇVRĞ¢fÇVW2æVæB‡fÇVRĞ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"÷"ç’‡fÇVW2“ ¢&W7VÇE¶æÖUÒÒfÇVW0¢&÷fVEöFW&—fVEöfö÷FW%ööæÇ’Ò&ööÂ€¢7G'V7GW&UöÖævWB‚'v–æF–öå÷6V7F–öç2"Â·Ò’ævWB‚&&÷fVB"¢æBç’€¢VçG'’ævWB‚&&÷fVEöFVÆWFR"¢æBVçG'’ævWB‚&Wf–FVæ6R"Â·Ò’ævWB‚&&÷fVEöFW&—fVEöfö÷FW%ööæÇ’"¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒ¢¢¢6æöæ–6Ã¢F–7E·7G"ÂÆ—7E·7G%ÕÒÒ·Ğ¢†VFW'3¢Æ—7E·7G%ÒÒµĞ¢fö÷FW'3¢Æ—7E·7G%ÒÒµĞ¢f÷"æÖRÂfÇVW2–â&W7VÇBæ—FV×2‚“ ¢–b&RægVÆÆÖF6‚‡"'v÷&Bö†VFW%ÆBµÂç†ÖÂ"ÂæÖR“ ¢†VFW'2æW‡FVæB‡fÇVW2¢VÆ–b&RægVÆÆÖF6‚‡"'v÷&Böfö÷FW%ÆBµÂç†ÖÂ"ÂæÖR“ ¢–bæ÷B&÷fVEöFW&—fVEöfö÷FW%ööæÇ“ ¢fö÷FW'2æW‡FVæB‡fÇVW2¢VÇ6S ¢6æöæ–6Å¶æÖUÒÒfÇVW0¢6æöæ–6Å²'v÷&Bõö†VFW'5ö'•ö6öçFVçB%ÒÒ6÷'FVB††VFW'2¢6æöæ–6Å²'v÷&Bõöfö÷FW'5ö'•ö6öçFVçB%ÒÒ6÷'FVB†fö÷FW'2¢&WGW&â6æöæ–6À Ğ Ğ¦FVb7G'V7GW&Uö6öçFVçEöf–ævW'&–çB‡Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•Ò’Óâ7G# Ğ¢&W7VÇBÒ7G'V7GW&Uö6öçFVçEö–çfVçF÷'’‡F‚Â7G'V7GW&UöÖĞ¢Væ6öFVBÒ§6öâæGV×2‡&W7VÇBÂVç7W&Uö66–“ÔfÇ6RÂ6W&F÷'3Ò‚"Â"Â#¢"’Â6÷'Eö¶W—3ÕG'VRĞ¢&WGW&â†6†Æ–"ç6†#Sb†Væ6öFVBæVæ6öFR‚'WFbÓ‚"’’æ†W†F–vW7B‚Ğ