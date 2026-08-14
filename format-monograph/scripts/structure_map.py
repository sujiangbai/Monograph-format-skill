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
from docx.image.image import Image as DocxImage
from docx.image.exceptions import UnrecognizedImageError
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from lxml import etree

from _common import (
    CONTENT_PART,
    FIELD_MARKER_PATTERN,
    NS,
    STRUCTURAL_INDENT_ATTRIBUTES,
    FormatMonographError,
    _numbering_level_for_style,
    _heading_prefix_pattern,
    _paragraph_text_without_field_results,
    _unique_row_cells,
    apply_table_properties,
    apply_style_properties,
    clear_controlled_direct_format,
    content_fingerprint,
    ensure_paragraph_style,
    normalize_structural_paragraph,
)
from docx_pagination import apply_pagination_sections, section_index_for_paragraph


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
BOOK_TITLE_STYLE = "Monograph Book Title"
TOC_HEADING_STYLE = "Monograph TOC Heading"
BLOCK_SPACER_STYLE = "Monograph Figure Table Spacer"
DEFAULT_BOOK_TITLE_FORMAT = {
    "font_name_east_asia": "é»‘ä½“",
    "font_name_ascii": "Times New Roman",
    "font_name_complex_script": "Times New Roman",
    "font_size_pt": 22,
    "bold": True,
    "alignment": "center",
    "first_line_indent_chars": 0,
    "left_indent_pt": 0,
    "right_indent_pt": 0,
    "line_spacing_rule": "at_least",
    "line_spacing_pt": 33,
    "space_before_pt": 0,
    "space_after_pt": 0,
}
EMU_PER_INCH = 914400
EMU_PER_MM = 36000
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
STRUCTURE_MAP_VERSIONS = {"1.0", "1.1", "1.2", "1.3", "1.4"}
STRUCTURE_MAP_VERSIONS.add("1.5")
SEMANTIC_STRUCTURE_MAP_VERSIONS = {"1.1", "1.2", "1.3", "1.4", "1.5"}
CAPTION_ACTIONS = {
    "preserve",
    "style_only",
    "replace_identifier",
    "convert_to_seq",
    "move_caption",
}

ROLE_ALIASES = {
    "body_text": "body",
    "heading_1": "chapter_title",
    "heading1": "chapter_title",
    "heading_2": "level_2_section",
    "heading2": "level_2_section",
    "heading_3": "level_3_section",
    "heading3": "level_3_section",
    "heading_4": "level_4_section",
    "heading4": "level_4_section",
    "appendix_heading": "chapter_title",
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
    "chapter_title": "Heading 1",
    "level_2_section": "Heading 2",
    "level_3_section": "Heading 3",
    "level_4_section": "Heading 4",
    "appendix_heading": "Heading 1",
    "figure_caption": "Caption",
    "figure_caption_unnumbered": "Caption",
    "figure_panel_label": "Caption",
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

APPENDIX_PATTERN = re.compile(
    r"^\s*(?:é™„\s*å½•|APPENDIX)\s*[A-Z0-9ä¸€äºŒä¸‰å››äº”å…­ä¸ƒå…«ä¹åç™¾]*",
    re.IGNORECASE,
)


def text_sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _heading_authored_text_hash(value: str, level: int) -> str:
    match = _heading_prefix_pattern(level).match(value)
    return text_sha256(value[match.end() :] if match else value)


def has_semantic_structure_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in SEMANTIC_STRUCTURE_MAP_VERSIONS


def has_caption_actions_map(structure_map: dict[str, Any]) -> bool:
    return structure_map.get("schema_version") in {"1.2", "1.3", "1.4", "1.5"}


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
    try:
        paragraph = resolve_paragraph_locator(document, locator)
    except FormatMonographError:
        normalized = entry.get("normalized_text_sha256")
        candidates = (
            [
                candidate
                for candidate in document.paragraphs
                if text_sha256(candidate.text) == normalized
            ]
            if normalized and locator.get("kind") == "body_paragraph"
            else []
        )
        if len(candidates) != 1:
            raise
        paragraph = candidates[0]
        cache = getattr(document, "_format_monograph_locator_cache", {})
        cache[key] = paragraph
        setattr(document, "_format_monograph_locator_cache", cache)
    expected = entry.get("text_sha256")
    actual = text_sha256(paragraph.text)
    if expected and actual != expected and actual != entry.get(
        "normalized_text_sha256"
    ):
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
    verified = getattr(ç_{òÚ$z{-®éÜj×‡V7FVBÒ&öG•öÆö6F÷"ævWB‚'FW‡E÷6†#Sb"Ð¢–bæ÷BW‡V7FVC Ð¢&WGW&â&W7VÇ@Ð¢W‡V7FVEö†6†W2Ò¶W‡V7FVGÐÐ¢†VF–ærÒæW‡B€Ð¢€Ð¢VçG'Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&†VF–æw2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢æBVçG'’ævWB‚&Æö6F÷""Â·Ò’ævWB‚'FW‡E÷6†#Sb"’ÓÒW‡V7FV@Ð¢’ÀÐ¢æöæRÀÐ¢Ð¢–b†VF–æræB†VF–ærævWB‚&æ÷&ÖÆ—¦VE÷FW‡E÷6†#Sb"“ Ð¢W‡V7FVEö†6†W2æFB††VF–æu²&æ÷&ÖÆ—¦VE÷FW‡E÷6†#Sb%ÒÐ¢&w&‡2Ò&ö÷Bç‡F‚‚"÷s¦Fö7VÖVçB÷s¦&öG’÷s§"ÂæÖW76W3Ôå2Ð¢ÖF6†W2Ò°Ð¢&w&€Ð¢f÷"&w&‚–â&w&‡0Ð¢–bFW‡E÷6†#Sb…÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚’Ð¢–âW‡V7FVEö†6†W0Ð¢ÐÐ¢–bÆVâ†ÖF6†W2’Ò Ð¢&WGW&â&W7VÇ@Ð¢&Wf–÷W2ÒÖF6†W5³ÒævWG&Wf–÷W2‚Ð¢–b€Ð¢&Wf–÷W2—2æ÷BæöæPÐ¢æB&Wf–÷W2çFrÓÒâ‚'s§"Ð¢æB&Wf–÷W2æf–æB‚"â÷s§"÷s§6V7E""ÂæÖW76W3Ôå2’—2æ÷BæöæPÐ¢æBæ÷B÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&Wf–÷W2Ð¢“ Ð¢&W7VÇBæFB‡&Wf–÷W2Ð¢&WGW&â&W7VÇ@Ð Ð Ð¦FVbö&÷fVE÷F–Å÷&w&‡2€Ð¢&ö÷C¢WG&VRåôVÆVÖVçBÂ7G'V7GW&UöÖ¢F–7E·7G"Âç•ÐÐ¢’Óâ6WE´ç•Ó Ð¢&÷fVBÒ°Ð¢VçG'Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"Ð¢ÐÐ¢–bæ÷B&÷fVC Ð¢&WGW&â6WB‚Ð¢&öG’Ò&ö÷Bæf–æB‡â‚'s¦&öG’"’Ð¢–b&öG’—2æöæS Ð¢&WGW&â6WB‚Ð¢F—&V7E÷&w&‡2Ò¶6†–ÆBf÷"6†–ÆB–â&öG’–b6†–ÆBçFrÓÒâ‚'s§"•ÐÐ¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2ÒµÐÐ¢6†–ÆG&VâÒÆ—7B†&öG’Ð¢f÷"VçG'’–â&÷fVC Ð¢W‡V7FVEö†6‚ÒVçG'’ævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"Ð¢&÷VæF'•ö—5öV×G’ÒW‡V7FVEö†6‚ÓÒFW‡E÷6†#Sb‚""Ð¢ÖF6†–ærÒ°Ð¢&w&€Ð¢f÷"&w&‚–âF—&V7E÷&w&‡0Ð¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb…÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚’Ð¢ÓÒW‡V7FVEö†6€Ð¢ÐÐ¢–bÆVâ†ÖF6†–ær’ÓÒ Ð¢&÷VæF'•÷&w&‚ÒÖF6†–æu³ÐÐ¢VÇ6S Ð¢&÷VæF'’Ò–çB†VçG'•²'&Wf–÷W5ö&÷VæF'•÷&w&‚%ÒÐ¢–bæ÷BÃÒ&÷VæF'’ÂÆVâ†F—&V7E÷&w&‡2“ Ð¢6öçF–çVPÐ¢6æF–FFRÒF—&V7E÷&w&‡5¶&÷VæF'•ÐÐ¢–bW‡V7FVEö†6‚æBFW‡E÷6†#Sb€Ð¢÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2†6æF–FFRÐ¢’ÒW‡V7FVEö†6ƒ Ð¢6öçF–çVPÐ¢&÷VæF'•÷&w&‚Ò6æF–FFPÐ¢–b&÷VæF'•ö—5öV×G’æBç’€Ð¢—FVÒævWB‚'&Wf–÷W5ö&÷VæF'•÷6†#Sb"’ÒFW‡E÷6†#Sb‚""Ð¢f÷"—FVÒ–â&÷fV@Ð¢“ Ð¢6öçF–çVPÐ¢&÷VæF'•÷÷6—F–öâÒ6†–ÆG&Vâæ–æFW‚†&÷VæF'•÷&w&‚Ð¢f—'7E÷&VÖ÷fVE÷÷6—F–öç2æVæB€Ð¢&÷VæF'•÷÷6—F–öâ–b&÷VæF'•ö—5öV×G’VÇ6R&÷VæF'•÷÷6—F–öâ²Ð¢Ð¢–bæ÷Bf—'7E÷&VÖ÷fVE÷÷6—F–öç3 Ð¢&WGW&â6WB‚Ð¢f—'7E÷&VÖ÷fVBÒÖ–â†f—'7E÷&VÖ÷fVE÷÷6—F–öç2Ð¢&WGW&â°Ð¢&w&€Ð¢f÷"6†–ÆB–â6†–ÆG&Vå¶f—'7E÷&VÖ÷fVC¥ÐÐ¢f÷"&w&‚–â6†–ÆBç‡F‚‚"âò÷s§Â6VÆc£§s§"ÂæÖW76W3Ôå2Ð¢ÐÐ Ð Ð¦FVbö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€Ð¢6¶vS¢¦—f–ÆRå¦—f–ÆRÀÐ¢Fö7VÖVçE÷&ö÷C¢WG&VRåôVÆVÖVçBÀÐ¢7G'V7GW&UöÖ¢F–7E·7G"Âç•ÒÀÐ¢’Óâ6WE·7G%Ó Ð¢FVÆWFVE÷6V7F–öç2Ò°Ð¢–çB†VçG'•²'6V7F–öâ%ÒÐ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVEöFVÆWFR"Ð¢ÐÐ¢–bæ÷BFVÆWFVE÷6V7F–öç3 Ð¢&WGW&â6WB‚Ð¢6V7F–öç2ÒFö7VÖVçE÷&ö÷Bç‡F‚‚"âò÷s§6V7E""ÂæÖW76W3Ôå2Ð¢FVÆWFVEö–G3¢6WE·7G%ÒÒ6WB‚Ð¢&WF–æVEö–G3¢6WE·7G%ÒÒ6WB‚Ð¢&VÆF–öç6†—öGG&–'WFRÒâ‚'#¦–B"Ð¢f÷"–æFW‚Â6V7F–öâ–âVçVÖW&FR‡6V7F–öç2“ Ð¢F&vWBÒFVÆWFVEö–G2–b–æFW‚–âFVÆWFVE÷6V7F–öç2VÇ6R&WF–æVEö–G0Ð¢f÷"&VfW&Væ6R–â6V7F–öâç‡F‚€Ð¢"â÷s¦†VFW%&VfW&Væ6RÂâ÷s¦fö÷FW%&VfW&Væ6R"ÂæÖW76W3Ôå0Ð¢“ Ð¢&VÆF–öç6†—ö–BÒ&VfW&Væ6RævWB‡&VÆF–öç6†—öGG&–'WFRÐ¢–b&VÆF–öç6†—ö–C Ð¢F&vWBæFB‡&VÆF–öç6†—ö–BÐ¢&VÖ÷f&ÆUö–G2ÒFVÆWFVEö–G2Ò&WF–æVEö–G0Ð¢–bæ÷B&VÖ÷f&ÆUö–G3 Ð¢&WGW&â6WB‚Ð¢&VÆF–öç6†—5öæÖRÒ'v÷&Bõ÷&VÇ2öFö7VÖVçBç†ÖÂç&VÇ2 Ð¢–b&VÆF–öç6†—5öæÖRæ÷B–â6¶vRææÖVÆ—7B‚“ Ð¢&WGW&â6WB‚Ð¢&VÆF–öç6†—2ÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‡&VÆF–öç6†—5öæÖR’Ð¢&W7VÇBÒ6WB‚Ð¢f÷"&VÆF–öç6†—–â&VÆF–öç6†—3 Ð¢–b&VÆF–öç6†—ævWB‚$–B"’æ÷B–â&VÖ÷f&ÆUö–G3 Ð¢6öçF–çVPÐ¢F&vWBÒ&VÆF–öç6†—ævWB‚%F&vWB"Â""Ð¢æ÷&ÖÆ—¦VBÒ÷6—‡F‚ææ÷&×F‚‡÷6—‡F‚æ¦ö–â‚'v÷&B"ÂF&vWB’Ð¢–b&RægVÆÆÖF6‚‡"'v÷&Bòƒó¦†VFW'Æfö÷FW"•ÆBµÂç†ÖÂ"Âæ÷&ÖÆ—¦VB“ Ð¢&W7VÇBæFB†æ÷&ÖÆ—¦VBÐ¢&WGW&â&W7VÇ@Ð Ð Ð¦FVbö&÷fVE÷F&ÆUö6VÆÅö6ÆVçW÷&w&‡2€Ð¢&ö÷C¢WG&VRåôVÆVÖVçBÂ7G'V7GW&UöÖ¢F–7E·7G"Âç•ÐÐ¢’Óâ6WE´ç•Ó Ð¢F&ÆW2Ò&ö÷Bç‡F‚‚"÷s¦Fö7VÖVçB÷s¦&öG’÷s§F&Â"ÂæÖW76W3Ôå2Ð¢–væ÷&VC¢6WE´ç•ÒÒ6WB‚Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'F&ÆUö6VÆÅö6ÆVçW2"ÂµÒ“ Ð¢–bæ÷BVçG'’ævWB‚&&÷fVB"“ Ð¢6öçF–çVPÐ¢F&ÆUö–æFW‚Ò–çB†VçG'•²'F&ÆR%ÒÐ¢&÷uö–æFW‚Ò–çB†VçG'•²'&÷r%ÒÐ¢6VÆÅö–æFW‚Ò–çB†VçG'•²&6VÆÂ%ÒÐ¢–bæ÷BÃÒF&ÆUö–æFW‚ÂÆVâ‡F&ÆW2“ Ð¢6öçF–çVPÐ¢&÷w2ÒF&ÆW5·F&ÆUö–æFW…Òç‡F‚‚"â÷s§G""ÂæÖW76W3Ôå2Ð¢–bæ÷BÃÒ&÷uö–æFW‚ÂÆVâ‡&÷w2“ Ð¢6öçF–çVPÐ¢6VÆÇ2Ò&÷w5·&÷uö–æFW…Òç‡F‚‚"â÷s§F2"ÂæÖW76W3Ôå2Ð¢–bæ÷BÃÒ6VÆÅö–æFW‚ÂÆVâ†6VÆÇ2“ Ð¢6öçF–çVPÐ¢&w&‡2Ò6VÆÇ5¶6VÆÅö–æFW…Òç‡F‚‚"â÷s§"ÂæÖW76W3Ôå2Ð¢f÷"&w&‚–â&w&‡5³¢–çB†VçG'•²&6÷VçB%Ò•Ó Ð¢FW‡BÒ""æ¦ö–â€Ð¢&w&‚ç‡F‚‚"âò÷s§B÷FW‡B‚’"ÂæÖW76W3Ôå2Ð¢’ç7G&—‚Ð¢&÷FV7FVBÒ&w&‚ç‡F‚€Ð¢"âò÷s§F&ÂÂâò÷s¦G&v–ærÂâò÷s¦ö&¦V7BÂâò÷s§–7BÂ Ð¢"âò÷s¦&öö¶Ö&µ7F'BÂâò÷s¦6öÖÖVçE&ævU7F'BÂ Ð¢"âò÷s¦fö÷Fæ÷FU&VfW&Væ6RÂâò÷s¦VæFæ÷FU&VfW&Væ6RÂâò÷s¦fÆD6†""ÀÐ¢æÖW76W3Ôå2ÀÐ¢Ð¢–bæ÷BFW‡BæBæ÷B&÷FV7FVC Ð¢–væ÷&VBæFB‡&w&‚Ð¢&WGW&â–væ÷&V@Ð Ð Ð¦FVb7G'V7GW&Uö6öçFVçEö–çfVçF÷'’€Ð¢Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•ÐÐ¢’ÓâF–7E·7G"ÂÆ—7E·7G%ÕÓ Ð¢Fö5ö–æFW†W2Ò°Ð¢–æFW€Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'Fö5÷&ævW2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢f÷"–æFW‚–â&ævR†–çB†VçG'•²'7F'E÷&w&‚%Ò’Â–çB†VçG'•²&VæE÷&w&‚%Ò’²Ð¢ÐÐ¢†VF–æuöVçG&–W2Ò°Ð¢VçG'’f÷"VçG'’–â7G'V7GW&UöÖævWB‚&†VF–æw2"ÂµÒ’–bVçG'’ævWB‚&&÷fVB"Ð¢ÐÐ¢6F–öåöVçG&–W2Ò°Ð¢VçG'Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢æB€Ð¢æ÷B†5ö6F–öåö7F–öç5öÖ‡7G'V7GW&UöÖÐ¢÷"VçG'’ævWB‚&7F–öâ"’ÓÒ&6öçfW'E÷Fõ÷6W Ð¢Ð¢ÐÐ¢–FVçF–f–W%÷&WÆ6VÖVçG2Ò°Ð¢VçG'Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"’æBVçG'’ævWB‚&7F–öâ"’ÓÒ'&WÆ6Uö–FVçF–f–W" Ð¢ÐÐ¢Ö–w&FVEö6F–öåö†6†W2Ò°Ð¢VçG'•²'FW‡E÷6†#Sb%ÐÐ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"Ð¢æBVçG'’ævWB‚&Æö6F÷""Â·Ò’ævWB‚&¶–æB"’ÓÒ'F&ÆUö6VÆÅ÷&w&‚ Ð¢ÐÐ¢ÖçVÅöÖ–w&FVEö6F–öåö†6†W2Ò°Ð¢VçG'•²'FW‡E÷6†#Sb%ÐÐ¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚&6F–öç2"ÂµÒÐ¢–bVçG'’ævWB‚&&÷fVB"Ð¢æBVçG'’ævWB‚&7F–öâ"’ÓÒ&Ö÷fUö6F–öâ Ð¢æBVçG'’ævWB‚&Ö–w&FUö÷WG6–FU÷F&ÆR"Ð¢ÐÐ¢†5öÖ–w&FVEö6F–öç2Ò&ööÂ€Ð¢Ö–w&FVEö6F–öåö†6†W2ÒÖçVÅöÖ–w&FVEö6F–öåö†6†W0Ð¢Ð¢&÷fVEög&öçEöÖGFW"Ò7G'V7GW&UöÖævWB‚&g&öçEöÖGFW""Â·Ò’ævWB‚&&÷fVB"Ð¢&÷fVEö&Æö6µ÷76–ærÒ7G'V7GW&UöÖævWB‚&&Æö6µ÷76–ær"Â·Ò’ævWB‚&&÷fVB"Ð¢Fö5ö†VF–æu÷7G–ÆUö–BÒDô5ô„TD”äuõ5E”ÄRç&WÆ6R‚""Â""Ð¢&Æö6µ÷76W%÷7G–ÆUö–BÒ$Äô4µõ54U%õ5E”ÄRç&WÆ6R‚""Â""Ð¢&W7VÇC¢F–7E·7G"ÂÆ—7E·7G%ÕÒÒ·ÐÐ¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ð¢Fö7VÖVçE÷&ö÷BÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‚'v÷&BöFö7VÖVçBç†ÖÂ"’Ð¢–væ÷&VE÷'G2Òö&÷fVEöFVÆWFVEö†VFW%öfö÷FW%÷'G2€Ð¢6¶vRÂFö7VÖVçE÷&ö÷BÂ7G'V7GW&UöÖ Ð¢Ð¢f÷"æÖR–â6÷'FVB‡6¶vRææÖVÆ—7B‚’“ Ð¢–bæ÷B4ôåDTåEõ%BæÖF6‚†æÖR’÷"æÖR–â–væ÷&VE÷'G3 Ð¢6öçF–çVPÐ¢&ö÷BÒFö7VÖVçE÷&ö÷B–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6RWG&VRæg&ö×7G&–ær€Ð¢6¶vRç&VB†æÖRÐ¢Ð¢Fö5öf–VÆE÷&w&‡2Ò€Ð¢÷Fö5öf–VÆE÷&w&‡2‡&ö÷B’–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"VÇ6R6WB‚Ð¢Ð¢ÆVv7•÷Fö5öæ6†÷'2Ò€Ð¢öÆVv7•÷Fö5öV×G•öæ6†÷'2‡&ö÷BÂFö5öf–VÆE÷&w&‡2Â7G'V7GW&UöÖÐ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢VÇ6R6WB‚Ð¢Ð¢&÷fVE÷F–Å÷&w&‡2Ò€Ð¢ö&÷fVE÷F–Å÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖÐ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢VÇ6R6WB‚Ð¢Ð¢v–æF–öåö&÷VæF&–W2Ò€Ð¢÷v–æF–öåö&÷VæF'•÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖÐ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢VÇ6R6WB‚Ð¢Ð¢&÷fVEö6VÆÅö6ÆVçW÷&w&‡2Ò€Ð¢ö&÷fVE÷F&ÆUö6VÆÅö6ÆVçW÷&w&‡2‡&ö÷BÂ7G'V7GW&UöÖÐ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢VÇ6R6WB‚Ð¢Ð¢fÇVW2ÒµÐÐ¢&öG•ö–æFW‚ÒÓÐ¢f÷"&w&‚–â&ö÷Bç‡F‚‚"âò÷s§"ÂæÖW76W3Ôå2“ Ð¢–b&w&‚–â&÷fVEö6VÆÅö6ÆVçW÷&w&‡3 Ð¢6öçF–çVPÐ¢7G–ÆUö–G2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ð¢–b&÷fVEög&öçEöÖGFW"æB7G–ÆUö–G2ÓÒ·Fö5ö†VF–æu÷7G–ÆUö–EÓ Ð¢6öçF–çVPÐ¢–b&÷fVEö&Æö6µ÷76–æræB7G–ÆUö–G2ÓÒ¶&Æö6µ÷76W%÷7G–ÆUö–EÓ Ð¢6öçF–çVPÐ¢F—&V7Eö&öG•÷&w&‚Ò€Ð¢æÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ Ð¢æB&w&‚ævWG&VçB‚’—2æ÷BæöæPÐ¢æB&w&‚ævWG&VçB‚’çFrÓÒâ‚'s¦&öG’"Ð¢Ð¢–bF—&V7Eö&öG•÷&w&ƒ Ð¢&öG•ö–æFW‚³ÒÐ¢7W'&VçEö–æFW‚Ò&öG•ö–æFW‚–bF—&V7Eö&öG•÷&w&‚VÇ6RæöæPÐ¢–b&w&‚–â&÷fVE÷F–Å÷&w&‡2÷"&w&‚–âv–æF–öåö&÷VæF&–W3 Ð¢6öçF–çVPÐ¢–b&w&‚–âFö5öf–VÆE÷&w&‡2÷"&w&‚–âÆVv7•÷Fö5öæ6†÷'3 Ð¢6öçF–çVPÐ¢–b€Ð¢7W'&VçEö–æFW‚—2æ÷BæöæPÐ¢æBæ÷BFö5öf–VÆE÷&w&‡0Ð¢æB7W'&VçEö–æFW‚–âFö5ö–æFW†W0Ð¢“ Ð¢6öçF–çVPÐ¢fÇVRÒ÷&w&…÷FW‡E÷v—F†÷WEöf–VÆE÷&W7VÇG2‡&w&‚Ð¢fÇVRÒd”TÄEôÔ$´U%õEDU$âç7V"‚""ÂfÇVRÐ¢fÇVRÒöæ÷&ÖÆ—¦UöÖçVÅö–FVçF–f–W"‡fÇVRÂ–FVçF–f–W%÷&WÆ6VÖVçG2Ð¢÷&–v–æÅö6F–öâÒ4D”ôåõEDU$âæÖF6‚‡fÇVRÐ¢–bFW‡E÷6†#Sb‡fÇVR’–âÖçVÅöÖ–w&FVEö6F–öåö†6†W3 Ð¢fÇVRÒ%µ´ÔõdTEôÔåTÅô4D”ôåÕÒ Ð¢VÆ–bFW‡E÷6†#Sb‡fÇVR’–âÖ–w&FVEö6F–öåö†6†W2æB÷&–v–æÅö6F–öã Ð¢fÇVRÒb'¶÷&–v–æÅö6F–öâæw&÷Wƒ—Ò¶÷&–v–æÅö6F–öâæw&÷WƒB—Ò Ð¢VÆ–b††5öÖ–w&FVEö6F–öç2÷"6F–öåöVçG&–W2’æBF—&V7Eö&öG•÷&w&ƒ Ð¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ð¢vVæW&FVBÒ&RæÖF6‚‡"%åÇ2¢ŽY»çÎŠ‚•Ç2¥²ÞûÈÞ(	N(	5ÕÇ2¢‚â¢’B"ÂfÇVRÐ¢–b7G–ÆW2ÓÒ²$6F–öâ%ÒæBvVæW&FVC Ð¢fÇVRÒb'¶vVæW&FVBæw&÷Wƒ—Ò¶vVæW&FVBæw&÷Wƒ"—Ò Ð¢†VF–æuöVçG'’ÒæW‡B€Ð¢€Ð¢VçG'Ð¢f÷"VçG'’–â†VF–æuöVçG&–W0Ð¢–bFW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"Ð¢’ÀÐ¢æöæRÀÐ¢Ð¢–b†VF–æuöVçG'’—2æ÷BæöæS Ð¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB††VF–æuöVçG'•²&ÆWfVÂ%Ò’’æÖF6‚‡fÇVRÐ¢–bÖF6ƒ Ð¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥ÐÐ¢VÇ6S Ð¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ð¢7G–ÆUöÖF6‚Ò€Ð¢&RægVÆÆÖF6‚‡"$†VF–ær…³ÓEÒ’"Â7G–ÆW5³Ò’–b7G–ÆW2VÇ6RæöæPÐ¢Ð¢–b7G–ÆUöÖF6ƒ Ð¢ÖF6‚Òö†VF–æu÷&Vf—…÷GFW&â†–çB‡7G–ÆUöÖF6‚æw&÷Wƒ’’’æÖF6‚€Ð¢fÇVPÐ¢Ð¢–bÖF6ƒ Ð¢fÇVRÒfÇVU¶ÖF6‚æVæB‚’¥ÐÐ¢–bç’€Ð¢FW‡E÷6†#Sb‡fÇVR’ÓÒVçG'’ævWB‚'FW‡E÷6†#Sb"Ð¢f÷"VçG'’–â6F–öåöVçG&–W0Ð¢“ Ð¢ÖF6‚Ò4D”ôåõEDU$âæÖF6‚‡fÇVRÐ¢–bÖF6ƒ Ð¢fÇVRÒb'¶ÖF6‚æw&÷Wƒ—Ò¶ÖF6‚æw&÷WƒB—Ò Ð¢VÇ6S Ð¢fÇVRÒ&Rç7V"‡"%âŽY»çÎŠ‚•Ç2µ²ÞûÈÞ(	N(	5ÓõÇ2¢"Â"%Ã"ÂfÇVRÐ¢fÇVW2æVæB‡fÇVRÐ¢–bæÖRÓÒ'v÷&BöFö7VÖVçBç†ÖÂ"÷"ç’‡fÇVW2“ Ð¢&W7VÇE¶æÖUÒÒfÇVW0Ð¢&÷fVEöFW&—fVEöfö÷FW%ööæÇ’Ò&ööÂ€Ð¢7G'V7GW&UöÖævWB‚'v–æF–öå÷6V7F–öç2"Â·Ò’ævWB‚&&÷fVB"Ð¢æBç’€Ð¢VçG'’ævWB‚&&÷fVEöFVÆWFR"Ð¢æBVçG'’ævWB‚&Wf–FVæ6R"Â·Ò’ævWB‚&&÷fVEöFW&—fVEöfö÷FW%ööæÇ’"Ð¢f÷"VçG'’–â7G'V7GW&UöÖævWB‚'G&–Æ–æuöV×G•÷6V7F–öç2"ÂµÒÐ¢Ð¢Ð¢6æöæ–6Ã¢F–7E·7G"ÂÆ—7E·7G%ÕÒÒ·ÐÐ¢†VFW'3¢Æ—7E·7G%ÒÒµÐÐ¢fö÷FW'3¢Æ—7E·7G%ÒÒµÐÐ¢f÷"æÖRÂfÇVW2–â&W7VÇBæ—FV×2‚“ Ð¢–b&RægVÆÆÖF6‚‡"'v÷&Bö†VFW%ÆBµÂç†ÖÂ"ÂæÖR“ Ð¢†VFW'2æW‡FVæB‡fÇVW2Ð¢VÆ–b&RægVÆÆÖF6‚‡"'v÷&Böfö÷FW%ÆBµÂç†ÖÂ"ÂæÖR“ Ð¢–bæ÷B&÷fVEöFW&—fVEöfö÷FW%ööæÇ“ Ð¢fö÷FW'2æW‡FVæB‡fÇVW2Ð¢VÇ6S Ð¢6æöæ–6Å¶æÖUÒÒfÇVW0Ð¢6æöæ–6Å²'v÷&Bõö†VFW'5ö'•ö6öçFVçB%ÒÒ6÷'FVB††VFW'2Ð¢6æöæ–6Å²'v÷&Bõöfö÷FW'5ö'•ö6öçFVçB%ÒÒ6÷'FVB†fö÷FW'2Ð¢&WGW&â6æöæ–6ÀÐ Ð Ð¦FVb7G'V7GW&Uö6öçFVçEöf–ævW'&–çB‡Fƒ¢F‚Â7G'V7GW&UöÖ¢F–7E·7G"Âç•Ò’Óâ7G# Ð¢&W7VÇBÒ7G'V7GW&Uö6öçFVçEö–çfVçF÷'’‡F‚Â7G'V7GW&UöÖÐ¢Væ6öFVBÒ§6öâæGV×2‡&W7VÇBÂVç7W&Uö66–“ÔfÇ6RÂ6W&F÷'3Ò‚"Â"Â#¢"’Â6÷'Eö¶W—3ÕG'VRÐ¢&WGW&â†6†Æ–"ç6†#Sb†Væ6öFVBæVæ6öFR‚'WFbÓ‚"’’æ†W†F–vW7B‚Ð