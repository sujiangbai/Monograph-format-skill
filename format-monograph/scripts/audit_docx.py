#!/usr/bin/env python3
"""Audit content, protected objects, and automatically verifiable profile rules."""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any

from docx.oxml.ns import qn

from _common import (
    STYLE_PROPERTIES,
    FormatMonographError,
    _field_instruction_for_marker,
    _heading_prefix_pattern,
    _table_effective_properties,
    content_fingerprint,
    equation_inventory,
    field_cache_inventory,
    font_alias_keys,
    load_document,
    protected_object_manifest,
    protected_payload_manifest,
    run_effective_font,
    style_effective_font,
    style_name_for_selector,
    write_json,
)
from validate_profile import validate
from docx_pagination import audit_pagination_sections
from structure_map import (
    approved_data_tables,
    approved_role_paragraphs,
    audit_caption_identifier_replacements,
    audit_structure_heading_operations,
    audit_structure_image_operations,
    audit_structure_table_operations,
    has_semantic_structure_map,
    load_structure_map,
    resolve_paragraph_locator,
    structure_content_fingerprint,
)


def close(actual: Any, expected: Any, tolerance: float = 0.05) -> bool:
    if actual is None:
        return expected is None
    return math.isclose(float(actual), float(expected), abs_tol=tolerance)


def style_value(document: Any, style: Any, key: str) -> Any:
    font, pf = style.font, style.paragraph_format
    if key == "font_name":
        return style_effective_font(document, style, "ascii")[0]
    if key == "font_name_ascii":
        return style_effective_font(document, style, "ascii")[0]
    if key == "font_name_east_asia":
        return style_effective_font(document, style, "eastAsia")[0]
    if key == "font_name_complex_script":
        return style_effective_font(document, style, "cs")[0]
    if key == "font_size_pt":
        current = style
        while current is not None:
            if current.font.size is not None:
                return current.font.size.pt
            current = current.base_style
        return None
    if key in {"bold", "italic"}:
        current = style
        while current is not None:
            value = getattr(current.font, key)
            if value is not None:
                return value
            current = current.base_style
        return False
    if key == "color_hex":
        return None if font.color.rgb is None else str(font.color.rgb)
    if key == "alignment":
        return None if pf.alignment is None else str(pf.alignment).split(".")[-1].lower()
    point_values = {
        "space_before_pt": pf.space_before,
        "space_after_pt": pf.space_after,
        "first_line_indent_pt": pf.first_line_indent,
        "left_indent_pt": pf.left_indent,
        "right_indent_pt": pf.right_indent,
    }
    if key in point_values:
        attribute = {
            "space_before_pt": "space_before",
            "space_after_pt": "space_after",
            "first_line_indent_pt": "first_line_indent",
            "left_indent_pt": "left_indent",
            "right_indent_pt": "right_indent",
        }[key]
        current = style
        while current is not None:
            value = getattr(current.paragraph_format, attribute)
            if value is not None:
                return value.pt
            current = current.base_style
        return 0.0 if key in {"space_before_pt", "space_after_pt"} else None
    if key == "first_line_indent_chars":
        current = style
        while current is not None:
            p_pr = current.element.pPr
            ind = None if p_pr is None else p_pr.find(qn("w:ind"))
            if ind is not None:
                value = ind.get(qn("w:firstLineChars"))
                if value is not None:
                    return int(value) / 100
                point_value = ind.get(qn("w:firstLine"))
                if point_value == "0":
                    return 0
            current = current.base_style
        return None
    if key == "line_spacing":
        value = pf.line_spacing
        return float(value) if isinstance(value, (int, float)) else None
    if key == "line_spacing_pt":
        value = pf.line_spacing
        return None if value is None or not hasattr(value, "pt") else value.pt
    if key == "line_spacing_rule":
        value = str(pf.line_spacing_rule).lower()
        if "exact" in value:
            return "exact"
        if "at_least" in value or "at least" in value:
            return "at_least"
        if "multiple" in value:
            return "multiple"
        if "one_point_five" in value or "1.5" in value:
            return "one_point_five"
        if "double" in value:
            return "double"
        if "single" in value:
            return "single"
        return value
    if key in {"keep_with_next", "keep_together", "page_break_before", "widow_control"}:
        current = style
        while current is not None:
            value = getattr(current.paragraph_format, key)
            if value is not None:
                return value
            current = current.base_style
        if key == "widow_control":
            return True
        return False
    return "<unsupported>"


def normalized_alignment(value: str) -> str:
    value = value.lower()
    for name in ("left", "center", "right", "justify", "distributed"):
        if name in value:
            return name
    return value


def compare_value(key: str, actual: Any, expected: Any) -> bool:
    if isinstance(actual, dict) and isinstance(expected, dict):
        return set(actual) == set(expected) and all(
            compare_value(
                f"{name}_mm" if key == "cell_margins_mm" else name,
                actual[name],
                expected[name],
            )
            for name in actual
        )
    if isinstance(actual, list):
        if isinstance(expected, list):
            return len(actual) == len(expected) and all(
                compare_value(key, left, right)
                for left, right in zip(actual, expected)
            )
        return bool(actual) and all(compare_value(key, item, expected) for item in actual)
    if key.startswith("font_name") and actual is not None:
        return bool(font_alias_keys(str(actual)) & font_alias_keys(str(expected)))
    if key.endswith(("_pt", "_mm", "_ratio")) or key in {
        "line_spacing",
        "first_line_indent_chars",
    }:
        return close(actual, expected)
    if key == "color_hex" and actual is not None:
        return str(actual).lstrip("#").upper() == str(expected).lstrip("#").upper()
    if key == "alignment" and actual is not None:
        return normalized_alignment(str(actual)) == str(expected).lower()
    return actual == expected


def run_font_value(document: Any, paragraph: Any, run: Any, key: str) -> Any:
    r_pr = run._r.rPr
    if key == "font_name":
        return run_effective_font(document, paragraph, run, "ascii")[0]
    if key == "font_name_ascii":
        return run_effective_font(document, paragraph, run, "ascii")[0]
    if key == "font_name_east_asia":
        return run_effective_font(document, paragraph, run, "eastAsia")[0]
    if key == "font_name_complex_script":
        return run_effective_font(document, paragraph, run, "cs")[0]
    if key == "font_size_pt":
        return None if run.font.size is None else run.font.size.pt
    if key in {"bold", "italic"}:
        return getattr(run.font, key)
    if key == "color_hex":
        return None if run.font.color.rgb is None else str(run.font.color.rgb)
    return None


def paragraph_effective_value(document: Any, paragraph: Any, key: str) -> Any:
    style = paragraph.style
    if key.startswith("font_name"):
        values = [
            run_font_value(document, paragraph, run, key)
            for run in paragraph.runs
            if run.text
        ]
        return values or style_value(document, style, key)
    if key in {
        "font_size_pt",
        "bold",
        "italic",
        "color_hex",
    }:
        fallback = style_value(document, style, key)
        values = []
        for run in paragraph.runs:
            if not run.text:
                continue
            direct = run_font_value(document, paragraph, run, key)
            values.append(fallback if direct is None else direct)
        return values or fallback

    pf = paragraph.paragraph_format
    direct: Any = None
    if key == "alignment":
        direct = pf.alignment
        if direct is not None:
            direct = str(direct).split(".")[-1].lower()
    elif key in {
        "space_before_pt",
        "space_after_pt",
        "first_line_indent_pt",
        "left_indent_pt",
        "right_indent_pt",
    }:
        attr = {
            "space_before_pt": "space_before",
            "space_after_pt": "space_after",
            "first_line_indent_pt": "first_line_indent",
            "left_indent_pt": "left_indent",
            "right_indent_pt": "right_indent",
        }[key]
        value = getattr(pf, attr)
        direct = None if value is None else value.pt
    elif key == "first_line_indent_chars":
        p_pr = paragraph._p.pPr
        ind = None if p_pr is None else p_pr.find(qn("w:ind"))
        value = None if ind is None else ind.get(qn("w:firstLineChars"))
        direct = None if value is None else int(value) / 100
        if direct is None and ind is not None and ind.get(qn("w:firstLine")) == "0":
            direct = 0
        if direct is None and paragraph.style is not None:
            match = re.fullmatch(r"Heading ([1-4])", paragraph.style.name)
            if match:
                lvl = heading_numbering_level(document, int(match.group(1)) - 1)
                lvl_ind = None if lvl is None else lvl.find(
                    qn("w:pPr") + "/" + qn("w:ind")
                )
                if lvl_ind is not None:
                    chars = lvl_ind.get(qn("w:firstLineChars"))
                    if chars is not None:
                        direct = int(chars) / 100
                    elif lvl_ind.get(qn("w:firstLine")) == "0":
                        direct = 0
    elif key == "line_spacing":
        value = pf.line_spacing
        direct = float(value) if isinstance(value, (int, float)) else None
    elif key == "line_spacing_pt":
        value = pf.line_spacing
        direct = None if value is None or not hasattr(value, "pt") else value.pt
    elif key == "line_spacing_rule":
        value = pf.line_spacing_rule
        if value is not None:
            raw = str(value).lower()
            direct = next(
                (
                    name
                    for name, marker in (
                        ("exact", "exact"),
                        ("at_least", "at_least"),
                        ("multiple", "multiple"),
                        ("one_point_five", "one_point_five"),
                        ("double", "double"),
                        ("single", "single"),
                    )
                    if marker in raw
                ),
                raw,
            )
    elif key in {"keep_with_next", "keep_together", "page_break_before", "widow_control"}:
        direct = getattr(pf, key)
    return style_value(document, style, key) if direct is None else direct


def audit_paragraph_rule(
    document: Any, rule: dict[str, Any], paragraphs: list[Any]
) -> list[dict[str, Any]]:
    failures = []
    for target_index, paragraph in enumerate(paragraphs):
        for key, expected in rule["properties"].items():
            if key not in STYLE_PROPERTIES:
                continue
            actual = paragraph_effective_value(document, paragraph, key)
            if (
                key == "page_break_before"
                and expected is True
                and actual is False
                and _paragraph_starts_new_page_section(paragraph)
            ):
                continue
            if not compare_value(key, actual, expected):
                failures.append(
                    {
                        "target": target_index,
                        "property": key,
                        "expected": expected,
                        "actual": actual,
                        "reason": "effective format differs after style and direct formatting",
                    }
                )
    return failures


def _paragraph_starts_new_page_section(paragraph: Any) -> bool:
    previous = paragraph._p.getprevious()
    if previous is None or previous.tag != qn("w:p"):
        return False
    p_pr = previous.find(qn("w:pPr"))
    sect_pr = None if p_pr is None else p_pr.find(qn("w:sectPr"))
    if sect_pr is None:
        return False
    section_type = sect_pr.find(qn("w:type"))
    section_value = (
        section_type.get(qn("w:val")) if section_type is not None else "nextPage"
    )
    return section_value != "continuous"


def audit_style_rule(document: Any, rule: dict) -> list[dict]:
    style_name = style_name_for_selector(rule["selector"])
    if style_name is None:
        return [{"property": "*", "expected": "supported style selector", "actual": "unsupported"}]
    try:
        style = document.styles[style_name]
    except KeyError:
        return [{"property": "*", "expected": style_name, "actual": "missing style"}]
    failures = []
    for key, expected in rule["properties"].items():
        if key not in STYLE_PROPERTIES:
            continue
        actual = style_value(document, style, key)
        if not compare_value(key, actual, expected):
            failures.append({"property": key, "expected": expected, "actual": actual})
    return failures


def document_toggle(document: Any, name: str) -> bool:
    return document.settings.element.find(qn(f"w:{name}")) is not None


def audit_section_rule(
    document: Any, rule: dict, structure_map: dict[str, Any] | None = None
) -> list[dict]:
    failures = []
    for index, section in enumerate(document.sections):
        width = float(section.page_width.mm)
        height = float(section.page_height.mm)
        values = {
            "page_width_mm": width,
            "page_height_mm": height,
            "margin_top_mm": section.top_margin.mm,
            "margin_bottom_mm": section.bottom_margin.mm,
            "margin_left_mm": section.left_margin.mm,
            "margin_right_mm": section.right_margin.mm,
            "gutter_mm": section.gutter.mm,
            "header_distance_ratio": section.header_distance.mm / height,
            "footer_distance_ratio": section.footer_distance.mm / height,
        ã~ô¶‰žËkºwµçL9½¹”è4(€€€€€€€É•ÑÕÉ¸9½¹”4(€€€…‰ÍÑÉ…Ñ}¥€ô…‰ÍÑÉ…Ñ}É•˜¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤¤4(€€€…‰ÍÑÉ…Ð€ô¹•áÐ 4(€€€€€€€€ 4(€€€€€€€€€€€¥Ñ•´4(€€€€€€€€€€€™½È¥Ñ•´¥¸É½½Ð¹™¥¹‘…±°¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õ´ˆ¤¤4(€€€€€€€€€€€¥˜¥Ñ•´¹•Ð¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õµ%ˆ¤¤€ôô…‰ÍÑÉ…Ñ}¥4(€€€€€€€€¤°4(€€€€€€€9½¹”°4(€€€€¤4(€€€¥˜…‰ÍÑÉ…Ð¥Ì9½¹”è4(€€€€€€€É•ÑÕÉ¸9½¹”4(€€€É•ÑÕÉ¸¹•áÐ 4(€€€€€€€€ 4(€€€€€€€€€€€¥Ñ•´4(€€€€€€€€€€€™½È¥Ñ•´¥¸…‰ÍÑÉ…Ð¹™¥¹‘…±°¡Å¸ ‰Üé±Ù°ˆ¤¤4(€€€€€€€€€€€¥˜¥Ñ•´¹•Ð¡Å¸ ‰Üé¥±Ù°ˆ¤¤€ôôÍÑÈ¡±•Ù•°¤4(€€€€€€€€¤°4(€€€€€€€9½¹”°4(€€€€¤4(4(4)‘•˜¡•…‘¥¹}¹Õµ‰•É¥¹}™½Éµ…Ñ}™…¥±ÕÉ•Ì¡‘½Õµ•¹Ðè¹ä°±•Ù•±Ìè¥¹Ð¤€´ø±¥ÍÑm‘¥Ñtè4(€€€™…¥±ÕÉ•Ì€ômt4(€€€™½È±•Ù•°¥¸É…¹”¡±•Ù•±Ì¤è4(€€€€€€€ÍÑå±”€ô‘½Õµ•¹Ð¹ÍÑå±•Ím˜‰!•…‘¥¹œí±•Ù•°€¬€Åô‰t4(€€€€€€€±Ù°€ô¡•…‘¥¹}¹Õµ‰•É¥¹}±•Ù•°¡‘½Õµ•¹Ð°±•Ù•°¤4(€€€€€€€¥˜±Ù°¥Ì9½¹”è4(€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€¥¹€ô±Ù°¹™¥¹¡Å¸ ‰ÜéÁAÈˆ¤€¬€ˆ¼ˆ€¬Å¸ ‰Üé¥¹ˆ¤¤4(€€€€€€€¡…Í}é•É½}¥¹‘•¹Ð€ô¥¹¥Ì¹½Ð9½¹”…¹€ 4(€€€€€€€€€€€¥¹¹•Ð¡Å¸ ‰Üé™¥ÉÍÑ1¥¹•¡…ÉÌˆ¤¤€ôô€ˆÀˆ4(€€€€€€€€€€€½È¥¹¹•Ð¡Å¸ ‰Üé™¥ÉÍÑ1¥¹”ˆ¤¤€ôô€ˆÀˆ4(€€€€€€€€¤4(€€€€€€€¥˜¹½Ð¡…Í}é•É½}¥¹‘•¹Ðè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰¡•…‘¥¹}™¥ÉÍÑ}±¥¹•}¥¹‘•¹Ðˆ°4(€€€€€€€€€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•°ˆè±•Ù•°€¬€Ä°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€À°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€É}ÁÈ€ô±Ù°¹™¥¹¡Å¸ ‰ÜéÉAÈˆ¤¤4(€€€€€€€É}™½¹ÑÌ€ô9½¹”¥˜É}ÁÈ¥Ì9½¹”•±Í”É}ÁÈ¹™¥¹¡Å¸ ‰ÜéÉ½¹ÑÌˆ¤¤4(€€€€€€€•áÁ•Ñ•‘}™½¹ÑÌ€ôì4(€€€€€€€€€€€€‰…Í¥¤ˆèÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°ÍÑå±”°€‰…Í¥¤ˆ¥lÁt°4(€€€€€€€€€€€€‰¡¹Í¤ˆèÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°ÍÑå±”°€‰…Í¥¤ˆ¥lÁt°4(€€€€€€€€€€€€‰•…ÍÑÍ¥„ˆèÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°ÍÑå±”°€‰•…ÍÑÍ¥„ˆ¥lÁt°4(€€€€€€€€€€€€‰ÌˆèÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°ÍÑå±”°€‰Ìˆ¥lÁt°4(€€€€€€€ô4(€€€€€€€…ÑÕ…±}™½¹ÑÌ€ôì4(€€€€€€€€€€€…ÑÑÉ¥‰ÕÑ”è9½¹”¥˜É}™½¹ÑÌ¥Ì9½¹”•±Í”É}™½¹ÑÌ¹•Ð¡Å¸¡˜‰Üéí…ÑÑÉ¥‰ÕÑ•ôˆ¤¤4(€€€€€€€€€€€™½È…ÑÑÉ¥‰ÕÑ”¥¸•áÁ•Ñ•‘}™½¹ÑÌ4(€€€€€€€ô4(€€€€€€€¥˜…ÑÕ…±}™½¹ÑÌ€„ô•áÁ•Ñ•‘}™½¹ÑÌè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰¡•…‘¥¹}¹Õµ‰•É¥¹}™½¹ÑÌˆ°4(€€€€€€€€€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•°ˆè±•Ù•°€¬€Ä°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}™½¹ÑÌ°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}™½¹ÑÌ°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€•áÁ•Ñ•‘}Í¥é”€ôÍÑå±•}Ù…±Õ”¡‘½Õµ•¹Ð°ÍÑå±”°€‰™½¹Ñ}Í¥é•}ÁÐˆ¤4(€€€€€€€Í¥é”€ô9½¹”¥˜É}ÁÈ¥Ì9½¹”•±Í”É}ÁÈ¹™¥¹¡Å¸ ‰ÜéÍèˆ¤¤4(€€€€€€€…ÑÕ…±}Í¥é”€ô9½¹”¥˜Í¥é”¥Ì9½¹”•±Í”¥¹Ð¡Í¥é”¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤¤¤€¼€È4(€€€€€€€¥˜¹½Ð±½Í”¡…ÑÕ…±}Í¥é”°•áÁ•Ñ•‘}Í¥é”¤è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰¡•…‘¥¹}¹Õµ‰•É¥¹}™½¹Ñ}Í¥é•}ÁÐˆ°4(€€€€€€€€€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•°ˆè±•Ù•°€¬€Ä°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}Í¥é”°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}Í¥é”°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€•áÁ•Ñ•‘}‰½±€ôÍÑå±•}Ù…±Õ”¡‘½Õµ•¹Ð°ÍÑå±”°€‰‰½±ˆ¤4(€€€€€€€‰½±€ô9½¹”¥˜É}ÁÈ¥Ì9½¹”•±Í”É}ÁÈ¹™¥¹¡Å¸ ‰Üéˆˆ¤¤4(€€€€€€€…ÑÕ…±}‰½±€ô9½¹”¥˜‰½±¥Ì9½¹”•±Í”‰½±¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€ˆÄˆ¤¹½Ð¥¸ìˆÀˆ°€‰™…±Í”‰ô4(€€€€€€€¥˜…ÑÕ…±}‰½±€„ô•áÁ•Ñ•‘}‰½±è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰¡•…‘¥¹}¹Õµ‰•É¥¹}‰½±ˆ°4(€€€€€€€€€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•°ˆè±•Ù•°€¬€Ä°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}‰½±°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}‰½±°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€É•ÑÕÉ¸™…¥±ÕÉ•Ì4(4(4)‘•˜…Õ‘¥Ñ}™¥•±‘}ÉÕ±” 4(€€€‘½Õµ•¹Ðè¹ä°4(€€€ÉÕ±”è‘¥Ð°4(€€€¡…ÁÑ•É}ÍÑ…ÉÐè¥¹Ðð9½¹”€ô9½¹”°4(€€€Á…Ñ èA…Ñ ð9½¹”€ô9½¹”°4(¤€´ø±¥ÍÑm‘¥Ñtè4(€€€™…¥±ÕÉ•Ì€ômt4(€€€ÁÉ½Á•ÉÑ¥•Ì€ôÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t4(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰ÕÁ‘…Ñ•}½¹}½Á•¸ˆ¤…¹¹½Ð‘½Õµ•¹Ñ}Ñ½±”¡‘½Õµ•¹Ð°€‰ÕÁ‘…Ñ•¥•±‘Ìˆ¤è4(€€€€€€€É•™É•Í¡•€ôÁ…Ñ ¥Ì¹½Ð9½¹”…¹™¥•±‘}…¡•}¥¹Ù•¹Ñ½Éä¡Á…Ñ ¥l‰ÍÑ…ÑÕÌ‰t€ôô€‰É•™É•Í¡•ˆ4(€€€€€€€¥˜¹½ÐÉ•™É•Í¡•è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹¡ì‰ÁÉ½Á•ÉÑäˆè€‰ÕÁ‘…Ñ•}½¹}½Á•¸ˆ°€‰•áÁ•Ñ•ˆèQÉÕ”°€‰…ÑÕ…°ˆè…±Í•ô¤4(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰½¹Ù•ÉÑ}•áÁ±¥¥Ñ}µ…É­•ÉÌˆ¤è4(€€€€€€€É•µ…¥¹¥¹œ€ôl4(€€€€€€€€€€€Á…É…É…Á ¹Ñ•áÐ4(€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸‘½Õµ•¹Ð¹Á…É…É…Á¡Ì4(€€€€€€€€€€€¥˜}™¥•±‘}¥¹ÍÑÉÕÑ¥½¹}™½É}µ…É­•È¡Á…É…É…Á ¹Ñ•áÐ¹ÍÑÉ¥À ¤¤¥Ì¹½Ð9½¹”4(€€€€€€€t4(€€€€€€€¥˜É•µ…¥¹¥¹œè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰½¹Ù•ÉÑ}•áÁ±¥¥Ñ}µ…É­•ÉÌˆ°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€‰¹¼µ…É­•ÉÌˆ°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèÉ•µ…¥¹¥¹lèÄÁt°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€±•Ù•±Ì€ô¥¹Ð¡ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰¡•…‘¥¹}±•Ù•±Ìˆ°€Ð¤¤4(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆ¤è4(€€€€€€€™½È±•Ù•°¥¸É…¹” Ä°±•Ù•±Ì€¬€Ä¤è4(€€€€€€€€€€€ÍÑå±”€ô‘½Õµ•¹Ð¹ÍÑå±•Ím˜‰!•…‘¥¹œí±•Ù•±ô‰t4(€€€€€€€€€€€Á}ÁÈ€ôÍÑå±”¹•±•µ•¹Ð¹ÁAÈ4(€€€€€€€€€€€¥˜Á}ÁÈ¥Ì9½¹”½ÈÁ}ÁÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤¥Ì9½¹”è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè˜‰¹Õµ‰•É¥¹œ½¸!•…‘¥¹œí±•Ù•±ôˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè€‰µ¥ÍÍ¥¹œˆ°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€™…¥±ÕÉ•Ì¹•áÑ•¹¡¡•…‘¥¹}¹Õµ‰•É¥¹}™½Éµ…Ñ}™…¥±ÕÉ•Ì¡‘½Õµ•¹Ð°±•Ù•±Ì¤¤4(€€€€€€€¥˜¡…ÁÑ•É}ÍÑ…ÉÐ¥Ì¹½Ð9½¹”è4(€€€€€€€€€€€…ÑÕ…±}ÍÑ…ÉÐ€ô¡•…‘¥¹}¹Õµ‰•É¥¹}ÍÑ…ÉÐ¡‘½Õµ•¹Ð¤4(€€€€€€€€€€€¥˜…ÑÕ…±}ÍÑ…ÉÐ€„ô¡…ÁÑ•É}ÍÑ…ÉÐè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰¡…ÁÑ•É}ÍÑ…ÉÐˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè¡…ÁÑ•É}ÍÑ…ÉÐ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}ÍÑ…ÉÐ°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰ÍÑÉ¥Á}µ…¹Õ…±}¡•…‘¥¹}ÁÉ•™¥á•Ìˆ¤è4(€€€€€€€™½ÈÁ…É…É…Á ¥¸‘½Õµ•¹Ð¹Á…É…É…Á¡Ìè4(€€€€€€€€€€€¥˜¹½ÐÁ…É…É…Á ¹ÍÑå±”½È¹½ÐÁ…É…É…Á ¹ÍÑå±”¹¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰!•…‘¥¹œ€ˆ¤è4(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€€€€€ÑÉäè4(€€€€€€€€€€€€€€€±•Ù•°€ô¥¹Ð¡Á…É…É…Á ¹ÍÑå±”¹¹…µ”¹ÍÁ±¥Ð ¥l´Åt¤4(€€€€€€€€€€€•á•ÁÐY…±Õ•ÉÉ½Èè4(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€€€€€¥˜±•Ù•°€ðô±•Ù•±Ì…¹}¡•…‘¥¹}ÁÉ•™¥á}Á…ÑÑ•É¸¡±•Ù•°¤¹µ…Ñ ¡Á…É…É…Á ¹Ñ•áÐ¤è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰ÍÑÉ¥Á}µ…¹Õ…±}¡•…‘¥¹}ÁÉ•™¥á•Ìˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€‰¹¼µ…¹Õ…°ÁÉ•™¥àˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèÁ…É…É…Á ¹Ñ•áÑlèàÁt°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€É•ÑÕÉ¸™…¥±ÕÉ•Ì4(4(4)‘•˜…Õ‘¥Ñ}•ÅÕ…Ñ¥½¹}ÉÕ±”¡Á…Ñ èA…Ñ °ÉÕ±”è‘¥Ð¤€´ø±¥ÍÑm‘¥Ñtè4(€€€Ù…±Õ•Ì€ô•ÅÕ…Ñ¥½¹}¥¹Ù•¹Ñ½Éä¡Á…Ñ ¤4(€€€™…¥±ÕÉ•Ì€ômt4(€€€ÁÉ½Á•ÉÑ¥•Ì€ôÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t4(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰‰±½­}™½ÉµÕ±…}¥µ…•Ìˆ¤…¹Ù…±Õ•Íl‰™½ÉµÕ±…}¥µ…•}…¹‘¥‘…Ñ•Ì‰tè4(€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰±½­}™½ÉµÕ±…}¥µ…•Ìˆ°4(€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€À°4(€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèÙ…±Õ•Íl‰™½ÉµÕ±…}¥µ…•}…¹‘¥‘…Ñ•Ì‰t°4(€€€€€€€€€€€ô4(€€€€€€€€¤4(€€€É•ÑÕÉ¸™…¥±ÕÉ•Ì4(4(4)‘•˜ÕÍ•Í}‘•É¥Ù•‘}¹½Éµ…±¥é…Ñ¥½¸¡ÁÉ½™¥±”è‘¥Ð¤€´ø‰½½°è4(€€€É•ÑÕÉ¸…¹ä 4(€€€€€€€ÉÕ±”¹•Ð ‰ÍÑ…ÑÕÌˆ¤€ôô€‰…ÁÁÉ½Ù•ˆ4(€€€€€€€…¹ÉÕ±”¹•Ð ‰…ÁÁ±¥…Ñ¥½¸ˆ¤€ôô€‰…ÕÑ½µ…Ñ¥Œˆ4(€€€€€€€…¹ÉÕ±”¹•Ð ‰Í•±•Ñ½Èˆ°íô¤¹•Ð ‰­¥¹ˆ¤€ôô€‰™¥•±‘}É½±”ˆ4(€€€€€€€…¹…¹ä 4(€€€€€€€€€€€ÉÕ±”¹•Ð ‰ÁÉ½Á•ÉÑ¥•Ìˆ°íô¤¹•Ð¡­•ä¤4(€€€€€€€€€€€™½È­•ä¥¸€ 4(€€€€€€€€€€€€€€€€‰½¹Ù•ÉÑ}•áÁ±¥¥Ñ}µ…É­•ÉÌˆ°4(€€€€€€€€€€€€€€€€‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆ°4(€€€€€€€€€€€€€€€€‰ÍÑÉ¥Á}µ…¹Õ…±}¡•…‘¥¹}ÁÉ•™¥á•Ìˆ°4(€€€€€€€€€€€€¤4(€€€€€€€€¤4(€€€€€€€™½ÈÉÕ±”¥¸ÁÉ½™¥±”¹•Ð ‰ÉÕ±•Ìˆ°mt¤4(€€€€¤4(4(4)‘•˜µ…¥¸ ¤€´ø¥¹Ðè4(€€€Á…ÉÍ•È€ô…ÉÁ…ÉÍ”¹ÉÕµ•¹ÑA…ÉÍ•È ¤4(€€€Á…ÉÍ•È¹…‘‘}…ÉÕµ•¹Ð ‰½É¥¥¹…°ˆ°ÑåÁ”õA…Ñ ¤4(€€€Á…ÉÍ•È¹…‘‘}…ÉÕµ•¹Ð ‰™½Éµ…ÑÑ•ˆ°ÑåÁ”õA…Ñ ¤4(€€€Á…ÉÍ•È¹…‘‘}…ÉÕµ•¹Ð ˆ´µÁÉ½™¥±”ˆ°É•ÅÕ¥É•õQÉÕ”°ÑåÁ”õA…Ñ ¤4(€€€Á…ÉÍ•È¹…‘‘}…ÉÕµ•¹Ð ˆ´µÍÑÉÕÑÕÉ”µµ…Àˆ°ÑåÁ”õA…Ñ ¤4(€€€Á…ÉÍ•È¹…‘‘}…ÉÕµ•¹Ð ˆ´µ½ÕÑÁÕÐˆ°ÑåÁ”õA…Ñ ¤4(€€€…ÉÌ€ôÁ…ÉÍ•È¹Á…ÉÍ•}…ÉÌ ¤4(4(€€€ÑÉäè4(€€€€€€€ÁÉ½™¥±•}•ÉÉ½ÉÌ°ÁÉ½™¥±”€ôÙ…±¥‘…Ñ”¡…ÉÌ¹ÁÉ½™¥±”¤4(€€€€€€€¥˜ÁÉ½™¥±•}•ÉÉ½ÉÌè4(€€€€€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È ‰AÉ½™¥±”Ù…±¥‘…Ñ¥½¸™…¥±•è€ˆ€¬€ˆì€ˆ¹©½¥¸¡ÁÉ½™¥±•}•ÉÉ½ÉÌ¤¤4(€€€€€€€¹½Éµ…±¥é”€ôÕÍ•Í}‘•É¥Ù•‘}¹½Éµ…±¥é…Ñ¥½¸¡ÁÉ½™¥±”¤4(€€€€€€€ÍÑÉÕÑÕÉ•}µ…À€ô±½…‘}ÍÑÉÕÑÕÉ•}µ…À¡…ÉÌ¹ÍÑÉÕÑÕÉ•}µ…À¤¥˜…ÉÌ¹ÍÑÉÕÑÕÉ•}µ…À•±Í”9½¹”4(€€€€€€€½É¥¥¹…±}™À€ô€ 4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡…ÉÌ¹½É¥¥¹…°°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡…ÉÌ¹½É¥¥¹…°°¹½Éµ…±¥é•}‘•É¥Ù•õ¹½Éµ…±¥é”¤4(€€€€€€€€¤4(€€€€€€€™½Éµ…ÑÑ•‘}™À€ô€ 4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡…ÉÌ¹™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡…ÉÌ¹™½Éµ…ÑÑ•°¹½Éµ…±¥é•}‘•É¥Ù•õ¹½Éµ…±¥é”¤4(€€€€€€€€¤4(€€€€€€€½É¥¥¹…±}½‰©•ÑÌ€ôÁÉ½Ñ•Ñ•‘}½‰©•Ñ}µ…¹¥™•ÍÐ¡…ÉÌ¹½É¥¥¹…°¤4(€€€€€€€™½Éµ…ÑÑ•‘}½‰©•ÑÌ€ôÁÉ½Ñ•Ñ•‘}½‰©•Ñ}µ…¹¥™•ÍÐ¡…ÉÌ¹™½Éµ…ÑÑ•¤4(€€€€€€€½É¥¥¹…±}Á…å±½…‘Ì€ôÁÉ½Ñ•Ñ•‘}Á…å±½…‘}µ…¹¥™•ÍÐ¡…ÉÌ¹½É¥¥¹…°¤4(€€€€€€€™½Éµ…ÑÑ•‘}Á…å±½…‘Ì€ôÁÉ½Ñ•Ñ•‘}Á…å±½…‘}µ…¹¥™•ÍÐ¡…ÉÌ¹™½Éµ…ÑÑ•¤4(€€€€€€€½‰©•ÑÍ}½¬€ô½É¥¥¹…±}Á…å±½…‘Ì€ôô™½Éµ…ÑÑ•‘}Á…å±½…‘Ì4(€€€€€€€‘½Õµ•¹Ð€ô±½…‘}‘½Õµ•¹Ð¡…ÉÌ¹™½Éµ…ÑÑ•¤4(€€€€€€€ÉÕ±•}É•ÍÕ±ÑÌ€ômt4(4(€€€€€€€™½ÈÉÕ±”¥¸ÁÉ½™¥±•l‰ÉÕ±•Ì‰tè4(€€€€€€€€€€€¥˜ÉÕ±•l‰ÍÑ…ÑÕÌ‰t€„ô€‰…ÁÁÉ½Ù•ˆè4(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€€€€€¥˜ÉÕ±•l‰…ÁÁ±¥…Ñ¥½¸‰t€ôô€‰µ…¹Õ…±}É•Ù¥•Üˆè4(€€€€€€€€€€€€€€€ÉÕ±•}É•ÍÕ±ÑÌ¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì‰¥ˆèÉÕ±•l‰¥‰t°€‰ÍÑ…ÑÕÌˆè€‰µ…¹Õ…±}É•Ù¥•Üˆ°€‰™…¥±ÕÉ•Ìˆèmuô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€€€€€­¥¹€ôÉÕ±•l‰Í•±•Ñ½È‰ul‰­¥¹‰t4(€€€€€€€€€€€¥˜­¥¹¥¸ì‰‘½Õµ•¹Ðˆ°€‰Í•Ñ¥½¹}É½±”‰ôè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}Í•Ñ¥½¹}ÉÕ±”¡‘½Õµ•¹Ð°ÉÕ±”°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€•±¥˜­¥¹€ôô€‰Ñ…‰±•}É½±”ˆè4(€€€€€€€€€€€€€€€Ñ…‰±•}Ñ…É•ÑÌ€ô€ 4(€€€€€€€€€€€€€€€€€€€…ÁÁÉ½Ù•‘}‘…Ñ…}Ñ…‰±•Ì¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À…¹¡…Í}Í•µ…¹Ñ¥}ÍÑÉÕÑÕÉ•}µ…À¡ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€€€€€€€€€•±Í”9½¹”4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}Ñ…‰±•}ÉÕ±”¡‘½Õµ•¹Ð°ÉÕ±”°Ñ…‰±•}Ñ…É•ÑÌ¤4(€€€€€€€€€€€•±¥˜­¥¹€ôô€‰™¥•±‘}É½±”ˆè4(€€€€€€€€€€€€€€€¹Õµ‰•É¥¹œ€ôÍÑÉÕÑÕÉ•}µ…À¹•Ð ‰¹Õµ‰•É¥¹œˆ°íô¤¥˜ÍÑÉÕÑÕÉ•}µ…À•±Í”íô4(€€€€€€€€€€€€€€€¡…ÁÑ•É}ÍÑ…ÉÐ€ô€ 4(€€€€€€€€€€€€€€€€€€€¥¹Ð¡¹Õµ‰•É¥¹l‰¡…ÁÑ•É}ÍÑ…ÉÐ‰t¤4(€€€€€€€€€€€€€€€€€€€¥˜¹Õµ‰•É¥¹œ¹•Ð ‰…ÁÁÉ½Ù•ˆ¤4(€€€€€€€€€€€€€€€€€€€•±Í”9½¹”4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}™¥•±‘}ÉÕ±” 4(€€€€€€€€€€€€€€€€€€€‘½Õµ•¹Ð°ÉÕ±”°¡…ÁÑ•É}ÍÑ…ÉÐ°…ÉÌ¹™½Éµ…ÑÑ•4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€•±¥˜­¥¹€ôô€‰•ÅÕ…Ñ¥½¹}É½±”ˆè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}•ÅÕ…Ñ¥½¹}ÉÕ±”¡…ÉÌ¹™½Éµ…ÑÑ•°ÉÕ±”¤4(€€€€€€€€€€€•±¥˜€ 4(€€€€€€€€€€€€€€€ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€€€€€…¹¡…Í}Í•µ…¹Ñ¥}ÍÑÉÕÑÕÉ•}µ…À¡ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€€€€€…¹­¥¹¥¸ì‰Á…É…É…Á¡}É½±”ˆ°€‰…ÁÑ¥½¹}É½±”ˆ°€‰‰¥‰±¥½É…Á¡å}É½±”‰ô4(€€€€€€€€€€€€¤è4(€€€€€€€€€€€€€€€Á…É…É…Á¡Ì€ô…ÁÁÉ½Ù•‘}É½±•}Á…É…É…Á¡Ì 4(€€€€€€€€€€€€€€€€€€€‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ•}µ…À°ÉÕ±•l‰Í•±•Ñ½È‰t4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}Á…É…É…Á¡}ÉÕ±”¡‘½Õµ•¹Ð°ÉÕ±”°Á…É…É…Á¡Ì¤4(€€€€€€€€€€€•±Í”è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}ÍÑå±•}ÉÕ±”¡‘½Õµ•¹Ð°ÉÕ±”¤4(€€€€€€€€€€€ÉÕ±•}É•ÍÕ±ÑÌ¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰¥ˆèÉÕ±•l‰¥‰t°4(€€€€€€€€€€€€€€€€€€€€‰ÍÑ…ÑÕÌˆè€‰Á…ÍÌˆ¥˜¹½Ð™…¥±ÕÉ•Ì•±Í”€‰™…¥°ˆ°4(€€€€€€€€€€€€€€€€€€€€‰™…¥±ÕÉ•Ìˆè™…¥±ÕÉ•Ì°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(4(€€€€€€€Á…¥¹…Ñ¥½¹}™…¥±ÕÉ•Ì°Á…¥¹…Ñ¥½¸€ô€ 4(€€€€€€€€€€€…Õ‘¥Ñ}Á…¥¹…Ñ¥½¹}Í•Ñ¥½¹Ì 4(€€€€€€€€€€€€€€€…ÉÌ¹™½Éµ…ÑÑ•°4(€€€€€€€€€€€€€€€‘½Õµ•¹Ð°4(€€€€€€€€€€€€€€€ÍÑÉÕÑÕÉ•}µ…À¹•Ð ‰Á…¥¹…Ñ¥½¹}Í•Ñ¥½¹Ìˆ°íô¤°4(€€€€€€€€€€€€€€€É•Í½±Ù•}Á…É…É…Á¡}±½…Ñ½È°4(€€€€€€€€€€€€€€€ÍÑÉÕÑÕÉ•}µ…À°4(€€€€€€€€€€€€¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”€¡mt°íô¤4(€€€€€€€€¤4(€€€€€€€½¹Ñ•¹Ñ}½¬€ô½É¥¥¹…±}™À€ôô™½Éµ…ÑÑ•‘}™À(€€€€€€€ÍÑÉÕÑÕÉ•}¡•…‘¥¹}™…¥±ÕÉ•Ì€ô€ (€€€€€€€€€€€…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}¡•…‘¥¹}½Á•É…Ñ¥½¹Ì¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ•}µ…À¤(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À(€€€€€€€€€€€•±Í”mt(€€€€€€€€¤(€€€€€€€ÍÑÉÕÑÕÉ•}Ñ…‰±•}™…¥±ÕÉ•Ì€ô€ (€€€€€€€€€€€…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}Ñ…‰±•}½Á•É…Ñ¥½¹Ì¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”mt4(€€€€€€€€¤4(€€€€€€€ÍÑÉÕÑÕÉ•}¥µ…•}™…¥±ÕÉ•Ì€ô€ 4(€€€€€€€€€€€…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}¥µ…•}½Á•É…Ñ¥½¹Ì¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ•}µ…À¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”mt4(€€€€€€€€¤4(€€€€€€€ÉÕ±•Í}½¬€ô€ 4(€€€€€€€€€€€…±°¡¥Ñ•µl‰ÍÑ…ÑÕÌ‰t€„ô€‰™…¥°ˆ™½È¥Ñ•´¥¸ÉÕ±•}É•ÍÕ±ÑÌ¤(€€€€€€€€€€€…¹¹½ÐÁ…¥¹…Ñ¥½¹}™…¥±ÕÉ•Ì(€€€€€€€€€€€…¹¹½ÐÍÑÉÕÑÕÉ•}¡•…‘¥¹}™…¥±ÕÉ•Ì(€€€€€€€€€€€…¹¹½ÐÍÑÉÕÑÕÉ•}Ñ…‰±•}™…¥±ÕÉ•Ì(€€€€€€€€€€€…¹¹½ÐÍÑÉÕÑÕÉ•}¥µ…•}™…¥±ÕÉ•Ì4(€€€€€€€€¤4(€€€€€€€…ÁÑ¥½¹}É•Á±…•µ•¹ÑÌ€ô€ 4(€€€€€€€€€€€…Õ‘¥Ñ}…ÁÑ¥½¹}¥‘•¹Ñ¥™¥•É}É•Á±…•µ•¹ÑÌ 4(€€€€€€€€€€€€€€€…ÉÌ¹½É¥¥¹…°°…ÉÌ¹™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€€¤4(€€€€€€€€€€€¥˜ÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€•±Í”mt4(€€€€€€€€¤4(€€€€€€€…ÁÑ¥½¹}É•Á±…•µ•¹ÑÍ}½¬€ô…±° 4(€€€€€€€€€€€¥Ñ•µl‰ÍÑ…ÑÕÌ‰t€ôô€‰Á…ÍÌˆ™½È¥Ñ•´¥¸…ÁÑ¥½¹}É•Á±…•µ•¹ÑÌ4(€€€€€€€€¤4(€€€€€€€É•ÍÕ±Ð€ôì4(€€€€€€€€€€€€‰Á…ÍÍ•ˆè½¹Ñ•¹Ñ}½¬…¹½‰©•ÑÍ}½¬…¹ÉÕ±•Í}½¬…¹…ÁÑ¥½¹}É•Á±…•µ•¹ÑÍ}½¬°4(€€€€€€€€€€€€‰½¹Ñ•¹Ñ}¥¹Ñ•É¥Ñäˆèì4(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè½¹Ñ•¹Ñ}½¬°4(€€€€€€€€€€€€€€€€‰½É¥¥¹…±}Í¡„ÈÔØˆè½É¥¥¹…±}™À°4(€€€€€€€€€€€€€€€€‰™½Éµ…ÑÑ•‘}Í¡„ÈÔØˆè™½Éµ…ÑÑ•‘}™À°4(€€€€€€€€€€€€€€€€‰™¥•±‘}É•ÍÕ±ÑÍ}…¹‘}…ÁÁÉ½Ù•‘}‘•É¥Ù•‘}Ù…±Õ•Í}•á±Õ‘•ˆè‰½½° 4(€€€€€€€€€€€€€€€€€€€¹½Éµ…±¥é”½ÈÍÑÉÕÑÕÉ•}µ…À4(€€€€€€€€€€€€€€€€¤°4(€€€€€€€€€€€€€€€€‰¹½Éµ…±¥é…Ñ¥½¹}Í½ÕÉ•Ìˆèl4(€€€€€€€€€€€€€€€€€€€Í½ÕÉ”4(€€€€€€€€€€€€€€€€€€€™½ÈÍ½ÕÉ”°•¹…‰±•¥¸€ 4(€€€€€€€€€€€€€€€€€€€€€€€€ ‰ÁÉ½™¥±•}™¥•±‘}ÉÕ±•Ìˆ°¹½Éµ…±¥é”¤°4(€€€€€€€€€€€€€€€€€€€€€€€€ ‰…ÁÁÉ½Ù•‘}ÍÑÉÕÑÕÉ•}µ…Àˆ°‰½½°¡ÍÑÉÕÑÕÉ•}µ…À¤¤°4(€€€€€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€€€€€€€€€¥˜•¹…‰±•4(€€€€€€€€€€€€€€€t°4(€€€€€€€€€€€ô°4(€€€€€€€€€€€€‰ÁÉ½Ñ•Ñ•‘}½‰©•Ñ}¥¹Ñ•É¥Ñäˆèì4(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè½‰©•ÑÍ}½¬°4(€€€€€€€€€€€€€€€€‰½É¥¥¹…°ˆè½É¥¥¹…±}½‰©•ÑÌ°4(€€€€€€€€€€€€€€€€‰™½Éµ…ÑÑ•ˆè™½Éµ…ÑÑ•‘}½‰©•ÑÌ°4(€€€€€€€€€€€€€€€€‰Á…å±½…‘}½µÁ…É¥Í½¸ˆèì4(€€€€€€€€€€€€€€€€€€€€‰Á…Ñ¡}¥¹‘•Á•¹‘•¹ÐˆèQÉÕ”°4(€€€€€€€€€€€€€€€€€€€€‰½É¥¥¹…°ˆè½É¥¥¹…±}Á…å±½…‘Ì°4(€€€€€€€€€€€€€€€€€€€€‰™½Éµ…ÑÑ•ˆè™½Éµ…ÑÑ•‘}Á…å±½…‘Ì°4(€€€€€€€€€€€€€€€ô°4(€€€€€€€€€€€ô°4(€€€€€€€€€€€€‰•ÅÕ…Ñ¥½¹Ìˆè•ÅÕ…Ñ¥½¹}¥¹Ù•¹Ñ½Éä¡…ÉÌ¹™½Éµ…ÑÑ•¤°4(€€€€€€€€€€€€‰…ÁÁÉ½Ù•‘}µ…¹Õ…±}¥‘•¹Ñ¥™¥•É}É•Á±…•µ•¹ÑÌˆè…ÁÑ¥½¹}É•Á±…•µ•¹ÑÌ°4(€€€€€€€€€€€€‰Á…¥¹…Ñ¥½¸ˆèì(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè¹½ÐÁ…¥¹…Ñ¥½¹}™…¥±ÕÉ•Ì°(€€€€€€€€€€€€€€€€‰™…¥±ÕÉ•ÌˆèÁ…¥¹…Ñ¥½¹}™…¥±ÕÉ•Ì°(€€€€€€€€€€€€€€€€‰¥¹Ù•¹Ñ½ÉäˆèÁ…¥¹…Ñ¥½¸°(€€€€€€€€€€€ô°(€€€€€€€€€€€€‰ÍÑÉÕÑÕÉ•}¡•…‘¥¹}½Á•É…Ñ¥½¹Ìˆèì(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè¹½ÐÍÑÉÕÑÕÉ•}¡•…‘¥¹}™…¥±ÕÉ•Ì°(€€€€€€€€€€€€€€€€‰™…¥±ÕÉ•ÌˆèÍÑÉÕÑÕÉ•}¡•…‘¥¹}™…¥±ÕÉ•Ì°(€€€€€€€€€€€ô°(€€€€€€€€€€€€‰ÍÑÉÕÑÕÉ•}Ñ…‰±•}½Á•É…Ñ¥½¹Ìˆèì(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè¹½ÐÍÑÉÕÑÕÉ•}Ñ…‰±•}™…¥±ÕÉ•Ì°4(€€€€€€€€€€€€€€€€‰™…¥±ÕÉ•ÌˆèÍÑÉÕÑÕÉ•}Ñ…‰±•}™…¥±ÕÉ•Ì°4(€€€€€€€€€€€ô°4(€€€€€€€€€€€€‰ÍÑÉÕÑÕÉ•}¥µ…•}½Á•É…Ñ¥½¹Ìˆèì4(€€€€€€€€€€€€€€€€‰Á…ÍÍ•ˆè¹½ÐÍÑÉÕÑÕÉ•}¥µ…•}™…¥±ÕÉ•Ì°4(€€€€€€€€€€€€€€€€‰™…¥±ÕÉ•ÌˆèÍÑÉÕÑÕÉ•}¥µ…•}™…¥±ÕÉ•Ì°4(€€€€€€€€€€€ô°4(€€€€€€€€€€€€‰ÉÕ±•ÌˆèÉÕ±•}É•ÍÕ±ÑÌ°4(€€€€€€€ô4(€€€€€€€¥˜…ÉÌ¹½ÕÑÁÕÐè4(€€€€€€€€€€€ÝÉ¥Ñ•}©Í½¸¡…ÉÌ¹½ÕÑÁÕÐ°É•ÍÕ±Ð¤4(€€€€€€€ÁÉ¥¹Ð¡©Í½¸¹‘ÕµÁÌ¡É•ÍÕ±Ð°•¹ÍÕÉ•}…Í¥¤õ…±Í”°¥¹‘•¹ÐôÈ¤¤4(€€€€€€€É•ÑÕÉ¸€À¥˜É•ÍÕ±Ñl‰Á…ÍÍ•‰t•±Í”€Ä4(€€€•á•ÁÐ€¡½Éµ…Ñ5½¹½É…Á¡ÉÉ½È°=MÉÉ½È°Y…±Õ•ÉÉ½È¤…Ì•áŒè4(€€€€€€€ÁÉ¥¹Ð¡ÍÑÈ¡•áŒ¤°™¥±”õÍåÌ¹ÍÑ‘•ÉÈ¤4(€€€€€€€É•ÑÕÉ¸€Ä4(4(4)¥˜}}¹…µ•}|€ôô€‰}}µ…¥¹}|ˆè4(€€€É…¥Í”MåÍÑ•µá¥Ð¡µ…¥¸ ¤¤4(