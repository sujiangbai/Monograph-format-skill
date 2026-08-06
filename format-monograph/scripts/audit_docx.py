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
        value = point_values[key]
        return None if value is None else value.pt
    if key == "first_line_indent_chars":
        p_pr = style.element.pPr
        ind = None if p_pr is None else p_pr.find(qn("w:ind"))
        value = None if ind is None else ind.get(qn("w:firstLineChars"))
        return None if value is None else int(value) / 100
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
        return getattr(pf, key)
    return "<unsupported>"


def normalized_alignment(value: str) -> str:
    value = value.lower()
    for name in ("left", "center", "right", "justify", "distributed"):
        if name in value:
            return name
    return value


def compare_value(key: str, actual: Any, expected: Any) -> bool:
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
        actual = style_value(style, key)
        if not compare_value(key, actual, expected):
            failures.append({"property": key, "expected": expected, "actual": actual})
    return failures


def document_toggle(document: Any, name: str) -> bool:
    return document.settings.element.find(qn(f"w:{name}")) is not None


def audit_section_rule(document: Any, rule: dict) -> list[dict]:
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
            "margin_inner_ratio": section.left_margin.mm / width,
            "margin_outer_ratio": section.right_margin.mm / width,
            "margin_top_ratio": section.top_margin.mm / height,
            "margin_bottom_ratio": section.bottom_margin.mm / height,
            "mirror_margins": document_toggle(document, "mirrorMargins"),
            "page_size_policy": "preserve",
            "different_first_page_header_footer": bool(section.different_first_page_header_footer),
            "odd_and_even_pages_header_footer": bool(
                document.settings.odd_and_even_pages_header_footer
            ),
            "orientation": "landscape" if section.page_width > section.page_height else "portrait",
        }
        for key, expected in rule["properties"].items():
            actual = values.get(key, "<unsupported>")
            if not compare_value(key, actual, expected):
                failures.append(
                    {
                        "section": index,
                        "property": key,
                        "expected": expected,
                        "actual": actual,
                    }
                )
    return failures


def row_property(row: Any, name: str) -> bool:
    tr_pr = row._tr.trPr
    return tr_pr is not None and tr_pr.find(qn(f"w:{name}")) is not None


def audit_table_rule(document: Any, rule: dict) -> list[dict]:
    failures = []
    for index, table in enumerate(document.tables):
        for key, expected in rule["properties"].items():
            if key == "table_style":
                actual = table.style.name if table.style else None
            elif key == "alignment":
                actual = normalized_alignment(str(table.alignment))
            elif key == "repeat_header_row":
                actual = bool(table.rows and row_property(table.rows[0], "tblHeader"))
            elif key == "prevent_row_split":
                actual = all(row_property(row, "cantSplit") for row in table.rows)
            else:
                actual = "<unsupported>"
            if not compare_value(key, actual, expected):
                failures.append(
                    {
                        "table": index,
                        "property": key,
                        "expected": expected,
                        "actual": actual,
                    }
                )
    return failures


def audit_field_rule(document: Any, rule: dict) -> list[dict]:
    failures = []
    properties = rule["properties"]
    if properties.get("update_on_open") and not document_toggle(document, "updateFields"):
        failures.append({"property": "update_on_open", "expected": True, "actual": False})
    if properties.get("convert_explicit_markers"):
        remaining = [
            paragraph.text
            for paragraph in document.paragraphs
            if _field_instruction_for_marker(paragraph.text.strip()) is not None
        ]
        if remaining:
            failures.append(
                {
                    "property": "convert_explicit_markers",
                    "expected": "no markers",
                    "actual": remaining[:10],
                }
            )
    levels = int(properties.get("heading_levels", 4))
    if properties.get("rebuild_heading_numbering"):
        for level in range(1, levels + 1):
            style = document.styles[f"Heading {level}"]
            p_pr = style.element.pPr
            if p_pr is None or p_pr.find(qn("w:numPr")) is None:
                failures.append(
                    {
                        "property": "rebuild_heading_numbering",
                        "expected": f"numbering on Heading {level}",
                        "actual": "missing",
                    }
                )
    if properties.get("strip_manual_heading_prefixes"):
        for paragraph in document.paragraphs:
            if not paragraph.style or not paragraph.style.name.startswith("Heading "):
                continue
            try:
                level = int(paragraph.style.name.split()[-1])
            except ValueError:
                continue
            if level <= levels and _heading_prefix_pattern(level).match(paragraph.text):
                failures.append(
                    {
                        "property": "strip_manual_heading_prefixes",
                        "expected": "no manual prefix",
                        "actual": paragraph.text[:80],
                    }
                )
    return failures


def audit_equation_rule(path: Path, rule: dict) -> list[dict]:
    values = equation_inventory(path)
    failures = []
    properties = rule["properties"]
    if properties.get("block_formula_images") and values["formula_image_candidates"]:
        failures.append(
            {
                "property": "block_formula_images",
                "expected": 0,
                "actual": values["formula_image_candidates"],
            }
        )
    return failures


def uses_derived_normalization(profile: dict) -> bool:
    return any(
        rule.get("status") == "approved"
        and rule.get("application") == "automatic"
        and rule.get("selector", {}).get("kind") == "field_role"
        and any(
            rule.get("properties", {}).get(key)
            for key in (
                "convert_explicit_markers",
                "rebuild_heading_numbering",
                "strip_manual_heading_prefixes",
            )
        )
        for rule in profile.get("rules", [])
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("original", type=Path)
    parser.add_argument("formatted", type=Path)
    parser.add_argument("--profile", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    try:
        profile_errors, profile = validate(args.profile)
        if profile_errors:
            raise FormatMonographError("Profile validation failed: " + "; ".join(profile_errors))
        normalize = uses_derived_normalization(profile)
        original_fp = content_fingerprint(args.original, normalize_derived=normalize)
        formatted_fp = content_fingerprint(args.formatted, normalize_derived=normalize)
        original_objects = protected_object_manifest(args.original)
        formatted_objects = protected_object_manifest(args.formatted)
        objects_ok = original_objects == formatted_objects
        document = load_document(args.formatted)
        rule_results = []

        for rule in profile["rules"]:
            if rule["status"] != "approved":
                continue
            if rule["application"] == "manual_review":
                rule_results.append(
                    {"id": rule["id"], "status": "manual_review", "failures": []}
                )
                continue
            kind = rule["selector"]["kind"]
            if kind in {"document", "section_role"}:
                failures = audit_section_rule(document, rule)
            elif kind == "table_role":
                failures = audit_table_rule(document, rule)
            elif kind == "field_role":
                failures = audit_field_rule(document, rule)
            elif kind == "equation_role":
                failures = audit_equation_rule(args.formatted, rule)
            else:
                failures = audit_style_rule(document, rule)
            rule_results.append(
                {
                    "id": rule["id"],
                    "status": "pass" if not failures else "fail",
                    "failures": failures,
                }
            )

        content_ok = original_fp == formatted_fp
        rules_ok = all(item["status"] != "fail" for item in rule_results)
        result = {
            "passed": content_ok and objects_ok and rules_ok,
            "content_integrity": {
                "passed": content_ok,
                "original_sha256": original_fp,
                "formatted_sha256": formatted_fp,
                "field_results_and_approved_derived_values_excluded": normalize,
            },
            "protected_object_integrity": {
                "passed": objects_ok,
                "original": original_objects,
                "formatted": formatted_objects,
            },
            "equations": equation_inventory(args.formatted),
            "rules": rule_results,
        }
        if args.output:
            write_json(args.output, result)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["passed"] else 1
    except (FormatMonographError, OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
