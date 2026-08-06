#!/usr/bin/env python3
"""Audit content preservation and automatically verifiable profile rules."""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

from _common import (
    FormatMonographError,
    content_fingerprint,
    load_document,
    style_name_for_selector,
    write_json,
)
from validate_profile import validate


def close(actual: Any, expected: Any, tolerance: float = 0.05) -> bool:
    if actual is None:
        return expected is None
    return math.isclose(float(actual), float(expected), abs_tol=tolerance)


def style_value(style: Any, key: str) -> Any:
    font, pf = style.font, style.paragraph_format
    if key == "font_name":
        return font.name
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
    if key == "line_spacing":
        value = pf.line_spacing
        return float(value) if isinstance(value, (int, float)) else None
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
    if key.endswith("_pt") or key == "line_spacing":
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


def audit_section_rule(document: Any, rule: dict) -> list[dict]:
    failures = []
    for index, section in enumerate(document.sections):
        values = {
            "page_width_mm": section.page_width.mm,
            "page_height_mm": section.page_height.mm,
            "margin_top_mm": section.top_margin.mm,
            "margin_bottom_mm": section.bottom_margin.mm,
            "margin_left_mm": section.left_margin.mm,
            "margin_right_mm": section.right_margin.mm,
            "gutter_mm": section.gutter.mm,
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


def audit_table_rule(document: Any, rule: dict) -> list[dict]:
    failures = []
    for index, table in enumerate(document.tables):
        for key, expected in rule["properties"].items():
            if key == "table_style":
                actual = table.style.name if table.style else None
            elif key == "alignment":
                actual = normalized_alignment(str(table.alignment))
            elif key == "repeat_header_row":
                if not table.rows:
                    actual = False
                else:
                    element = table.rows[0]._tr.trPr
                    header = None if element is None else element.find(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblHeader"
                    )
                    actual = header is not None
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
        original_fp = content_fingerprint(args.original)
        formatted_fp = content_fingerprint(args.formatted)
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
            "passed": content_ok and rules_ok,
            "content_integrity": {
                "passed": content_ok,
                "original_sha256": original_fp,
                "formatted_sha256": formatted_fp,
                "field_results_excluded": True,
            },
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
