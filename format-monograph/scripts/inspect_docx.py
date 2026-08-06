#!/usr/bin/env python3
"""Create a structural and editable-object inventory for a monograph DOCX."""

from __future__ import annotations

import argparse
import collections
import json
import sys
from pathlib import Path

from docx.oxml.ns import qn

from _common import (
    FormatMonographError,
    content_fingerprint,
    equation_inventory,
    field_inventory,
    load_document,
    protected_object_manifest,
    word_xml_counts,
    write_json,
)


def length_points(value) -> float | None:
    return None if value is None else round(value.pt, 3)


def length_mm(value) -> float | None:
    return None if value is None else round(value.mm, 3)


def style_definition(style) -> dict:
    r_pr = style.element.rPr
    r_fonts = None if r_pr is None else r_pr.rFonts
    attrs = {}
    if r_fonts is not None:
        for name in ("ascii", "hAnsi", "eastAsia", "cs"):
            attrs[name] = r_fonts.get(qn(f"w:{name}"))
    spacing = style.paragraph_format.line_spacing
    return {
        "font_name": style.font.name,
        "font_name_ascii": attrs.get("ascii") or attrs.get("hAnsi"),
        "font_name_east_asia": attrs.get("eastAsia"),
        "font_name_complex_script": attrs.get("cs"),
        "font_size_pt": length_points(style.font.size),
        "bold": style.font.bold,
        "line_spacing": (
            float(spacing) if isinstance(spacing, (int, float)) else length_points(spacing)
        ),
        "line_spacing_rule": (
            None
            if style.paragraph_format.line_spacing_rule is None
            else str(style.paragraph_format.line_spacing_rule)
        ),
    }


def row_has_property(row, name: str) -> bool:
    tr_pr = row._tr.trPr
    return tr_pr is not None and tr_pr.find(qn(f"w:{name}")) is not None


def inventory(path: Path) -> dict:
    document = load_document(path)
    style_counts = collections.Counter(
        paragraph.style.name if paragraph.style else "<none>"
        for paragraph in document.paragraphs
    )
    headings = [
        {
            "index": index,
            "style": paragraph.style.name,
            "text": paragraph.text[:160],
        }
        for index, paragraph in enumerate(document.paragraphs)
        if paragraph.style and paragraph.style.name.startswith("Heading")
    ]
    sections = []
    for index, section in enumerate(document.sections):
        sections.append(
            {
                "index": index,
                "orientation": str(section.orientation),
                "page_width_mm": length_mm(section.page_width),
                "page_height_mm": length_mm(section.page_height),
                "margin_top_mm": length_mm(section.top_margin),
                "margin_bottom_mm": length_mm(section.bottom_margin),
                "margin_left_mm": length_mm(section.left_margin),
                "margin_right_mm": length_mm(section.right_margin),
                "gutter_mm": length_mm(section.gutter),
                "header_distance_mm": length_mm(section.header_distance),
                "footer_distance_mm": length_mm(section.footer_distance),
                "different_first_page_header_footer": bool(
                    section.different_first_page_header_footer
                ),
            }
        )

    tables = []
    for index, table in enumerate(document.tables):
        tables.append(
            {
                "index": index,
                "rows": len(table.rows),
                "columns": len(table.columns),
                "style": table.style.name if table.style else None,
                "repeat_header_row": bool(
                    table.rows and row_has_property(table.rows[0], "tblHeader")
                ),
                "prevent_split_rows": [
                    row_index
                    for row_index, row in enumerate(table.rows)
                    if row_has_property(row, "cantSplit")
                ],
            }
        )

    settings = document.settings.element
    object_manifest = protected_object_manifest(path)
    return {
        "file": str(path.resolve()),
        "content_fingerprint_sha256": content_fingerprint(path),
        "paragraph_count": len(document.paragraphs),
        "nonempty_paragraph_count": sum(bool(p.text.strip()) for p in document.paragraphs),
        "style_counts": dict(sorted(style_counts.items())),
        "style_definitions": {
            name: style_definition(document.styles[name])
            for name in ("Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Caption")
            if name in document.styles
        },
        "headings": headings,
        "section_count": len(document.sections),
        "sections": sections,
        "table_count": len(document.tables),
        "tables": tables,
        "inline_shape_count": len(document.inline_shapes),
        "settings": {
            "odd_and_even_pages_header_footer": bool(
                document.settings.odd_and_even_pages_header_footer
            ),
            "mirror_margins": settings.find(qn("w:mirrorMargins")) is not None,
            "update_fields_on_open": settings.find(qn("w:updateFields")) is not None,
        },
        "fields": field_inventory(path),
        "equations": equation_inventory(path),
        "protected_objects": {
            "embedding_count": len(object_manifest["embeddings"]),
            "media_count": len(object_manifest["media"]),
            "omml_hash_count": len(object_manifest["omml_sha256"]),
            "manifest": object_manifest,
        },
        "ooxml": word_xml_counts(path),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    try:
        result = inventory(args.input)
        write_json(args.output, result)
    except FormatMonographError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(
        json.dumps(
            {"output": str(args.output), "fingerprint": result["content_fingerprint_sha256"]},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
