#!/usr/bin/env python3
"""Create a structural inventory for a monograph DOCX."""

from __future__ import annotations

import argparse
import collections
import json
import sys
from pathlib import Path

from _common import (
    FormatMonographError,
    content_fingerprint,
    load_document,
    word_xml_counts,
    write_json,
)


def length_points(value) -> float | None:
    return None if value is None else round(value.pt, 3)


def length_mm(value) -> float | None:
    return None if value is None else round(value.mm, 3)


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
            }
        )

    return {
        "file": str(path.resolve()),
        "content_fingerprint_sha256": content_fingerprint(path),
        "paragraph_count": len(document.paragraphs),
        "nonempty_paragraph_count": sum(bool(p.text.strip()) for p in document.paragraphs),
        "style_counts": dict(sorted(style_counts.items())),
        "headings": headings,
        "section_count": len(document.sections),
        "sections": sections,
        "table_count": len(document.tables),
        "tables": tables,
        "inline_shape_count": len(document.inline_shapes),
        "settings": {
            "odd_and_even_pages_header_footer": bool(
                document.settings.odd_and_even_pages_header_footer
            )
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
    print(json.dumps({"output": str(args.output), "fingerprint": result["content_fingerprint_sha256"]}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
