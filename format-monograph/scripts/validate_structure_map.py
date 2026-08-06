#!/usr/bin/env python3
"""Validate an approved structure map and optionally bind it to a DOCX."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from _common import FormatMonographError
from structure_map import load_structure_map, validate_structure_map_source


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("structure_map", type=Path)
    parser.add_argument("--source", type=Path)
    args = parser.parse_args()
    try:
        value = load_structure_map(args.structure_map)
        if args.source:
            validate_structure_map_source(args.source, value)
        print(
            json.dumps(
                {
                    "valid": True,
                    "status": value["status"],
                    "source_bound": bool(args.source),
                },
                ensure_ascii=False,
            )
        )
        return 0
    except FormatMonographError as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
