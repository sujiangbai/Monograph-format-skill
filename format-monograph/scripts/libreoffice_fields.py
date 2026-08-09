#!/usr/bin/env python3
"""Refresh DOCX indexes and fields through a running LibreOffice UNO service."""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


def property_value(name: str, value: object) -> PropertyValue:
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def connect(port: int):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    last_error = None
    for _ in range(60):
        try:
            return resolver.resolve(
                f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
            )
        except Exception as exc:  # UNO exposes implementation-specific exceptions.
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"Could not connect to LibreOffice UNO: {last_error}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--port", required=True, type=int)
    args = parser.parse_args()

    context = connect(args.port)
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    document = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(args.input.resolve())),
        "_blank",
        0,
        (
            property_value("Hidden", True),
            property_value("ReadOnly", False),
            property_value("UpdateDocMode", 3),
        ),
    )
    if document is None:
        raise RuntimeError("LibreOffice could not open the DOCX.")
    try:
        indexes = document.getDocumentIndexes()
        for index in range(indexes.getCount()):
            indexes.getByIndex(index).update()
        fields = document.getTextFields()
        if hasattr(fields, "refresh"):
            fields.refresh()
        if hasattr(document, "calculateAll"):
            document.calculateAll()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        document.storeAsURL(
            uno.systemPathToFileUrl(str(args.output.resolve())),
            (
                property_value("FilterName", "Office Open XML Text"),
                property_value("Overwrite", True),
            ),
        )
        print(
            json.dumps(
                {"indexes_updated": indexes.getCount(), "fields_refreshed": True}
            )
        )
    finally:
        document.close(True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
