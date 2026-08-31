#!/usr/bin/env python3
"""Generate a one-TOC DOCX using LibreOffice for round-trip calibration."""

from __future__ import annotations

import json
import os
import traceback
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


OUTPUT_ENV = "FORMAT_MONOGRAPH_TOC_FIXTURE_OUTPUT"
RESULT_ENV = "FORMAT_MONOGRAPH_TOC_FIXTURE_RESULT"


def property_value(name: str, value: object) -> PropertyValue:
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def write_result(path: Path, payload: dict[str, object]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload), encoding="utf-8")
    os.replace(temporary, path)


def create_toc_fixture(*_args: object) -> bool:
    output = Path(os.environ[OUTPUT_ENV]).resolve()
    result = Path(os.environ[RESULT_ENV]).resolve()
    document = None
    payload: dict[str, object]
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()  # type: ignore[name-defined]
        document = desktop.loadComponentFromURL(
            "private:factory/swriter",
            "_blank",
            0,
            (property_value("Hidden", True),),
        )
        text = document.Text
        cursor = text.createTextCursor()
        cursor.ParaStyleName = "Heading 1"
        text.insertString(cursor, "Synthetic heading", False)
        text.insertControlCharacter(cursor, 0, False)
        cursor.ParaStyleName = "Standard"
        index = document.createInstance("com.sun.star.text.ContentIndex")
        index.CreateFromOutline = True
        index.CreateFromMarks = False
        index.Level = 3
        text.insertTextContent(cursor, index, False)
        index.update()
        output.parent.mkdir(parents=True, exist_ok=True)
        document.storeAsURL(
            uno.systemPathToFileUrl(str(output)),
            (
                property_value("FilterName", "Office Open XML Text"),
                property_value("Overwrite", True),
            ),
        )
        document.close(True)
        document = None
        payload = {"ok": True, "document_index_count": 1}
    except Exception as exc:
        payload = {
            "ok": False,
            "error": str(exc),
            "traceback": traceback.format_exc(),
        }
    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                pass
    write_result(result, payload)
    return bool(payload["ok"])


g_exportedScripts = (create_toc_fixture,)
