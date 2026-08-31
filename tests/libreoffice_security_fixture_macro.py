#!/usr/bin/env python3
"""Create a synthetic ODT whose OnLoad event points at a document macro."""

from __future__ import annotations

import json
import os
import traceback
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


OUTPUT_ENV = "FORMAT_MONOGRAPH_SECURITY_ODT"
MARKER_ENV = "FORMAT_MONOGRAPH_SECURITY_MACRO_MARKER"
RESULT_ENV = "FORMAT_MONOGRAPH_SECURITY_FIXTURE_RESULT"


def property_value(name: str, value: object) -> PropertyValue:
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def write_result(path: Path, payload: dict[str, object]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload), encoding="utf-8")
    os.replace(temporary, path)


def create_fixture(*_args: object) -> bool:
    output = Path(os.environ[OUTPUT_ENV]).resolve()
    marker = Path(os.environ[MARKER_ENV]).resolve()
    result = Path(os.environ[RESULT_ENV]).resolve()
    document = None
    payload: dict[str, object]
    try:
        document = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(  # type: ignore[name-defined]
            "private:factory/swriter",
            "_blank",
            0,
            (property_value("Hidden", True),),
        )
        if document is None:
            raise RuntimeError("LibreOffice could not create the security fixture.")
        libraries = document.BasicLibraries
        if not libraries.hasByName("Standard"):
            libraries.createLibrary("Standard")
        library = libraries.getByName("Standard")
        module_source = (
            "Sub OnLoad\n"
            "  Dim channel As Integer\n"
            "  channel = FreeFile\n"
            f'  Open "{str(marker).replace(chr(34), chr(34) * 2)}" For Output As #channel\n'
            '  Print #channel, "EXECUTED"\n'
            "  Close #channel\n"
            "End Sub\n"
        )
        if library.hasByName("SecurityProbe"):
            library.replaceByName("SecurityProbe", module_source)
        else:
            library.insertByName("SecurityProbe", module_source)
        document.Text.String = "Synthetic document-macro security fixture."
        output.parent.mkdir(parents=True, exist_ok=True)
        document.storeAsURL(
            uno.systemPathToFileUrl(str(output)),
            (
                property_value("FilterName", "writer8"),
                property_value("Overwrite", True),
            ),
        )
        document.close(True)
        document = None
        payload = {"ok": True, "fixture": str(output)}
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


g_exportedScripts = (create_fixture,)
