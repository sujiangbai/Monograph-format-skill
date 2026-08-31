#!/usr/bin/env python3
"""Unsafe positive control for a synthetic loopback-only external-field fixture."""

from __future__ import annotations

import json
import os
import traceback
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
from com.sun.star.document.UpdateDocMode import FULL_UPDATE


INPUT_ENV = "FORMAT_MONOGRAPH_PROBE_INPUT"
RESULT_ENV = "FORMAT_MONOGRAPH_PROBE_RESULT"
EXPECTED_URL_ENV = "FORMAT_MONOGRAPH_PROBE_EXPECTED_URL"


def property_value(name: str, value: object) -> PropertyValue:
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def write_result(path: Path, payload: dict[str, object]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload), encoding="utf-8")
    os.replace(temporary, path)


def external_graphic_urls(document: object) -> list[str]:
    urls = []
    graphics = document.getGraphicObjects()
    for name in graphics.getElementNames():
        graphic_object = graphics.getByName(name)
        candidates = []
        for property_name in ("GraphicURL", "LinkDisplayName"):
            try:
                candidates.append(getattr(graphic_object, property_name))
            except Exception:
                pass
        for candidate in candidates:
            if (
                isinstance(candidate, str)
                and candidate.startswith(("http://", "https://"))
                and candidate not in urls
            ):
                urls.append(candidate)
    return urls


def probe_external_refresh(*_args: object) -> bool:
    result_path = Path(os.environ[RESULT_ENV])
    document = None
    payload: dict[str, object]
    try:
        input_path = Path(os.environ[INPUT_ENV]).resolve()
        document = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(  # type: ignore[name-defined]
            uno.systemPathToFileUrl(str(input_path)),
            "_blank",
            0,
            (
                property_value("Hidden", True),
                property_value("ReadOnly", False),
                property_value("UpdateDocMode", FULL_UPDATE),
                property_value("MacroExecutionMode", NEVER_EXECUTE),
            ),
        )
        if document is None:
            raise RuntimeError("LibreOffice could not open the external probe fixture.")
        expected_url = os.environ[EXPECTED_URL_ENV]
        if not expected_url.startswith("http://127.0.0.1:"):
            raise RuntimeError("External positive control URL is not loopback-only.")
        urls = external_graphic_urls(document)
        graphics = document.getGraphicObjects()
        refreshed = 0
        for name in graphics.getElementNames():
            graphic_object = graphics.getByName(name)
            try:
                graphic_object.Graphic
            except Exception as exc:
                raise RuntimeError(
                    f"Unsafe positive control could not load graphic {name}: {exc}"
                ) from exc
            refreshed += 1
        payload = {
            "ok": True,
            "expected_external_url": expected_url,
            "recognized_external_connections": urls,
            "recognized_external_connection_count": len(urls),
            "unsafe_full_update_load_completed": True,
            "unsafe_graphics_loaded": refreshed,
        }
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
            except Exception as exc:
                payload = {
                    "ok": False,
                    "error": f"LibreOffice external probe close failed: {exc}",
                    "traceback": traceback.format_exc(),
                }
    write_result(result_path, payload)
    return bool(payload["ok"])


g_exportedScripts = (probe_external_refresh,)
