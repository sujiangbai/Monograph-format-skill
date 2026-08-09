#!/usr/bin/env python3
"""Report portable execution capabilities for format-monograph."""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import platform
import shutil
import sys
from pathlib import Path

PACKAGES = {
    "python-docx": "docx",
    "jsonschema": "jsonschema",
    "lxml": "lxml",
    "PyMuPDF": "fitz",
}


def package_status() -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    for distribution, module in PACKAGES.items():
        try:
            version = importlib.metadata.version(distribution)
            available = True
        except importlib.metadata.PackageNotFoundError:
            version = None
            available = False
        result[distribution] = {
            "module": module,
            "available": available,
            "version": version,
        }
    return result


def font_directories() -> list[str]:
    candidates: list[Path] = []
    system = platform.system()
    if system == "Windows":
        windir = os.environ.get("WINDIR")
        if windir:
            candidates.append(Path(windir) / "Fonts")
        local_app = os.environ.get("LOCALAPPDATA")
        if local_app:
            candidates.append(Path(local_app) / "Microsoft" / "Windows" / "Fonts")
    elif system == "Darwin":
        candidates.extend(
            [Path("/System/Library/Fonts"), Path("/Library/Fonts"), Path.home() / "Library/Fonts"]
        )
    else:
        candidates.extend(
            [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".fonts"]
        )
    return [str(path) for path in candidates if path.is_dir()]


def resolve_renderer(requested: str | None) -> tuple[str | None, str | None]:
    if requested:
        path = Path(requested).expanduser()
        return (str(path.resolve()), "argument") if path.is_file() else (None, "argument")
    configured = os.environ.get("FORMAT_MONOGRAPH_RENDERER")
    if configured:
        path = Path(configured).expanduser()
        return (str(path.resolve()), "environment") if path.is_file() else (None, "environment")
    discovered = shutil.which("soffice") or shutil.which("libreoffice")
    return discovered, "path" if discovered else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON.")
    parser.add_argument("--renderer", help="Explicit LibreOffice-compatible renderer path.")
    args = parser.parse_args()

    packages = package_status()
    python_ok = sys.version_info >= (3, 11)
    inspection_ok = python_ok and all(
        packages[name]["available"] for name in ("python-docx", "lxml")
    )
    validation_ok = python_ok and bool(packages["jsonschema"]["available"])
    editing_ok = inspection_ok
    soffice, renderer_source = resolve_renderer(args.renderer)
    pymupdf_ok = bool(packages["PyMuPDF"]["available"])

    if editing_ok and validation_ok and soffice and pymupdf_ok:
        mode = "full"
    elif editing_ok and validation_ok:
        mode = "structural"
    else:
        mode = "analysis"

    result = {
        "mode": mode,
        "python": {
            "available": True,
            "executable": sys.executable,
            "version": platform.python_version(),
            "supported": python_ok,
        },
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
        },
        "packages": packages,
        "capabilities": {
            "inspection": inspection_ok,
            "profile_validation": validation_ok,
            "docx_editing": editing_ok,
            "field_finalization": bool(editing_ok and soffice),
            "rendering": bool(soffice and pymupdf_ok),
        },
        "rendering": {
            "soffice": soffice,
            "renderer": soffice,
            "source": renderer_source,
            "pymupdf": pymupdf_ok,
            "available": bool(soffice and pymupdf_ok),
            "field_refresh_candidate": bool(soffice),
        },
        "font_directories": font_directories(),
        "limitations": [],
    }
    if not python_ok:
        result["limitations"].append("Python 3.11 or newer is required.")
    if not inspection_ok:
        result["limitations"].append("DOCX inspection dependencies are incomplete.")
    if not validation_ok:
        result["limitations"].append("Profile validation dependency jsonschema was not found.")
    if not soffice:
        result["limitations"].append("LibreOffice soffice was not found.")
    if not pymupdf_ok:
        result["limitations"].append("PyMuPDF was not found.")

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Capability mode: {mode}")
        for limitation in result["limitations"]:
            print(f"- {limitation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
