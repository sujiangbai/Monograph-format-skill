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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON.")
    args = parser.parse_args()

    packages = package_status()
    python_ok = sys.version_info >= (3, 11)
    docx_ok = python_ok and all(
        packages[name]["available"] for name in ("python-docx", "jsonschema", "lxml")
    )
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pymupdf_ok = bool(packages["PyMuPDF"]["available"])

    if docx_ok and soffice and pymupdf_ok:
        mode = "full"
    elif docx_ok:
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
        "rendering": {
            "soffice": soffice,
            "pymupdf": pymupdf_ok,
            "available": bool(soffice and pymupdf_ok),
        },
        "font_directories": font_directories(),
        "limitations": [],
    }
    if not python_ok:
        result["limitations"].append("Python 3.11 or newer is required.")
    if not docx_ok:
        result["limitations"].append("DOCX processing dependencies are incomplete.")
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
