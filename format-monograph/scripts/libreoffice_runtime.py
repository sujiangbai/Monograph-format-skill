#!/usr/bin/env python3
"""Resolve LibreOffice runtime capabilities without executing untrusted documents."""

from __future__ import annotations

import os
import platform
import re
from pathlib import Path


WRAPPED_MACOS_SOFFICE = re.compile(
    r"[\"'](?P<path>/[^\"']*/LibreOffice\.app/Contents/MacOS/soffice)[\"']"
)


def default_macos_soffice(*, system: str | None = None) -> str | None:
    if (system or platform.system()) != "Darwin":
        return None
    for candidate in (
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate.resolve())
    return None


def _supports_internal_python_macro(candidate: Path) -> bool:
    candidate = candidate.expanduser().resolve()
    if (
        not candidate.is_file()
        or not os.access(candidate, os.X_OK)
        or candidate.name != "soffice"
    ):
        return False
    try:
        contents = candidate.parents[1]
    except IndexError:
        return False
    return bool(
        candidate.parent.name == "MacOS"
        and contents.name == "Contents"
        and (contents / "Resources" / "Scripts" / "python").is_dir()
        and (contents / "Frameworks" / "LibreOfficePython.framework").exists()
    )


def macos_internal_macro_soffice(
    soffice: str | None, *, system: str | None = None
) -> str | None:
    """Return the app-bundle executable that can host embedded Python macros."""
    if not soffice or (system or platform.system()) != "Darwin":
        return None
    requested = Path(soffice).expanduser().resolve()
    if not requested.is_file() or not os.access(requested, os.X_OK):
        return None
    if _supports_internal_python_macro(requested):
        return str(requested)
    try:
        if requested.stat().st_size > 8192:
            return None
        wrapper = requested.read_text(encoding="utf-8")
    except (OSError, UnicodeError):
        return None
    match = WRAPPED_MACOS_SOFFICE.search(wrapper)
    if not match:
        return None
    target = Path(match.group("path"))
    return str(target.resolve()) if _supports_internal_python_macro(target) else None
