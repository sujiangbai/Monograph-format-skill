#!/usr/bin/env python3
"""Disabled legacy LibreOffice UNO server helper.

Field refresh is permitted only through ``libreoffice_fields_macro.py`` hosted
by a verified macOS LibreOffice app bundle. This compatibility entry point is
kept solely to fail closed for stale callers.
"""

from __future__ import annotations

import sys


DISABLED_REASON = (
    "The legacy LibreOffice UNO server/helper backend is disabled; use the "
    "verified macOS internal-Python macro host or an approved deferred result."
)


def main() -> int:
    print(DISABLED_REASON, file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
