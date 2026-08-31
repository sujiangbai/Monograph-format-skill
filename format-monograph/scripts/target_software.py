#!/usr/bin/env python3
"""Resolve allowlisted target-application names to stable workflow IDs."""

from __future__ import annotations

from typing import Final


MICROSOFT_WORD: Final = "microsoft_word"
LIBREOFFICE: Final = "libreoffice"
UNSUPPORTED: Final = "unsupported"


def _normalized_alias(value: object) -> str:
    return " ".join(str(value or "").split()).casefold()


_TARGET_ALIASES: Final[dict[str, str]] = {
    MICROSOFT_WORD: MICROSOFT_WORD,
    "microsoft word": MICROSOFT_WORD,
    "microsoft word 2016": MICROSOFT_WORD,
    "microsoft word 2019": MICROSOFT_WORD,
    "microsoft word 2021": MICROSOFT_WORD,
    "microsoft word for mac": MICROSOFT_WORD,
    "microsoft 365": MICROSOFT_WORD,
    LIBREOFFICE: LIBREOFFICE,
    "libreoffice writer": LIBREOFFICE,
    UNSUPPORTED: UNSUPPORTED,
}


def resolve_target_id(value: object) -> str:
    """Return one stable ID; arbitrary strings never become Microsoft Word."""
    return _TARGET_ALIASES.get(_normalized_alias(value), UNSUPPORTED)


def target_alias_is_supported(value: object) -> bool:
    return resolve_target_id(value) != UNSUPPORTED
