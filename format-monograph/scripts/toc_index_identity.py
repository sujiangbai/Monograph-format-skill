#!/usr/bin/env python3
"""Pure-stdlib canonical identity for approved OOXML/UNO content indexes."""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{WORD_NS}}}p"
W_FLD_CHAR = f"{{{WORD_NS}}}fldChar"
W_FLD_CHAR_TYPE = f"{{{WORD_NS}}}fldCharType"
W_INSTR_TEXT = f"{{{WORD_NS}}}instrText"
W_FLD_SIMPLE = f"{{{WORD_NS}}}fldSimple"
W_INSTR = f"{{{WORD_NS}}}instr"
CONTENT_INDEX_SERVICE = "com.sun.star.text.ContentIndex"
TOC_OUTLINE_INSTRUCTION = re.compile(
    r'^TOC \\o "(?P<first>[1-9])-(?P<last>[1-9])" \\h(?: \\z)?$'
)


def canonical_json_hash(value: Any) -> str:
    serialized = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def _normalize_instruction(value: str) -> str:
    return " ".join(value.split())


def ooxml_toc_identities(path: Path) -> list[dict[str, Any]]:
    """Read exact TOC instructions and stable main-story positions from OOXML."""
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    identities: list[dict[str, Any]] = []
    field_ordinal = 0
    toc_occurrence = 0
    for paragraph_ordinal, paragraph in enumerate(root.iter(W_P)):
        stack: list[dict[str, Any]] = []
        for element in paragraph.iter():
            if element.tag == W_FLD_SIMPLE:
                instruction = _normalize_instruction(element.get(W_INSTR, ""))
                current_ordinal = field_ordinal
                field_ordinal += 1
                if instruction.startswith("TOC "):
                    identities.append(
                        {
                            "part": "word/document.xml",
                            "paragraph_ordinal": paragraph_ordinal,
                            "field_ordinal": current_ordinal,
                            "occurrence": toc_occurrence,
                            "form": "simple",
                            "instruction": instruction,
                        }
                    )
                    toc_occurrence += 1
            elif element.tag == W_FLD_CHAR:
                marker = element.get(W_FLD_CHAR_TYPE)
                if marker == "begin":
                    stack.append(
                        {
                            "field_ordinal": field_ordinal,
                            "instruction_fragments": [],
                            "separated": False,
                        }
                    )
                    field_ordinal += 1
                elif marker == "separate" and stack:
                    frame = stack[-1]
                    if not frame["separated"]:
                        instruction = _normalize_instruction(
                            "".join(frame["instruction_fragments"])
                        )
                        if instruction.startswith("TOC "):
                            identities.append(
                                {
                                    "part": "word/document.xml",
                                    "paragraph_ordinal": paragraph_ordinal,
                                    "field_ordinal": frame["field_ordinal"],
                                    "occurrence": toc_occurrence,
                                    "form": "complex",
                                    "instruction": instruction,
                                }
                            )
                            toc_occurrence += 1
                        frame["separated"] = True
                elif marker == "end" and stack:
                    stack.pop()
            elif element.tag == W_INSTR_TEXT and stack and not stack[-1]["separated"]:
                stack[-1]["instruction_fragments"].append(element.text or "")
    return identities


def expected_uno_identity(instruction: str) -> dict[str, Any]:
    """Map the narrowly supported outline-only TOC instruction to UNO settings."""
    match = TOC_OUTLINE_INSTRUCTION.fullmatch(instruction)
    if match is None or int(match.group("first")) != 1:
        raise ValueError(
            "Approved TOC instruction has no stable outline-only UNO identity mapping."
        )
    return {
        "create_from_outline": True,
        "create_from_marks": False,
        "level": int(match.group("last")),
    }


def canonical_index_descriptor(
    path: Path,
    structure_contract_sha256: str,
    observed_uno: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    toc_fields = ooxml_toc_identities(path)
    if observed_uno is not None and len(observed_uno) != len(toc_fields):
        raise ValueError("UNO index count differs from the OOXML TOC count.")
    indexes = []
    for ordinal, field in enumerate(toc_fields):
        uno_identity = (
            observed_uno[ordinal]
            if observed_uno is not None
            else expected_uno_identity(field["instruction"])
        )
        indexes.append(
            {
                "ordinal": ordinal,
                "occurrence": field["occurrence"],
                "service": CONTENT_INDEX_SERVICE,
                "ooxml": field,
                "uno": uno_identity,
            }
        )
    return {
        "version": 2,
        "structure_contract_sha256": structure_contract_sha256,
        "expected_document_index_count": len(indexes),
        "indexes": indexes,
    }


def authorization_with_hash(descriptor: dict[str, Any]) -> dict[str, Any]:
    result = dict(descriptor)
    result["authorization_id"] = canonical_json_hash(descriptor)
    return result
