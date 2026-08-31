#!/usr/bin/env python3
"""Refresh DOCX fields inside LibreOffice's embedded Python runtime."""

from __future__ import annotations

import hashlib
import json
import os
import traceback
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
from com.sun.star.document.UpdateDocMode import NO_UPDATE

INPUT_ENV = "FORMAT_MONOGRAPH_FIELD_INPUT"
OUTPUT_ENV = "FORMAT_MONOGRAPH_FIELD_OUTPUT"
RESULT_ENV = "FORMAT_MONOGRAPH_FIELD_RESULT"
TOC_AUTHORIZATION_ENV = "FORMAT_MONOGRAPH_TOC_AUTHORIZATION"
TOC_CONTRACT_ENV = "FORMAT_MONOGRAPH_TOC_CONTRACT"
CONTENT_INDEX_SERVICE = "com.sun.star.text.ContentIndex"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{WORD_NS}}}p"
W_FLD_CHAR = f"{{{WORD_NS}}}fldChar"
W_FLD_CHAR_TYPE = f"{{{WORD_NS}}}fldCharType"
W_INSTR_TEXT = f"{{{WORD_NS}}}instrText"
W_FLD_SIMPLE = f"{{{WORD_NS}}}fldSimple"
W_INSTR = f"{{{WORD_NS}}}instr"


def canonical_json_hash(value: object) -> str:
    serialized = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def ooxml_toc_identities(path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    identities = []
    field_ordinal = 0
    toc_occurrence = 0
    for paragraph_ordinal, paragraph in enumerate(root.iter(W_P)):
        stack = []
        for element in paragraph.iter():
            if element.tag == W_FLD_SIMPLE:
                instruction = " ".join(element.get(W_INSTR, "").split())
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
                        instruction = " ".join(
                            "".join(frame["instruction_fragments"]).split()
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


def build_observed_index_descriptor(
    path: Path, structure_contract_sha256: str, observed_uno: list[dict[str, object]]
) -> dict[str, object]:
    toc_fields = ooxml_toc_identities(path)
    indexes = []
    for ordinal, field in enumerate(toc_fields):
        if ordinal >= len(observed_uno):
            raise RuntimeError("UNO index count differs from the OOXML TOC count.")
        indexes.append(
            {
                "ordinal": ordinal,
                "occurrence": field["occurrence"],
                "service": CONTENT_INDEX_SERVICE,
                "ooxml": field,
                "uno": observed_uno[ordinal],
            }
        )
    return {
        "version": 2,
        "structure_contract_sha256": structure_contract_sha256,
        "expected_document_index_count": len(indexes),
        "indexes": indexes,
    }


def property_value(name: str, value: object) -> PropertyValue:
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def write_result(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload), encoding="utf-8")
    os.replace(temporary, path)


def toc_authorization() -> dict[str, object] | None:
    serialized = os.environ.get(TOC_AUTHORIZATION_ENV, "")
    if not serialized:
        return None
    value = json.loads(serialized)
    if (
        not isinstance(value, dict)
        or type(value.get("version")) is not int
        or value.get("version") != 2
        or value.get("expected_document_index_count") != 1
        or not isinstance(value.get("authorization_id"), str)
        or len(value["authorization_id"]) != 64
        or not isinstance(value.get("structure_contract_sha256"), str)
        or len(value["structure_contract_sha256"]) != 64
        or not isinstance(value.get("indexes"), list)
        or len(value["indexes"]) != 1
    ):
        raise RuntimeError("TOC index authorization descriptor is invalid.")
    descriptor = dict(value)
    descriptor.pop("authorization_id", None)
    if canonical_json_hash(descriptor) != value["authorization_id"]:
        raise RuntimeError("TOC index authorization hash is invalid.")
    return value


def toc_contract() -> list[dict[str, object]] | None:
    serialized = os.environ.get(TOC_CONTRACT_ENV, "")
    if not serialized:
        return None
    value = json.loads(serialized)
    if not isinstance(value, list) or not value:
        raise RuntimeError("Approved TOC contract anchor must be a non-empty list.")
    for item in value:
        if (
            not isinstance(item, dict)
            or set(item) != {"level", "kind", "text_sha256"}
            or not isinstance(item.get("level"), int)
            or isinstance(item.get("level"), bool)
            or not 1 <= item["level"] <= 4
            or item.get("kind") not in {"heading", "appendix"}
            or not isinstance(item.get("text_sha256"), str)
            or len(item["text_sha256"]) != 64
            or any(
                character not in "0123456789abcdef"
                for character in item["text_sha256"]
            )
        ):
            raise RuntimeError("Approved TOC contract anchor has an invalid item.")
    return value


def observed_uno_index_identity(candidate: object) -> dict[str, object]:
    return {
        "create_from_outline": bool(candidate.CreateFromOutline),
        "create_from_marks": bool(candidate.CreateFromMarks),
        "level": int(candidate.Level),
    }


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


def refresh_from_environment(*_args: object) -> bool:
    """LibreOffice macro entry point; paths arrive through the isolated process env."""
    result_path = Path(os.environ[RESULT_ENV])
    document = None
    output_path = None
    payload: dict[str, object]
    try:
        input_path = Path(os.environ[INPUT_ENV]).resolve()
        output_path = Path(os.environ[OUTPUT_ENV]).resolve()
        desktop = XSCRIPTCONTEXT.getDesktop()  # type: ignore[name-defined]
        document = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(input_path)),
            "_blank",
            0,
            (
                property_value("Hidden", True),
                property_value("ReadOnly", False),
                # MediaDescriptor safety controls. The user macro is already
                # running; these values prohibit macros embedded in the DOCX
                # and automatic external-link updates while the DOCX loads.
                property_value("UpdateDocMode", NO_UPDATE),
                property_value("MacroExecutionMode", NEVER_EXECUTE),
            ),
        )
        if document is None:
            raise RuntimeError("LibreOffice could not open the DOCX.")

        embedded_script_events = 0
        if hasattr(document, "getEvents"):
            events = document.getEvents()
            for event_name in events.getElementNames():
                properties = events.getByName(event_name)
                if any(
                    getattr(item, "Name", None) == "Script"
                    and bool(getattr(item, "Value", None))
                    for item in properties
                ):
                    embedded_script_events += 1

        fields = document.getTextFields()
        field_enumeration = fields.createEnumeration()
        recognized_text_fields = 0
        recognized_field_services = []
        while field_enumeration.hasMoreElements():
            field = field_enumeration.nextElement()
            recognized_text_fields += 1
            recognized_field_services.extend(field.getSupportedServiceNames())
        recognized_external_connections = external_graphic_urls(document)

        indexes = document.getDocumentIndexes()
        approved_indexes_updated = 0
        skipped_indexes = 0
        authorization = toc_authorization()
        contract = toc_contract()
        index_count = indexes.getCount()
        approved_candidates = []
        observed_index_descriptor = None
        contract_sha256 = None
        if authorization is not None:
            if contract is None:
                raise RuntimeError("Approved TOC contract anchor is missing.")
            contract_sha256 = canonical_json_hash(contract)
            if contract_sha256 != authorization["structure_contract_sha256"]:
                raise RuntimeError(
                    "Approved TOC contract anchor hash does not match authorization."
                )
            if index_count != authorization["expected_document_index_count"]:
                raise RuntimeError(
                    "Document index count does not match the approved TOC authorization."
                )
            observed_uno = []
            for descriptor in authorization["indexes"]:
                ordinal = descriptor["ordinal"]
                candidate = indexes.getByIndex(ordinal)
                if not candidate.supportsService(descriptor["service"]):
                    raise RuntimeError(
                        "Document index identity does not match the approved TOC authorization."
                    )
                observed_uno.append(observed_uno_index_identity(candidate))
                approved_candidates.append(candidate)
            observed_index_descriptor = build_observed_index_descriptor(
                input_path,
                authorization["structure_contract_sha256"],
                observed_uno,
            )
            observed_authorization_id = canonical_json_hash(observed_index_descriptor)
            expected_descriptor = dict(authorization)
            expected_descriptor.pop("authorization_id", None)
            if (
                observed_index_descriptor != expected_descriptor
                or observed_authorization_id != authorization["authorization_id"]
            ):
                raise RuntimeError(
                    "Document index identity does not match the approved TOC authorization."
                )
        elif contract is not None:
            raise RuntimeError(
                "Approved TOC contract anchor was provided without authorization."
            )
        skipped_indexes = index_count - len(approved_candidates)
        for candidate in approved_candidates:
            candidate.update()
            approved_indexes_updated += 1
        # Do not call getTextFields().refresh(): it can actively refresh
        # INCLUDETEXT, INCLUDEPICTURE, DDE, database, or other external fields.
        # calculateAll() is retained for document-internal calculations only.
        calculation_performed = False
        if hasattr(document, "calculateAll"):
            document.calculateAll()
            calculation_performed = True

        output_path.parent.mkdir(parents=True, exist_ok=True)
        document.storeAsURL(
            uno.systemPathToFileUrl(str(output_path)),
            (
                property_value("FilterName", "Office Open XML Text"),
                property_value("Overwrite", True),
            ),
        )
        payload = {
            "ok": True,
            "approved_indexes_updated": approved_indexes_updated,
            "skipped_indexes": skipped_indexes,
            "document_index_count": index_count,
            "toc_authorization_id": (
                authorization.get("authorization_id")
                if authorization is not None
                else None
            ),
            "toc_contract_sha256": contract_sha256,
            "observed_index_descriptor": observed_index_descriptor,
            "text_fields_collection_refreshed": False,
            "calculation_performed": calculation_performed,
            "embedded_script_events": embedded_script_events,
            "recognized_text_fields": recognized_text_fields,
            "recognized_field_services": sorted(set(recognized_field_services)),
            "recognized_external_connections": recognized_external_connections,
            "recognized_external_connection_count": len(
                recognized_external_connections
            ),
        }
    except Exception as exc:  # UNO exposes implementation-specific exceptions.
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
                if output_path is not None:
                    try:
                        output_path.unlink(missing_ok=True)
                    except OSError:
                        pass
                payload = {
                    "ok": False,
                    "error": f"LibreOffice document close failed: {exc}",
                    "traceback": traceback.format_exc(),
                }
    write_result(result_path, payload)
    return bool(payload["ok"])


g_exportedScripts = (refresh_from_environment,)
