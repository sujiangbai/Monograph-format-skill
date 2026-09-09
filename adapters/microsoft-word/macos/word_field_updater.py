#!/usr/bin/env python3
"""Protocol 1.1 adapter for bounded Microsoft Word finalization on macOS."""
from __future__ import annotations

import hashlib
import copy
import json
import os
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any
import zipfile

from lxml import etree


ROOT = Path(__file__).resolve().parents[3]
SCRIPTS = ROOT / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import NS, FormatMonographError, load_document  # noqa: E402
from field_writeback import (  # noqa: E402
    DEFAULT_ALLOWED_FIELD_TYPES,
    _result_text_nodes,
    _validate_page_offset_formula,
    parse_fields,
    _story_roles,
    _semantic_part_sources,
    _matched_records,
    _set_scalar_result,
    selective_field_result_writeback,
)
from structure_map import audit_structure_toc_source_operations  # noqa: E402


PROTOCOL_VERSION = "1.1"
TARGET_ID = "microsoft_word"
BACKEND = "microsoft_word_macos_applescript"
OPERATIONS = frozenset({"measure_layout", "refresh_fields", "verify_only"})
ALLOWED_FIELD_TYPES = DEFAULT_ALLOWED_FIELD_TYPES | {"="}
APPLESCRIPT = Path(__file__).with_name("word_field_updater.applescript")
OSASCRIPT = "/usr/bin/osascript"
TIMEOUT_SECONDS = 570
FIELD_PARTS = (
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
)
STORY_LABELS = {
    ("header", "default"): "header_primary",
    ("header", "first"): "header_first",
    ("header", "even"): "header_even",
    ("footer", "default"): "footer_primary",
    ("footer", "first"): "footer_first",
    ("footer", "even"): "footer_even",
}
STORY_RELATIONSHIP_TYPES = {
    "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
}


class AdapterError(RuntimeError):
    """Fail-closed adapter contract error."""


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _require_host() -> None:
    if (
        sys.platform != "darwin"
        or not Path(OSASCRIPT).is_file()
        or not os.access(OSASCRIPT, os.X_OK)
    ):
        raise AdapterError(
            "The macOS Microsoft Word adapter requires executable /usr/bin/osascript."
        )


def _exact_input(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        raise AdapterError(f"{label} must be an absolute path.")
    path = Path(value)
    if not path.is_absolute() or path.is_symlink() or not path.is_file():
        raise AdapterError(f"{label} must be an existing non-symlink file.")
    resolved = path.resolve(strict=True)
    if str(path) != str(resolved):
        raise AdapterError(f"{label} must already be its exact resolved path.")
    return resolved


def _new_output(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        raise AdapterError(f"{label} must be an absolute path.")
    path = Path(value)
    if not path.is_absolute() or path.exists() or path.is_symlink():
        raise AdapterError(f"{label} must be a new absolute path.")
    parent = path.parent
    if not parent.is_dir() or parent.is_symlink() or parent.resolve() != parent:
        raise AdapterError(f"{label} parent must be an existing exact directory.")
    return path


def _load_object(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise AdapterError(f"{label} is not valid UTF-8 JSON.") from exc
    if not isinstance(value, dict):
        raise AdapterError(f"{label} must be a JSON object.")
    return value


def _field_parts(package: zipfile.ZipFile) -> list[str]:
    names = set(package.namelist())
    dynamic = sorted(
        name
        for name in names
        if name.startswith(("word/header", "word/footer")) and name.endswith(".xml")
    )
    return [name for name in FIELD_PARTS if name in names] + dynamic


def _instruction_manifest(path: Path, allowed: set[str]) -> list[dict[str, Any]]:
    manifest: list[dict[str, Any]] = []
    with zipfile.ZipFile(path) as package:
        for part_name in _field_parts(package):
            try:
                root = etree.fromstring(package.read(part_name))
            except (KeyError, etree.XMLSyntaxError) as exc:
                raise AdapterError(f"Invalid field-bearing part: {part_name}") from exc
            records = parse_fields(root)
            if records and part_name in {
                "word/footnotes.xml",
                "word/endnotes.xml",
            }:
                raise AdapterError(
                    "The macOS adapter does not support fields in footnotes, "
                    "endnotes, or text boxes."
                )
            if root.xpath(
                ".//w:txbxContent//w:fldSimple | .//w:txbxContent//w:fldChar",
                namespaces=NS,
            ):
                raise AdapterError(
                    "The macOS adapter does not support fields in footnotes, "
                    "endnotes, or text boxes."
                )
            if "=" in allowed:
                for record in records:
                    if record.field_type == "=":
                        try:
                            _validate_page_offset_formula(record, records)
                        except FormatMonographError as exc:
                            raise AdapterError(str(exc)) from exc
            for record in records:
                manifest.append(
                    {
                        "part": part_name,
                        "order": record.order,
                        "parent_order": record.parent_order,
                        "form": record.form,
                        "field_type": record.field_type,
                        "instruction": record.instruction,
                    }
                )
    return manifest


def _supplement_footer_caches(
    baseline_path: Path, candidate_path: Path, allowed: set[str],
    plan: list[dict[str, Any]], rows: list[list[Any]],
) -> dict[str, Any]:
    """Bounded model-result supplement; the original Word XML is checked first."""
    original_hash = sha256(candidate_path)
    evidence = {"source": "word_saved_xml", "word_candidate_sha256": original_hash,
                "supplemented_fields": 0}
    if not rows:
        return evidence
    owners = {(item["section_index"], item["story_label"]) for item in plan}
    if any((row[0], row[1]) not in owners or row[2] not in allowed for row in rows):
        raise AdapterError("Footer result has an unapproved owner or field.")
    changes = []
    with zipfile.ZipFile(baseline_path) as baseline, zipfile.ZipFile(candidate_path) as candidate:
        roles = _story_roles(baseline)
        sources, _ = _semantic_part_sources(baseline, candidate)
        for owner in sorted(owners):
            section, label = owner
            if not label.startswith("footer_"):
                continue
            kind = {"footer_primary": "default", "footer_first": "first", "footer_even": "even"}[label]
            part = roles[("footer", section - 1, kind)]
            root = etree.fromstring(baseline.read(part))
            records = parse_fields(root)
            # No new authority for nested or duplicated identities.
            if any(record.parent_order is not None or any(
                parent.tag == f"{{{NS['w']}}}fldSimple"
                for parent in (record.simple if record.form == "simple" else record.begin).iterancestors()
            ) for record in records):
                continue
            source = sources[part][0]
            target_root = etree.fromstring(candidate.read(source))
            target_records = parse_fields(target_root)
            matches = _matched_records(root, records, target_root, target_records, allowed)
            for record, target in matches:
                if record.field_type not in allowed & {"PAGE", "NUMPAGES"}:
                    continue
                if sum(item.semantic_key == record.semantic_key for item in records) != 1:
                    continue
                observed = [row for row in rows if (row[0], row[1]) == owner
                            and row[2] == record.field_type
                            and " ".join(row[3].split()) == record.instruction]
                if len(observed) != 1:
                    continue
                value = observed[0][4]
                old_value = "".join(node.text or "" for node in _result_text_nodes(target_root, target))
                if value != old_value:
                    changes.append((source, target.order, value))
        if not changes:
            return evidence
        if len({(part, order) for part, order, _ in changes}) != len(changes):
            raise AdapterError("Ambiguous footer result destination.")
        # Run the existing whole-package checker on the untouched Word candidate.
        # This temporary output is discarded; it does not become a new parent.
        with tempfile.TemporaryDirectory(prefix="word-footer-cache-") as temporary:
            selective_field_result_writeback(
                baseline_path, candidate_path, Path(temporary) / "checked.docx",
                allowed_field_types=allowed,
            )
            patched = {}
            for part in {part for part, _, _ in changes}:
                target_root = etree.fromstring(candidate.read(part))
                target_records = parse_fields(target_root)
                for _, order, value in (change for change in changes if change[0] == part):
                    target = target_records[order]
                    model_root = copy.deepcopy(target_root)
                    model_record = parse_fields(model_root)[order]
                    nodes = _result_text_nodes(model_root, model_record)
                    if not nodes:
                        raise AdapterError("Footer result has no scalar container.")
                    nodes[0].text = value
                    for node in nodes[1:]:
                        node.text = ""
                    # Existing scalar checks own numeric/payload/container policy.
                    _set_scalar_result(target_root, target, model_root, model_record)
                patched[part] = etree.tostring(target_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            output = Path(temporary) / "supplemented.docx"
            with zipfile.ZipFile(output, "w") as package:
                for info in candidate.infolist():
                    package.writestr(info, patched.get(info.filename, candidate.read(info.filename)))
            selective_field_result_writeback(
                baseline_path, output, Path(temporary) / "rechecked.docx",
                allowed_field_types=allowed,
            )
            # Word has already closed. Only the disposable adapter output changes.
            shutil.copyfile(output, candidate_path)
    evidence.update(source="word_model_footer_supplement", supplemented_fields=len(changes),
                    supplemented_candidate_sha256=sha256(candidate_path))
    return evidence


def _require_verified_footer_caches(
    path: Path, allowed: set[str], plan: list[dict[str, Any]], rows: list[list[Any]],
) -> None:
    """Verify outgoing caches without guessing correspondence or writing values."""
    with zipfile.ZipFile(path) as package:
        roles = _story_roles(package)
        for item in plan:
            label = item["story_label"]
            if not label.startswith("footer_"):
                continue
            section = item["section_index"]
            kind = {"footer_primary": "default", "footer_first": "first", "footer_even": "even"}[label]
            part = roles.get(("footer", section - 1, kind))
            if part is None:
                raise AdapterError("Saved footer cache owner is unavailable.")
            root = etree.fromstring(package.read(part))
            records = parse_fields(root)
            targets = []
            for record in records:
                marker = record.simple if record.form == "simple" else record.begin
                if (record.field_type in allowed & {"PAGE", "NUMPAGES"}
                    and record.parent_order is None
                    and not any(parent.tag == f"{{{NS['w']}}}fldSimple" for parent in marker.iterancestors())):
                    targets.append(record)
            if not targets:
                continue  # No new condition on unrelated/nested-child field kinds.
            if [[record.field_type, record.instruction] for record in records] != item["field_identities"]:
                raise AdapterError("Saved footer cache inventory differs from its observed owner.")
            for record in targets:
                marker = record.simple if record.form == "simple" else record.begin
                if any(child.parent_order == record.order for child in records) or (
                    record.form == "simple" and any(
                        child is not marker and child.tag in {f"{{{NS['w']}}}fldSimple", f"{{{NS['w']}}}fldChar"}
                        for child in marker.iter()
                    )
                ):
                    raise AdapterError("Nested footer cache cannot be verified by the existing association.")
                # Reuse the existing dirty/scalar/container checks on a copy;
                # confirmation never changes the saved candidate.
                checked_root = copy.deepcopy(root)
                _set_scalar_result(checked_root, parse_fields(checked_root)[record.order], root, record)
            for key in {record.semantic_key for record in targets}:
                group = [record for record in targets if record.semantic_key == key]
                observed = [row[4] for row in rows
                            if (row[0], row[1]) == (section, label)
                            and (row[2], " ".join(row[3].split()).casefold()) == key]
                if len(observed) != len(group):
                    raise AdapterError("Saved footer cache has insufficient owner-associated results.")
                # Equal model values prove every matching occurrence without
                # assigning them by ordinal. Different duplicate values cannot.
                if len(set(observed)) != 1:
                    raise AdapterError("Distinct duplicate footer cache results cannot be matched safely.")
                if any("".join(node.text or "" for node in _result_text_nodes(root, record)) != observed[0]
                       for record in group):
                    raise AdapterError("Saved footer cache differs from the converged Word result.")


def _require_saved_default_numpages(path: Path, page_count: int) -> None:
    """Reject stale default NUMPAGES caches; never manufacture saved results.

    Only the unswitched decimal instruction has this direct representation.
    PAGE is location-dependent; formatted NUMPAGES retains existing contracts.
    """
    with zipfile.ZipFile(path) as package:
        for part in _field_parts(package):
            root = etree.fromstring(package.read(part))
            for record in parse_fields(root):
                if record.instruction.strip().upper() != "NUMPAGES":
                    continue
                cached = "".join(
                    node.text or "" for node in _result_text_nodes(root, record)
                )
                if cached.strip() != str(page_count):
                    raise AdapterError("Saved default NUMPAGES cache differs from Word page count.")


def _require_effective_primary_footers(path: Path) -> None:
    """Prove every section inherits or declares one valid primary footer."""
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        try:
            document = etree.fromstring(package.read("word/document.xml"))
            relationships = etree.fromstring(
                package.read("word/_rels/document.xml.rels")
            )
        except (KeyError, etree.XMLSyntaxError) as exc:
            raise AdapterError("Invalid effective primary footer structure.") from exc

        if document.tag != f"{{{NS['w']}}}document":
            raise AdapterError("Invalid effective primary footer structure.")
        sections = document.xpath(
            "./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr",
            namespaces=NS,
        )
        if not sections:
            raise AdapterError("Invalid effective primary footer structure.")

        relationships_namespace = (
            "http://schemas.openxmlformats.org/package/2006/relationships"
        )
        if relationships.tag != f"{{{relationships_namespace}}}Relationships":
            raise AdapterError("Invalid effective primary footer relationship.")
        relationship_map: dict[str, etree._Element] = {}
        for relationship in relationships:
            if relationship.tag != f"{{{relationships_namespace}}}Relationship":
                raise AdapterError("Invalid effective primary footer relationship.")
            identifier = relationship.get("Id")
            if not identifier or identifier in relationship_map:
                raise AdapterError("Invalid effective primary footer relationship.")
            relationship_map[identifier] = relationship

        effective_part: str | None = None
        validated_parts: set[str] = set()
        for section in sections:
            references = section.xpath(
                "./w:footerReference[not(@w:type) or @w:type='default']",
                namespaces=NS,
            )
            if len(references) > 1:
                raise AdapterError("Invalid effective primary footer owner.")
            if references:
                rel_id = references[0].get(f"{{{NS['r']}}}id")
                relationship = relationship_map.get(rel_id or "")
                if relationship is None:
                    raise AdapterError("Invalid effective primary footer relationship.")
                if (
                    relationship.get("TargetMode") not in {None, "Internal"}
                    or relationship.get("Type")
                    != STORY_RELATIONSHIP_TYPES["footer"]
                ):
                    raise AdapterError("Invalid effective primary footer relationship.")
                target = relationship.get("Target")
                if not target:
                    raise AdapterError("Invalid effective primary footer relationship.")
                target_part = posixpath.normpath(posixpath.join("word", target))
                if (
                    not re.fullmatch(r"word/footer\d+\.xml", target_part)
                    or target_part not in names
                ):
                    raise AdapterError("Invalid effective primary footer part.")
                effective_part = target_part

            if effective_part is None:
                raise AdapterError("Missing effective primary footer.")
            if effective_part not in validated_parts:
                try:
                    footer = etree.fromstring(package.read(effective_part))
                except (KeyError, etree.XMLSyntaxError) as exc:
                    raise AdapterError("Invalid effective primary footer part.") from exc
                if footer.tag != f"{{{NS['w']}}}ftr":
                    raise AdapterError("Invalid effective primary footer part.")
                validated_parts.add(effective_part)


def _story_access_plan(path: Path) -> list[dict[str, Any]]:
    """Select one explicit Word owner for every field-bearing header/footer part."""
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        try:
            document = etree.fromstring(package.read("word/document.xml"))
        except (KeyError, etree.XMLSyntaxError) as exc:
            raise AdapterError("Invalid document story structure.") from exc

        field_records_by_part: dict[str, list[Any]] = {}
        for part_name in sorted(
            name
            for name in names
            if name.startswith(("word/header", "word/footer"))
            and name.endswith(".xml")
        ):
            try:
                root = etree.fromstring(package.read(part_name))
                records = parse_fields(root)
            except (KeyError, etree.XMLSyntaxError, FormatMonographError) as exc:
                raise AdapterError("Invalid field-bearing story part.") from exc
            if records:
                if not re.fullmatch(r"word/(?:header|footer)\d+\.xml", part_name):
                    raise AdapterError("Invalid field-bearing story part.")
                story = (
                    "header" if part_name.startswith("word/header") else "footer"
                )
                expected_root = f"{{{NS['w']}}}{'hdr' if story == 'header' else 'ftr'}"
                if root.tag != expected_root:
                    raise AdapterError("Invalid field-bearing story part.")
                field_records_by_part[part_name] = records
        if not field_records_by_part:
            return []

        try:
            relationships = etree.fromstring(
                package.read("word/_rels/document.xml.rels")
            )
        except (KeyError, etree.XMLSyntaxError) as exc:
            raise AdapterError("Invalid document story relationships.") from exc

        relationship_map: dict[str, tuple[str, str]] = {}
        for relationship in relationships:
            identifier = relationship.get("Id")
            relationship_type = relationship.get("Type")
            target = relationship.get("Target")
            if not identifier or identifier in relationship_map:
                raise AdapterError("Invalid document story relationships.")
            if relationship.get("TargetMode") == "External":
                continue
            if relationship_type and target:
                target_part = posixpath.normpath(posixpath.join("word", target))
                relationship_map[identifier] = (relationship_type, target_part)

        owners_by_part: dict[str, list[tuple[int, str]]] = {}
        owner_targets: dict[tuple[int, str], str] = {}
        for section_index, section in enumerate(
            document.xpath(".//w:sectPr", namespaces=NS), start=1
        ):
            for story in ("header", "footer"):
                for reference in section.xpath(f"./w:{story}Reference", namespaces=NS):
                    kind = reference.get(f"{{{NS['w']}}}type", "default")
                    label = STORY_LABELS.get((story, kind))
                    rel_id = reference.get(f"{{{NS['r']}}}id")
                    relationship = relationship_map.get(rel_id or "")
                    if label is None or relationship is None:
                        raise AdapterError("Invalid document story owner.")
                    relationship_type, target_part = relationship
                    if relationship_type != STORY_RELATIONSHIP_TYPES[story]:
                        raise AdapterError("Invalid document story owner.")
                    if not re.fullmatch(rf"word/{story}\d+\.xml", target_part):
                        raise AdapterError("Invalid document story owner.")
                    if target_part not in names:
                        raise AdapterError("Invalid document story owner.")
                    owner_key = (section_index, label)
                    if owner_key in owner_targets:
                        raise AdapterError("Field-bearing story owners are not unique.")
                    owner_targets[owner_key] = target_part
                    owners_by_part.setdefault(target_part, []).append(owner_key)

        plan: list[dict[str, Any]] = []
        for part_name, records in field_records_by_part.items():
            owners = sorted(set(owners_by_part.get(part_name, [])))
            if not owners:
                raise AdapterError("Field-bearing story part has no explicit owner.")
            section_index, label = owners[0]
            plan.append(
                {
                    "section_index": section_index,
                    "story_label": label,
                    "ownership": (
                        "first_section" if section_index == 1 else "independent"
                    ),
                    "field_identities": [
                        [record.field_type, record.instruction] for record in records
                    ],
                }
            )
        owner_keys = {
            (item["section_index"], item["story_label"]) for item in plan
        }
        if len(owner_keys) != len(plan):
            raise AdapterError("Field-bearing story owners are not unique.")
        return plan


def _encode_story_access_plan(plan: list[dict[str, Any]]) -> str:
    return ";".join(
        f"{item['section_index']},{item['story_label']}" for item in plan
    )


def _validate_story_access_observations(
    raw: Any, plan: list[dict[str, Any]]
) -> None:
    if not isinstance(raw, list) or len(raw) != len(plan):
        raise AdapterError("Word story access differed from the approved plan.")
    observed: list[dict[str, Any]] = []
    for row in raw:
        if (
            not isinstance(row, list)
            or len(row) != 5
            or type(row[0]) is not int
            or row[1] not in set(STORY_LABELS.values())
            or row[2] not in {"first_section", "independent"}
            or type(row[3]) is not int
            or row[3] < 1
            or not isinstance(row[4], list)
            or len(row[4]) != row[3]
        ):
            raise AdapterError("Word returned invalid story access evidence.")
        identities = []
        for identity in row[4]:
            if (
                not isinstance(identity, list)
                or len(identity) != 2
                or not isinstance(identity[0], str)
                or not isinstance(identity[1], str)
            ):
                raise AdapterError("Word returned invalid story field identity evidence.")
            identities.append([identity[0], " ".join(identity[1].split())])
        observed.append(
            {
                "section_index": row[0],
                "story_label": row[1],
                "ownership": row[2],
                "field_identities": identities,
            }
        )
    if observed != plan:
        raise AdapterError("Word story access differed from the approved plan.")


def _effective_allowed(
    requested: Any, input_path: Path, structure_map: dict[str, Any]
) -> set[str]:
    if not isinstance(requested, list) or any(
        not isinstance(item, str) for item in requested
    ):
        raise AdapterError("allowed_field_types must be a string list.")
    allowed = {item.upper() for item in requested}
    if not allowed or not allowed <= ALLOWED_FIELD_TYPES:
        raise AdapterError("The request contains an empty or non-approved field whitelist.")
    toc_approved = structure_map.get("toc_source", {}).get("approved") is True
    if not toc_approved:
        allowed.discard("TOC")
    elif "TOC" in allowed:
        failures = audit_structure_toc_source_operations(
            load_document(input_path), structure_map
        )
        if failures:
            raise AdapterError(
                "The existing TOC does not match the approved structure source: "
                + json.dumps(failures, ensure_ascii=False, sort_keys=True)
            )
    return allowed


def _split_toc_entries(value: str) -> list[str]:
    entries = value.split("\r")
    while entries and entries[-1] == "":
        entries.pop()
    return entries


def _hash_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _sanitize_snapshot(
    raw: Any, round_index: int, allowed: set[str]
) -> dict[str, Any]:
    if not isinstance(raw, list) or len(raw) not in {6, 7}:
        raise AdapterError("Word returned an incomplete convergence tuple.")
    toc_texts, toc_spans, section_rows, page_count, field_rows, spacers = raw[:6]
    if not isinstance(toc_texts, list) or any(not isinstance(v, str) for v in toc_texts):
        raise AdapterError("Word returned invalid TOC observations.")
    if (
        not isinstance(toc_spans, list)
        or len(toc_spans) != len(toc_texts)
        or any(type(v) is not int or v < 1 for v in toc_spans)
    ):
        raise AdapterError("Word returned invalid TOC page spans.")
    if type(page_count) is not int or page_count < 1:
        raise AdapterError("Word returned an invalid page count.")
    if not isinstance(section_rows, list) or not section_rows:
        raise AdapterError("Word returned no section measurement.")
    sections = []
    for expected, row in enumerate(section_rows):
        if (
            not isinstance(row, list)
            or len(row) != 8
            or row[0] != expected
            or any(
                type(row[index]) is not int or row[index] < 1
                for index in range(1, 5)
            )
            or type(row[5]) is not bool
            or type(row[6]) is not int
            or row[7] not in {"decimal", "lowerRoman"}
        ):
            raise AdapterError("Word returned an invalid section measurement tuple.")
        sections.append(
            {
                "section_index": row[0],
                "first_physical_page": row[1],
                "last_physical_page": row[2],
                "last_content_page": row[3],
                "first_logical_page": row[4],
                "restart_numbering": row[5],
                "page_number_start": row[6],
                "page_number_format": row[7],
            }
        )
    if not isinstance(field_rows, list):
        raise AdapterError("Word returned invalid field observations.")
    fields = []
    for ordinal, row in enumerate(field_rows, start=1):
        if (
            not isinstance(row, list)
            or len(row) != 3
            or row[0] not in allowed - {"TOC"}
            or not isinstance(row[1], str)
            or not isinstance(row[2], str)
        ):
            raise AdapterError("Word returned an invalid approved-field tuple.")
        fields.append(
            {
                "ordinal": ordinal,
                "field_type": row[0],
                "instruction_sha256": _hash_text(row[1]),
                "result_sha256": _hash_text(row[2]),
            }
        )
    if (
        not isinstance(spacers, list)
        or len(set(spacers)) != len(spacers)
        or any(type(value) is not int or value < 0 for value in spacers)
    ):
        raise AdapterError("Word returned invalid page-boundary spacer ordinals.")
    return {
        "round": round_index,
        "toc_count": len(toc_texts),
        "toc_entries": [
            {
                "toc_ordinal": ordinal,
                "entry_count": len(entries),
                "entry_text_sha256": [_hash_text(entry) for entry in entries],
                "page_span": toc_spans[ordinal - 1],
            }
            for ordinal, text in enumerate(toc_texts, start=1)
            for entries in [_split_toc_entries(text)]
        ],
        "sections": sections,
        "page_count": page_count,
        "field_results": fields,
        "page_boundary_spacer_ordinals": spacers,
    }


def _parse_word_result(
    stdout: str,
    operation: str,
    allowed: set[str],
    story_plan: list[dict[str, Any]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    try:
        raw = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise AdapterError("Word did not return one JSON result.") from exc
    if not isinstance(raw, list) or len(raw) != 16:
        raise AdapterError("Word returned an incomplete adapter result.")
    (
        status,
        observed_operation,
        version,
        docx_saved,
        read_only,
        pdf_exported,
        close_outcome,
        controls_restored,
        structural_changes,
        page_count,
        toc_count,
        verified_count,
        updated_pairs,
        snapshots,
        saved_observations,
        story_observations,
    ) = raw
    _validate_story_access_observations(story_observations, story_plan)
    if (
        status != "success"
        or observed_operation != operation
        or not isinstance(version, str)
        or not version
        or type(docx_saved) is not bool
        or type(read_only) is not bool
        or type(pdf_exported) is not bool
        or close_outcome != "exact_document_closed_without_save"
        or controls_restored is not True
        or structural_changes != 0
    ):
        raise AdapterError("Word did not satisfy the bounded operation contract.")
    expected_flags = {
        "measure_layout": (False, True, False),
        "refresh_fields": (True, False, False),
        "verify_only": (False, True, True),
    }[operation]
    if (docx_saved, read_only, pdf_exported) != expected_flags:
        raise AdapterError("Word returned contradictory save/read-only/PDF evidence.")
    if type(page_count) is not int or page_count < 1 or type(toc_count) is not int or toc_count < 0:
        raise AdapterError("Word returned invalid page or TOC counts.")
    if type(verified_count) is not int or verified_count < 0:
        raise AdapterError("Word returned an invalid verified field count.")
    if not isinstance(updated_pairs, list):
        raise AdapterError("Word returned invalid update counts.")
    updated: dict[str, int] = {}
    for pair in updated_pairs:
        if (
            not isinstance(pair, list)
            or len(pair) != 2
            or pair[0] not in allowed
            or type(pair[1]) is not int
            or pair[1] < 0
            or pair[0] in updated
        ):
            raise AdapterError("Word reported a non-approved field update.")
        updated[pair[0]] = pair[1]
    if operation != "refresh_fields" and updated:
        raise AdapterError("A read-only Word operation reported field updates.")
    if not isinstance(saved_observations, list) or any(type(v) is not bool for v in saved_observations):
        raise AdapterError("Word returned invalid saved-state observations.")
    expected_saved_observations = 4 if operation == "verify_only" else 1
    if (
        len(saved_observations) != expected_saved_observations
        or saved_observations[0] is not True
    ):
        raise AdapterError("Word did not confirm a clean on-disk DOCX before work.")
    if not isinstance(snapshots, list):
        raise AdapterError("Word returned invalid convergence evidence.")
    if operation == "refresh_fields":
        if not 2 <= len(snapshots) <= 3 or snapshots[-1] != snapshots[-2]:
            raise AdapterError("Word fields did not converge within three rounds.")
    elif operation == "measure_layout":
        if len(snapshots) != 1:
            raise AdapterError("Word measurement must return exactly one snapshot.")
    elif len(snapshots) != 2 or snapshots[0] != snapshots[1]:
        raise AdapterError("Word read-only/PDF snapshots changed in memory.")
    sanitized = [
        _sanitize_snapshot(snapshot, index, allowed)
        for index, snapshot in enumerate(snapshots, start=1)
    ]
    footer_rows = []
    for snapshot in snapshots:
        footer_rows = []  # Never reuse associations from an earlier, unconverged round.
        if len(snapshot) == 7:
            association = snapshot[6]
            if not isinstance(association, list) or len(association) != 2:
                raise AdapterError("Invalid footer result association.")
            _validate_story_access_observations(association[0], story_plan)
            rows = association[1]
            if not isinstance(rows, list) or any(
                not isinstance(row, list) or len(row) != 5
                or type(row[0]) is not int
                or row[1] not in {"footer_primary", "footer_first", "footer_even"}
                or row[2] not in {"PAGE", "NUMPAGES"}
                or not isinstance(row[3], str) or not isinstance(row[4], str)
                for row in rows
            ):
                raise AdapterError("Invalid footer result association.")
            remaining = list(snapshot[4])
            for row in rows:
                if row[2:] not in remaining:
                    raise AdapterError("Footer association differs from sampled field results.")
                remaining.remove(row[2:])
            footer_rows = rows
    if sanitized[-1]["page_count"] != page_count or sanitized[-1]["toc_count"] != toc_count:
        raise AdapterError("Word summary conflicts with its complete snapshot.")
    if ("TOC" not in allowed and toc_count != 0) or toc_count > 1:
        raise AdapterError("Word reported an unapproved or ambiguous TOC.")
    # Measurement may observe old caches; correctness is a completion gate,
    # not a prerequisite for read-only layout planning or earlier refresh rounds.
    if operation in {"refresh_fields", "verify_only"}:
        for row in snapshots[-1][4]:
            if (
                row[0] == "NUMPAGES"
                and row[1].strip().upper() == "NUMPAGES"
                and row[2].strip() != str(page_count)
            ):
                raise AdapterError("Default NUMPAGES observation differs from Word page count.")
    return {
        "version": version,
        "updated_fields": updated,
        "verified_field_count": verified_count,
        "page_count": page_count,
        "toc_count": toc_count,
        "saved_observations": saved_observations,
        "footer_rows": footer_rows,
    }, sanitized


def run_request(request: dict[str, Any]) -> dict[str, Any]:
    _require_host()
    if request.get("protocol_version") != PROTOCOL_VERSION:
        raise AdapterError("Unsupported external field protocol version.")
    operation = request.get("operation")
    if operation not in OPERATIONS:
        raise AdapterError("Unsupported external field operation.")
    if request.get("target_software") != TARGET_ID:
        raise AdapterError("This adapter only supports target ID microsoft_word.")
    input_path = _exact_input(request.get("input_path"), "input_path")
    profile_path = _exact_input(request.get("profile_path"), "profile_path")
    structure_path = _exact_input(request.get("structure_map_path"), "structure_map_path")
    structure_map = _load_object(structure_path, "structure_map_path")
    allowed = _effective_allowed(request.get("allowed_field_types"), input_path, structure_map)
    original_hash = sha256(input_path)
    _require_effective_primary_footers(input_path)
    before_instructions = _instruction_manifest(input_path, allowed)
    story_plan = _story_access_plan(input_path)
    expected_fields = sum(
        item["parent_order"] is None and item["field_type"] in allowed
        for item in before_instructions
    )
    expected_toc_count = sum(
        "TOC" in allowed
        and item["parent_order"] is None
        and item["field_type"] == "TOC"
        for item in before_instructions
    )
    output_path: Path | None = None
    pdf_path: Path | None = None
    pdf_workspace: tempfile.TemporaryDirectory | None = None
    created_pdf_identity: tuple[int, int] | None = None
    word_path = input_path
    if operation == "refresh_fields":
        output_path = _new_output(request.get("output_path"), "output_path")
        if output_path in {input_path, profile_path, structure_path}:
            raise AdapterError("The disposable Word output overlaps a protected input.")
        shutil.copy2(input_path, output_path)
        word_path = output_path
    elif operation == "verify_only":
        pdf_path = _new_output(request.get("pdf_output_path"), "pdf_output_path")
        if pdf_path in {input_path, profile_path, structure_path}:
            raise AdapterError("The verification PDF overlaps a protected input.")

    command = [
        OSASCRIPT,
        str(APPLESCRIPT),
        operation,
        str(word_path),
        "",
        "|".join(sorted(allowed)),
        _encode_story_access_plan(story_plan),
    ]
    try:
        if pdf_path is not None:
            # One Word export, always to an adapter-owned short private path.
            # The requested protocol target remains unchanged and is not
            # created until Word's complete verification has succeeded.
            pdf_workspace = tempfile.TemporaryDirectory(prefix="mw-pdf-", dir="/private/tmp")
            word_pdf = Path(pdf_workspace.name).resolve() / "verification.pdf"
            command[4] = str(word_pdf)
        completed = subprocess.run(
            command,
            text=True,
            capture_output=True,
            timeout=TIMEOUT_SECONDS,
            check=False,
        )
        if sha256(input_path) != original_hash:
            raise AdapterError("Microsoft Word changed the protected input DOCX.")
        if completed.returncode != 0:
            detail = completed.stderr.strip().splitlines()
            raise AdapterError(
                "Microsoft Word operation failed: "
                + (detail[-1] if detail else "no bounded error detail")
            )
        parsed, snapshots = _parse_word_result(
            completed.stdout, operation, allowed, story_plan
        )
        if parsed["verified_field_count"] != expected_fields:
            raise AdapterError(
                "Microsoft Word approved-field count differs from the input."
            )
        if parsed["toc_count"] != expected_toc_count:
            raise AdapterError("Microsoft Word TOC count differs from the approved input.")
        if operation == "refresh_fields":
            assert output_path is not None
            if not output_path.is_file() or output_path.stat().st_size == 0:
                raise AdapterError("Microsoft Word did not save its disposable output.")
            footer_evidence = _supplement_footer_caches(
                input_path, output_path, allowed, story_plan, parsed["footer_rows"]
            )
            _require_verified_footer_caches(output_path, allowed, story_plan, parsed["footer_rows"])
            _require_saved_default_numpages(output_path, parsed["page_count"])
        if operation == "verify_only":
            assert pdf_path is not None
            _require_saved_default_numpages(input_path, parsed["page_count"])
            if word_pdf.is_symlink() or not word_pdf.is_file() or word_pdf.stat().st_size == 0:
                raise AdapterError("Microsoft Word did not create the verification PDF.")
            try:
                import pymupdf

                with pymupdf.open(word_pdf) as pdf:
                    pdf_pages = pdf.page_count
            except Exception as exc:
                raise AdapterError("Microsoft Word created an unreadable verification PDF.") from exc
            if pdf_pages != parsed["page_count"]:
                raise AdapterError("Verification PDF page count differs from Word.")
            expected_pdf_hash = sha256(word_pdf)
            _new_output(str(pdf_path), "pdf_output_path")
            with pdf_path.open("xb") as destination:
                owned_stat = os.fstat(destination.fileno())
                created_pdf_identity = (owned_stat.st_dev, owned_stat.st_ino)
                with word_pdf.open("rb") as source:
                    shutil.copyfileobj(source, destination)
            if sha256(pdf_path) != expected_pdf_hash:
                raise AdapterError("Verification PDF transfer changed the bytes.")
        result = {
            "protocol_version": PROTOCOL_VERSION,
            "operation": operation,
            "status": "success",
            "backend": BACKEND,
            "software": "Microsoft Word",
            "software_version": parsed["version"],
            "repaginated": True,
            "saved": operation == "refresh_fields",
            "read_only_verified": operation != "refresh_fields",
            "pdf_exported": operation == "verify_only",
            "structural_changes_applied": 0,
            "page_count": parsed["page_count"],
            "toc_count": parsed["toc_count"],
            "verified_field_count": parsed["verified_field_count"],
            "updated_fields": parsed["updated_fields"],
            "updated_field_types": sorted(parsed["updated_fields"]),
            "field_cache_verified": operation == "refresh_fields",
            "convergence": snapshots,
            "saved_observations": parsed["saved_observations"],
            "input_sha256_unchanged": True,
        }
        final_snapshot = snapshots[-1]
        if operation == "refresh_fields":
            result["footer_cache_source"] = footer_evidence
        if operation == "measure_layout":
            result["sections"] = final_snapshot["sections"]
            result["page_boundary_spacer_ordinals"] = final_snapshot[
                "page_boundary_spacer_ordinals"
            ]
        if pdf_workspace is not None:
            pdf_workspace.cleanup()
            pdf_workspace = None
        return result
    except subprocess.TimeoutExpired as exc:
        for partial in (output_path,):
            if partial is not None:
                try:
                    partial.unlink(missing_ok=True)
                except OSError:
                    pass
        raise AdapterError(
            "Microsoft Word operation timed out; exact close and preference restoration are unconfirmed."
        ) from exc
    except Exception:
        if output_path is not None and output_path.exists():
            output_path.unlink()
        if pdf_path is not None and created_pdf_identity is not None:
            try:
                current = pdf_path.lstat()
                if (current.st_dev, current.st_ino) == created_pdf_identity:
                    pdf_path.unlink()
            except FileNotFoundError:
                pass
        raise
    finally:
        if pdf_workspace is not None:
            pdf_workspace.cleanup()


def main() -> int:
    try:
        request = json.load(sys.stdin)
        if not isinstance(request, dict):
            raise AdapterError("External field request must be a JSON object.")
        response = run_request(request)
    except (AdapterError, OSError, ValueError, zipfile.BadZipFile) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(json.dumps(response, ensure_ascii=False, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
