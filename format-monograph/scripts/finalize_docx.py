#!/usr/bin/env python3
"""Finalize field caches and re-audit a formatted DOCX without overwriting inputs."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shlex
import shutil
import socket
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree
from docx.enum.section import WD_HEADER_FOOTER
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.hdrftr import FooterPart

from _common import (
    NS,
    FormatMonographError,
    ensure_docx,
    field_cache_inventory,
    font_alias_keys,
    load_document,
    protected_payload_manifest,
    run_effective_font,
    style_effective_font,
    style_name_for_selector,
    write_json,
)
from render_docx import locate_soffice
from structure_map import (
    approved_data_tables,
    approved_role_paragraphs,
    has_semantic_structure_map,
    toc_result_contract,
)
from structure_map import (
    load_structure_map,
    structure_content_fingerprint,
    validate_structure_map_source,
)
from validate_profile import validate
from field_writeback import (
    DEFAULT_ALLOWED_FIELD_TYPES,
    selective_field_result_writeback,
)
from docx_pagination import _page_only_footer, _replace_with_page_field


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def controlled_field_result_writeback(
    baseline_path: Path, refreshed_path: Path, output_path: Path
) -> list[str]:
    """Compatibility wrapper around V0.3.2 selective field-result writeback."""
    report = selective_field_result_writeback(
        baseline_path,
        refreshed_path,
        output_path,
        allowed_field_types=DEFAULT_ALLOWED_FIELD_TYPES,
    )
    return list(report["patched_parts"])


def rewrite_field_flags(path: Path, *, deferred: bool) -> None:
    temp_path = path.with_name(f".{path.name}.fields.tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temp_path, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                for field in root.xpath(
                    ".//w:fldSimple | .//w:fldChar[@w:fldCharType='begin']",
                    namespaces=NS,
                ):
                    dirty = f"{{{NS['w']}}}dirty"
                    if deferred:
                        field.set(dirty, "true")
                    else:
                        field.attrib.pop(dirty, None)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif info.filename == "word/settings.xml":
                root = etree.fromstring(data)
                updates = root.xpath("./w:updateFields", namespaces=NS)
                if deferred and not updates:
                    update = etree.Element(f"{{{NS['w']}}}updateFields")
                    update.set(f"{{{NS['w']}}}val", "true")
                    root.append(update)
                elif not deferred:
                    for update in updates:
                        root.remove(update)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            target.writestr(info, data)
    os.replace(temp_path, path)


def free_port() -> int:
    with socket.socket() as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def uno_python_candidates(soffice: str) -> list[str]:
    program = Path(soffice).resolve().parent
    candidates = [
        program / ("python.exe" if os.name == "nt" else "python"),
        Path("/usr/bin/python3"),
        Path(sys.executable),
    ]
    result = []
    for candidate in candidates:
        value = str(candidate)
        if candidate.is_file() and value not in result:
            result.append(value)
    return result


def locate_uno_python(soffice: str, env: dict[str, str]) -> str:
    for candidate in uno_python_candidates(soffice):
        checked = subprocess.run(
            [candidate, "-c", "import uno"],
            capture_output=True,
            text=True,
            env=env,
            timeout=20,
            check=False,
        )
        if checked.returncode == 0:
            return candidate
    raise FormatMonographError(
        "LibreOffice was found, but no Python runtime with the UNO module was available."
    )


def libreoffice_refresh(
    input_path: Path, output_path: Path, renderer: str | None
) -> dict:
    soffice, renderer_source = locate_soffice(renderer)
    helper = Path(__file__).with_name("libreoffice_fields.py")
    port = free_port()
    with tempfile.TemporaryDirectory(prefix="format-monograph-fields-") as temp_name:
        temp = Path(temp_name)
        profile = temp / "lo-profile"
        profile.mkdir()
        env = os.environ.copy()
        env["PATH"] = str(Path(soffice).resolve().parent) + os.pathsep + env.get(
            "PATH", ""
        )
        python = locate_uno_python(soffice, env)
        server = subprocess.Popen(
            [
                soffice,
                "--headless",
                "--invisible",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile.resolve().as_uri()}",
                f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
        )
        try:
            completed = subprocess.run(
                [
                    python,
                    str(helper),
                    str(input_path.resolve()),
                    str(output_path.resolve()),
                    "--port",
                    str(port),
                ],
                capture_output=True,
                text=True,
                env=env,
                timeout=300,
                check=False,
            )
        finally:
            server.terminate()
            try:
                server.wait(timeout=15)
            except subprocess.TimeoutExpired:
                server.kill()
        if completed.returncode != 0 or not output_path.is_file():
            raise FormatMonographError(
                "LibreOffice field refresh failed. "
                f"stdout={completed.stdout.strip()} stderr={completed.stderr.strip()}"
            )
        rewrite_field_flags(output_path, deferred=False)
        details = json.loads(completed.stdout or "{}")
        details.update(
            {
                "backend": "libreoffice_uno",
                "renderer": soffice,
                "renderer_source": renderer_source,
                "uno_python": python,
            }
        )
        return details


def field_contract_preserved(before: dict, after: dict) -> bool:
    if before["main_toc_fields"] > after["main_toc_fields"]:
        return False
    after_types = after.get("field_types", {})
    return all(
        int(after_types.get(name, 0)) >= int(count)
        for name, count in before.get("field_types", {}).items()
    )


def effective_font_failures(
    path: Path, profile: dict, structure_map: dict | None = None
) -> list[dict[str, Any]]:
    document = load_document(path)
    result = []
    property_attributes = {
        "font_name": "ascii",
        "font_name_ascii": "ascii",
        "font_name_east_asia": "eastAsia",
        "font_name_complex_script": "cs",
    }
    for rule in profile.get("rules", []):
        if rule.get("status") != "approved" or rule.get("application") != "automatic":
            continue
        selector = rule.get("selector", {})
        selector_kind = selector.get("kind")
        if selector_kind == "table_role":
            try:
                targets = (
                    approved_data_tables(document, structure_map)
                    if structure_map and has_semantic_structure_map(structure_map)
                    else [(table, {}) for table in document.tables]
                )
            except FormatMonographError:
                result.append(
                    {
                        "rule": rule.get("id"),
                        "selector": selector_kind,
                        "reason": "semantic_target_unresolvable",
                    }
                )
                continue
            for table_index, (table, entry) in enumerate(targets):
                for property_name, attribute in property_attributes.items():
                    expected = rule.get("properties", {}).get(property_name)
                    if not expected:
                        continue
                    for row_index, row in enumerate(table.rows):
                        if entry.get("caption_row") is not None and row_index == int(
                            entry["caption_row"]
                        ):
                            continue
                        for cell_index, cell in enumerate(row.cells):
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    if not run.text:
                                        continue
                                    actual, source = run_effective_font(
                                        document, paragraph, run, attribute
                                    )
                                    if not actual or not (
                                        font_alias_keys(str(actual))
                                        & font_alias_keys(str(expected))
                                    ):
                                        result.append(
                                            {
                                                "rule": rule.get("id"),
                                                "table": table_index,
                                                "row": row_index,
                                                "cell": cell_index,
                                                "property": property_name,
                                                "expected": str(expected),
                                                "actual": actual,
                                                "source": source,
                                            }
                                        )
            continue
        if (
            structure_map
            and has_semantic_structure_map(structure_map)
            and selector_kind
            in {"paragraph_role", "caption_role", "bibliography_role"}
        ):
            try:
                targets = approved_role_paragraphs(document, structure_map, selector)
            except FormatMonographError:
                result.append(
                    {
                        "rule": rule.get("id"),
                        "selector": selector_kind,
                        "reason": "semantic_target_unresolvable",
                    }
                )
                continue
            if targets:
                for property_name, attribute in property_attributes.items():
                    expected = rule.get("properties", {}).get(property_name)
                    if not expected:
                        continue
                    for paragraph_index, paragraph in enumerate(targets):
                        for run_index, run in enumerate(paragraph.runs):
                            if not run.text:
                                continue
                            actual, source = run_effective_font(
                                document, paragraph, run, attribute
                            )
                            if not actual or not (
                                font_alias_keys(str(actual))
                                & font_alias_keys(str(expected))
                            ):
                                result.append(
                                    {
                                        "rule": rule.get("id"),
                                        "paragraph": paragraph_index,
                                        "run": run_index,
                                        "property": property_name,
                                        "expected": str(expected),
                                        "actual": actual,
                                        "source": source,
                                    }
                                )
                continue
        style_name = style_name_for_selector(rule.get("selector", {}))
        if not style_name:
            continue
        try:
            style = document.styles[style_name]
        except KeyError:
            result.append(
                {"rule": rule.get("id"), "style": style_name, "reason": "missing_style"}
            )
            continue
        for property_name, attribute in property_attributes.items():
            expected = rule.get("properties", {}).get(property_name)
            if not expected:
                continue
            actual, source = style_effective_font(document, style, attribute)
            if not actual or not (
                font_alias_keys(str(actual)) & font_alias_keys(str(expected))
            ):
                result.append(
                    {
                        "rule": rule.get("id"),
                        "style": style_name,
                        "property": property_name,
                        "expected": str(expected),
                        "actual": actual,
                        "source": source,
                    }
                )
    return result


def use_deferred_output(input_path: Path, output_path: Path, reason: str) -> dict:
    output_path.unlink(missing_ok=True)
    shutil.copy2(input_path, output_path)
    rewrite_field_flags(output_path, deferred=True)
    return {"backend": "deferred_on_open", "fallback_from": reason}


def _external_command(value: str) -> list[str]:
    stripped = value.strip()
    if stripped.startswith("["):
        parsed = json.loads(stripped)
        if not isinstance(parsed, list) or not parsed or not all(
            isinstance(item, str) and item for item in parsed
        ):
            raise FormatMonographError(
                "--field-updater-command JSON must be a non-empty string array."
            )
        return parsed
    parts = shlex.split(stripped, posix=os.name != "nt")
    if not parts:
        raise FormatMonographError("--field-updater-command cannot be empty.")
    return parts


def external_refresh(
    input_path: Path,
    output_path: Path,
    command: str,
    profile_path: Path,
    structure_map_path: Path,
    pdf_output: Path | None,
    target_software: str,
    *,
    allowed_field_types: set[str] | frozenset[str] = DEFAULT_ALLOWED_FIELD_TYPES,
) -> dict:
    request = {
        "protocol_version": "1.1",
        "operation": "refresh_fields",
        "input_path": str(input_path.resolve()),
        "output_path": str(output_path.resolve()),
        "profile_path": str(profile_path.resolve()),
        "structure_map_path": str(structure_map_path.resolve()),
        "allowed_field_types": sorted(allowed_field_types),
        "target_software": target_software,
        "pdf_output_path": str(pdf_output.resolve()) if pdf_output else None,
    }
    input_hash = file_sha256(input_path)
    completed = subprocess.run(
        _external_command(command),
        input=json.dumps(request, ensure_ascii=False),
        capture_output=True,
        text=True,
        timeout=600,
        check=False,
    )
    if file_sha256(input_path) != input_hash:
        raise FormatMonographError("External field updater changed its input DOCX.")
    if completed.returncode != 0:
        raise FormatMonographError(
            "External field updater failed. "
            f"stdout={completed.stdout.strip()} stderr={completed.stderr.strip()}"
        )
    try:
        response = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise FormatMonographError(
            "External field updater did not return one JSON response."
        ) from exc
    if not isinstance(response, dict):
        raise FormatMonographError("External field updater response must be an object.")
    if not str(response.get("software", "")).strip():
        raise FormatMonographError(
            "External field updater must report its target software."
        )
    required_true = ("repaginated", "saved", "field_cache_verified")
    if response.get("status") != "success" or any(
        response.get(name) is not True for name in required_true
    ):
        raise FormatMonographError(
            "External field updater did not confirm repagination, save, and cache verification."
        )
    if response.get("structural_changes_applied") != 0:
        raise FormatMonographError(
            "External field updater changed pagination or document structure."
        )
    if not output_path.is_file() or output_path.stat().st_size == 0:
        raise FormatMonographError("External field updater did not create the output DOCX.")
    updated_types = set(response.get("updated_field_types", []))
    allowed_types = set(request["allowed_field_types"])
    if not updated_types <= allowed_types:
        raise FormatMonographError(
            "External field updater reported a non-approved field type."
        )
    response.setdefault("backend", "external")
    response["command"] = _external_command(command)[0]
    return response


def external_measure(
    input_path: Path,
    command: str,
    profile_path: Path,
    structure_map_path: Path,
    target_software: str,
) -> dict:
    request = {
        "protocol_version": "1.1",
        "operation": "measure_layout",
        "input_path": str(input_path.resolve()),
        "profile_path": str(profile_path.resolve()),
        "structure_map_path": str(structure_map_path.resolve()),
        "allowed_field_types": sorted(DEFAULT_ALLOWED_FIELD_TYPES),
        "target_software": target_software,
        "block_spacer_style_name": "Monograph Figure Table Spacer",
    }
    input_hash = file_sha256(input_path)
    completed = subprocess.run(
        _external_command(command),
        input=json.dumps(request, ensure_ascii=False),
        capture_output=True,
        text=True,
        timeout=600,
        check=False,
    )
    if file_sha256(input_path) != input_hash:
        raise FormatMonographError("External layout measurer changed its input DOCX.")
    if completed.returncode != 0:
        raise FormatMonographError(
            "External layout measurement failed. "
            f"stdout={completed.stdout.strip()} stderr={completed.stderr.strip()}"
        )
    try:
        response = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise FormatMonographError(
            "External layout measurer did not return one JSON response."
        ) from exc
    required = {
        "status": "success",
        "operation": "measure_layout",
        "repaginated": True,
        "saved": False,
        "read_only_verified": True,
        "structural_changes_applied": 0,
    }
    if not isinstance(response, dict) or any(
        response.get(name) != value for name, value in required.items()
    ):
        raise FormatMonographError(
            "External layout measurer did not satisfy the read-only contract."
        )
    ordinals = response.get("page_boundary_spacer_ordinals", [])
    if not isinstance(ordinals, list) or any(
        not isinstance(value, int) or value < 0 for value in ordinals
    ):
        raise FormatMonographError(
            "External layout measurer returned invalid spacer ordinals."
        )
    sections = response.get("sections", [])
    if not isinstance(sections, list) or any(
        not isinstance(item, dict) for item in sections
    ):
        raise FormatMonographError(
            "External layout measurer returned invalid section metrics."
        )
    page_count = response.get("page_count")
    if not isinstance(page_count, int) or isinstance(page_count, bool) or page_count < 1:
        raise FormatMonographError(
            "External layout measurer returned an invalid page count."
        )
    response.setdefault("backend", "external")
    response["command"] = _external_command(command)[0]
    return response


def approved_front_matter_section_indexes(
    input_path: Path,
    structure_map: dict[str, Any],
) -> set[int]:
    front_matter = structure_map.get("front_matter", {})
    pagination = structure_map.get("pagination_sections", {})
    if not front_matter.get("approved") or not pagination.get("approved"):
        return set()
    with zipfile.ZipFile(input_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    sections = root.xpath(".//w:sectPr", namespaces=NS)
    restart_indexes = [
        index
        for index, section in enumerate(sections)
        if section.find(qn("w:pgNumType")) is not None
        and section.find(qn("w:pgNumType")).get(qn("w:start")) is not None
    ]
    if len(restart_indexes) != 2 or restart_indexes[1] != restart_indexes[0] + 1:
        raise FormatMonographError(
            "Approved title/TOC/body pagination requires exactly two adjacent restarts."
        )
    return set(restart_indexes)


def approved_front_matter_section_types(
    section_indexes: set[int],
    measurement: dict[str, Any],
    input_path: Path,
) -> dict[int, str]:
    if not section_indexes:
        return {}
    document = load_document(input_path)
    metrics = {
        int(item["section_index"]): item
        for item in measurement.get("sections", [])
        if isinstance(item, dict) and "section_index" in item
    }
    result = {}
    for index in sorted(section_indexes):
        previous = metrics.get(index - 1)
        if previous is None or previous.get("last_content_page") is None:
            raise FormatMonographError(
                "Target layout measurement omitted a front-matter section boundary."
            )
        desired_page = int(previous["last_content_page"]) + 1
        target = "evenPage" if desired_page % 2 == 0 else "oddPage"
        section_type = document.sections[index]._sectPr.find(qn("w:type"))
        current = (
            "nextPage"
            if section_type is None
            else section_type.get(qn("w:val"), "nextPage")
        )
        if current != target:
            result[index] = target
    return result


def apply_measured_layout_adjustments(
    input_path: Path,
    output_path: Path,
    ordinals: list[int],
    section_types: dict[int, str] | None = None,
) -> int:
    section_types = section_types or {}
    if any(value not in {"evenPage", "oddPage"} for value in section_types.values()):
        raise FormatMonographError("Measured section type is not an approved parity start.")
    selected = set(ordinals)
    if len(selected) != len(ordinals):
        raise FormatMonographError("Measured spacer ordinals must be unique.")
    with zipfile.ZipFile(input_path) as package:
        document_data = package.read("word/document.xml")
        root = etree.fromstring(document_data)
        sections = root.xpath(".//w:sectPr", namespaces=NS)
        for index, value in section_types.items():
            if not 0 <= index < len(sections):
                raise FormatMonographError(
                    "Measured section index is outside the DOCX section set."
                )
            section_type = sections[index].find(qn("w:type"))
            if section_type is None:
                section_type = etree.Element(qn("w:type"))
                sections[index].insert(0, section_type)
            section_type.set(qn("w:val"), value)
        spacers = root.xpath(
            ".//w:p[w:pPr/w:pStyle[@w:val='MonographFigureTableSpacer']]",
            namespaces=NS,
        )
        if selected and max(selected) >= len(spacers):
            raise FormatMonographError(
                "Measured spacer ordinal is outside the approved spacer set."
            )
        for ordinal in sorted(selected, reverse=True):
            spacer = spacers[ordinal]
            if spacer.xpath(
                ".//w:t[normalize-space(.) != ''] | .//w:drawing | .//w:object | "
                ".//w:pict | .//w:fldChar | .//w:instrText | .//w:sectPr",
                namespaces=NS,
            ):
                raise FormatMonographError(
                    "Measured page-boundary spacer contains authored or structural payload."
                )
            parent = spacer.getparent()
            if parent is None:
                raise FormatMonographError("Measured spacer has no document parent.")
            parent.remove(spacer)
        patched = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        temp = output_path.with_name(f".{output_path.name}.spacers.tmp")
        temp.unlink(missing_ok=True)
        try:
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
                for info in package.infolist():
                    target.writestr(
                        info,
                        patched if info.filename == "word/document.xml" else package.read(info.filename),
                    )
            temp.replace(output_path)
        finally:
            temp.unlink(missing_ok=True)
    if protected_payload_manifest(input_path) != protected_payload_manifest(output_path):
        output_path.unlink(missing_ok=True)
        raise FormatMonographError(
            "Core spacer normalization changed a protected payload."
        )
    return len(selected) + len(section_types)


def remove_measured_block_spacers(
    input_path: Path,
    output_path: Path,
    ordinals: list[int],
) -> int:
    return apply_measured_layout_adjustments(
        input_path,
        output_path,
        ordinals,
    )


def _append_page_offset_formula(paragraph: Any, offset: int) -> None:
    if offset != 1:
        raise FormatMonographError("Only the approved PAGE-minus-one offset is supported.")

    def marker(kind: str) -> Any:
        run = OxmlElement("w:r")
        field = OxmlElement("w:fldChar")
        field.set(qn("w:fldCharType"), kind)
        run.append(field)
        return run

    def instruction(value: str) -> Any:
        run = OxmlElement("w:r")
        node = OxmlElement("w:instrText")
        node.set(qn("xml:space"), "preserve")
        node.text = value
        run.append(node)
        return run

    paragraph._p.append(marker("begin"))
    paragraph._p.append(instruction(" = "))
    paragraph._p.append(marker("begin"))
    paragraph._p.append(instruction(" PAGE "))
    paragraph._p.append(marker("separate"))
    inner_result = OxmlElement("w:r")
    inner_text = OxmlElement("w:t")
    inner_text.text = "2"
    inner_result.append(inner_text)
    paragraph._p.append(inner_result)
    paragraph._p.append(marker("end"))
    paragraph._p.append(instruction(" - 1 "))
    paragraph._p.append(marker("separate"))
    outer_result = OxmlElement("w:r")
    outer_text = OxmlElement("w:t")
    outer_text.text = "1"
    outer_result.append(outer_text)
    paragraph._p.append(outer_result)
    paragraph._p.append(marker("end"))


def _isolate_page_footer(
    document: Any,
    section: Any,
    footer_type: WD_HEADER_FOOTER,
) -> Any:
    section._sectPr.remove_footerReference(footer_type)
    footer_part = FooterPart.new(document.part.package)
    relationship_id = document.part.relate_to(footer_part, RT.FOOTER)
    section._sectPr.add_footerReference(footer_type, relationship_id)
    return (
        section.footer
        if footer_type == WD_HEADER_FOOTER.PRIMARY
        else section.even_page_footer
    )


def _drop_unused_footer_relationships(document: Any) -> None:
    used = {
        reference.get(qn("r:id"))
        for section in document.sections
        for reference in section._sectPr.xpath("./w:footerReference")
    }
    for relationship_id, relationship in list(document.part.rels.items()):
        if relationship.reltype == RT.FOOTER and relationship_id not in used:
            document.part.drop_rel(relationship_id)


def apply_page_display_offsets(
    input_path: Path,
    output_path: Path,
    section_offsets: dict[int, int],
) -> int:
    if any(value != 1 for value in section_offsets.values()):
        raise FormatMonographError("Only a one-page parity offset is supported.")
    document = load_document(input_path)
    changed = 0
    for index, offset in sorted(section_offsets.items()):
        if not 0 <= index < len(document.sections):
            raise FormatMonographError("Page display offset section is out of range.")
        section = document.sections[index]
        for footer_type, alignment in (
            (WD_HEADER_FOOTER.PRIMARY, WD_ALIGN_PARAGRAPH.RIGHT),
            (WD_HEADER_FOOTER.EVEN_PAGE, WD_ALIGN_PARAGRAPH.LEFT),
        ):
            footer = (
                section.footer
                if footer_type == WD_HEADER_FOOTER.PRIMARY
                else section.even_page_footer
            )
            if not _page_only_footer(footer):
                raise FormatMonographError(
                    "Page display offset requires a page-only footer."
                )
            if index + 1 < len(document.sections):
                next_section = document.sections[index + 1]
                if next_section._sectPr.get_footerReference(footer_type) is None:
                    next_footer = (
                        next_section.footer
                        if footer_type == WD_HEADER_FOOTER.PRIMARY
                        else next_section.even_page_footer
                    )
                    if not _page_only_footer(next_footer):
                        raise FormatMonographError(
                            "The following section inherits a non-page footer."
                        )
                    next_footer = _isolate_page_footer(
                        document,
                        next_section,
                        footer_type,
                    )
                    _replace_with_page_field(next_footer, alignment)
                    changed += 1
            footer = _isolate_page_footer(document, section, footer_type)
            paragraphs = list(footer.paragraphs)
            paragraph = paragraphs[0] if paragraphs else footer.add_paragraph()
            for extra in paragraphs[1:]:
                footer._element.remove(extra._p)
            for child in list(paragraph._p):
                if child.tag != qn("w:pPr"):
                    paragraph._p.remove(child)
            paragraph.alignment = alignment
            _append_page_offset_formula(paragraph, offset)
            changed += 1
    _drop_unused_footer_relationships(document)
    document.save(output_path)
    if protected_payload_manifest(input_path) != protected_payload_manifest(output_path):
        output_path.unlink(missing_ok=True)
        raise FormatMonographError(
            "Core page display offset changed a protected payload."
        )
    return changed


def external_verify(
    input_path: Path,
    command: str,
    profile_path: Path,
    structure_map_path: Path,
    pdf_output: Path,
    target_software: str,
    *,
    expected_page_count: int | None = None,
    allowed_field_types: set[str] | frozenset[str] = DEFAULT_ALLOWED_FIELD_TYPES,
) -> dict:
    request = {
        "protocol_version": "1.1",
        "operation": "verify_only",
        "input_path": str(input_path.resolve()),
        "profile_path": str(profile_path.resolve()),
        "structure_map_path": str(structure_map_path.resolve()),
        "allowed_field_types": sorted(allowed_field_types),
        "target_software": target_software,
        "pdf_output_path": str(pdf_output.resolve()),
        "expected_page_count": expected_page_count,
    }
    input_hash = file_sha256(input_path)
    completed = subprocess.run(
        _external_command(command),
        input=json.dumps(request, ensure_ascii=False),
        capture_output=True,
        text=True,
        timeout=600,
        check=False,
    )
    if file_sha256(input_path) != input_hash:
        raise FormatMonographError("External verifier changed its input DOCX.")
    if completed.returncode != 0:
        raise FormatMonographError(
            "External read-only verification failed. "
            f"stdout={completed.stdout.strip()} stderr={completed.stderr.strip()}"
        )
    try:
        response = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise FormatMonographError(
            "External verifier did not return one JSON response."
        ) from exc
    required = {
        "status": "success",
        "operation": "verify_only",
        "repaginated": True,
        "saved": False,
        "read_only_verified": True,
        "pdf_exported": True,
    }
    if not isinstance(response, dict) or any(
        response.get(name) != value for name, value in required.items()
    ):
        raise FormatMonographError(
            "External verifier did not satisfy the read-only verification contract."
        )
    if not pdf_output.is_file() or pdf_output.stat().st_size == 0:
        raise FormatMonographError("External verifier did not create its target PDF.")
    actual_page_count = response.get("page_count")
    if expected_page_count is not None:
        if (
            not isinstance(actual_page_count, int)
            or isinstance(actual_page_count, bool)
            or actual_page_count < 1
        ):
            raise FormatMonographError(
                "External verifier omitted a valid page count."
            )
        if actual_page_count != int(expected_page_count):
            raise FormatMonographError(
                "Selective output page count differs from the field calculation session."
            )
    response.setdefault("backend", "external")
    response["command"] = _external_command(command)[0]
    return response


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--profile", required=True, type=Path)
    parser.add_argument("--structure-map", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--status-output", type=Path)
    parser.add_argument(
        "--field-updater",
        choices=("auto", "external", "libreoffice", "deferred"),
        default="auto",
    )
    parser.add_argument(
        "--field-updater-command",
        help="External updater command or a JSON array of command arguments.",
    )
    parser.add_argument(
        "--pdf-output",
        type=Path,
        help="Optional target-software PDF path requested from an external backend.",
    )
    parser.add_argument(
        "--target-software",
        help="Target application requested from an external field backend.",
    )
    parser.add_argument("--renderer")
    parser.add_argument("--approve-deferred", action="store_true")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    try:
        ensure_docx(args.input)
        if args.source:
            ensure_docx(args.source)
        errors, profile = validate(args.profile)
        if errors or profile["approval"]["status"] != "approved":
            raise FormatMonographError(
                "Finalization requires an approved valid profile: " + "; ".join(errors)
            )
        structure_map = load_structure_map(args.structure_map)
        if args.source:
            validate_structure_map_source(args.source, structure_map)
        if args.output.exists() and not args.force:
            raise FormatMonographError("Final output exists; use --force to replace it.")
        if args.output.resolve() == args.input.resolve() or (
            args.source and args.output.resolve() == args.source.resolve()
        ):
            raise FormatMonographError("Finalization must not overwrite an input DOCX.")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        if args.pdf_output:
            args.pdf_output.parent.mkdir(parents=True, exist_ok=True)
            if args.pdf_output.exists() and not args.force:
                raise FormatMonographError("PDF output exists; use --force to replace it.")
            args.pdf_output.unlink(missing_ok=True)
        args.output.unlink(missing_ok=True)

        baseline = args.source or args.input
        baseline_fp = structure_content_fingerprint(baseline, structure_map)
        input_fp = structure_content_fingerprint(args.input, structure_map)
        if baseline_fp != input_fp:
            raise FormatMonographError(
                "Formatted input failed the stable pre-finalization content audit."
            )
        baseline_objects = protected_payload_manifest(baseline)
        approved_toc_contract = toc_result_contract(
            load_document(args.input), structure_map
        )
        input_font_failures = effective_font_failures(
            args.input, profile, structure_map
        )
        if input_font_failures:
            raise FormatMonographError(
                "Formatted input failed deterministic effective-font validation: "
                + json.dumps(input_font_failures, ensure_ascii=False)
            )
        input_fields = field_cache_inventory(args.input)
        backend: dict = {"backend": "not_needed"}
        delivery_status = input_fields["status"]

        external_requested = args.field_updater == "external" or (
            args.field_updater == "auto" and bool(args.field_updater_command)
        )
        if external_requested:
            if not args.field_updater_command:
                raise FormatMonographError(
                    "External field update requires --field-updater-command."
                )
            try:
                with tempfile.TemporaryDirectory(
                    prefix="format-monograph-external-fields-"
                ) as refresh_name:
                    refresh_root = Path(refresh_name)
                    refresh_input = args.input
                    measurements = []
                    section_adjustments = 0
                    spacers_removed = 0
                    pagination_section_indexes = (
                        approved_front_matter_section_indexes(
                            args.input,
                            structure_map,
                        )
                    )
                    remove_boundary_spacers = bool(
                        structure_map.get("block_spacing", {}).get("approved")
                        and structure_map.get("block_spacing", {}).get(
                            "same_page_only"
                        )
                    )
                    for measurement_pass in range(10):
                        measurement = external_measure(
                            refresh_input,
                            args.field_updater_command,
                            args.profile,
                            args.structure_map,
                            args.target_software
                            or profile["target_applications"][0],
                        )
                        measurements.append(measurement)
                        ordinals = (
                            measurement.get("page_boundary_spacer_ordinals", [])
                            if remove_boundary_spacers
                            else []
                        )
                        section_types = approved_front_matter_section_types(
                            pagination_section_indexes,
                            measurement,
                            refresh_input,
                        )
                        if not ordinals and not section_types:
                            break
                        normalized = refresh_root / (
                            f"layout-normalized-{measurement_pass + 1}.docx"
                        )
                        apply_measured_layout_adjustments(
                            refresh_input,
                            normalized,
                            ordinals,
                            section_types,
                        )
                        section_adjustments += len(section_types)
                        spacers_removed += len(ordinals)
                        refresh_input = normalized
                    else:
                        raise FormatMonographError(
                            "Page-boundary spacer normalization did not converge."
                        )
                    display_offsets = {}
                    normalized_document = load_document(refresh_input)
                    for index in sorted(pagination_section_indexes):
                        section_type = normalized_document.sections[
                            index
                        ]._sectPr.find(qn("w:type"))
                        if (
                            section_type is not None
                            and section_type.get(qn("w:val")) == "evenPage"
                        ):
                            display_offsets[index] = 1
                    allowed_field_types = set(DEFAULT_ALLOWED_FIELD_TYPES)
                    if display_offsets:
                        offset_input = refresh_root / "page-display-offsets.docx"
                        apply_page_display_offsets(
                            refresh_input,
                            offset_input,
                            display_offsets,
                        )
                        refresh_input = offset_input
                        allowed_field_types.add("=")
                    writeback_field_types = set(allowed_field_types)
                    if approved_toc_contract is not None:
                        writeback_field_types.add("TC")
                    refreshed = refresh_root / "refreshed.docx"
                    backend = external_refresh(
                        refresh_input,
                        refreshed,
                        args.field_updater_command,
                        args.profile,
                        args.structure_map,
                        None,
                        args.target_software or profile["target_applications"][0],
                        allowed_field_types=allowed_field_types,
                    )
                    writeback = selective_field_result_writeback(
                        refresh_input,
                        refreshed,
                        args.output,
                        allowed_field_types=writeback_field_types,
                        toc_contract=approved_toc_contract,
                    )
                    backend["layout_measurements"] = measurements
                    backend["removed_page_boundary_spacers"] = spacers_removed
                    backend["core_section_start_adjustments"] = section_adjustments
                    backend["core_page_display_offsets"] = {
                        str(index): value for index, value in display_offsets.items()
                    }
                    backend["selective_writeback"] = writeback
                    verification_pdf = args.pdf_output or (
                        Path(refresh_name) / "read-only-verification.pdf"
                    )
                    backend["read_only_verification"] = external_verify(
                        args.output,
                        args.field_updater_command,
                        args.profile,
                        args.structure_map,
                        verification_pdf,
                        args.target_software
                        or profile["target_applications"][0],
                        expected_page_count=backend.get("page_count"),
                        allowed_field_types=allowed_field_types,
                    )
                delivery_status = "selective_verified"
            except FormatMonographError:
                if args.field_updater != "auto" or not args.approve_deferred:
                    raise
                backend = use_deferred_output(
                    args.input, args.output, "external_error"
                )
                delivery_status = "deferred"
        elif input_fields["status"] in {"absent", "refreshed"}:
            shutil.copy2(args.input, args.output)
        elif args.field_updater == "deferred":
            if not args.approve_deferred:
                raise FormatMonographError(
                    "Deferred field update requires caller QA and --approve-deferred."
                )
            shutil.copy2(args.input, args.output)
            rewrite_field_flags(args.output, deferred=True)
            backend = {"backend": "deferred_on_open"}
            delivery_status = "deferred"
        else:
            with tempfile.TemporaryDirectory(
                prefix="format-monograph-libreoffice-fields-"
            ) as refresh_name:
                refreshed = Path(refresh_name) / "refreshed.docx"
                try:
                    backend = libreoffice_refresh(
                        args.input, refreshed, args.renderer
                    )
                except FormatMonographError:
                    if args.field_updater != "auto" or not args.approve_deferred:
                        raise
                    backend = use_deferred_output(
                        args.input, args.output, "libreoffice_error"
                    )
                    delivery_status = "deferred"
                else:
                    try:
                        backend["selective_writeback"] = (
                            selective_field_result_writeback(
                                args.input,
                                refreshed,
                                args.output,
                                allowed_field_types=(
                                    set(DEFAULT_ALLOWED_FIELD_TYPES) | {"TC"}
                                    if approved_toc_contract is not None
                                    else DEFAULT_ALLOWED_FIELD_TYPES
                                ),
                                toc_contract=approved_toc_contract,
                            )
                        )
                        delivery_status = "selective_verified"
                    except FormatMonographError:
                        if args.field_updater != "auto" or not args.approve_deferred:
                            raise
                        backend = use_deferred_output(
                            args.input,
                            args.output,
                            "libreoffice_contract_or_integrity",
                        )
                        delivery_status = "deferred"

        output_fields = field_cache_inventory(args.output)
        strict_backend = backend.get("backend") not in {
            "not_needed",
            "deferred_on_open",
        }
        if strict_backend:
            selective_ok = (
                backend.get("selective_writeback", {}).get("status")
                == "selective_verified"
            )
            delivery_status = (
                "selective_verified" if selective_ok else output_fields["status"]
            )
            field_contract_ok = field_contract_preserved(input_fields, output_fields)
            refreshed_ok = (
                not input_fields["main_toc_fields"] or delivery_status == "refreshed"
            )
            if backend.get("backend") not in {"libreoffice_uno"}:
                refreshed_ok = bool(
                    backend.get("field_cache_verified")
                    and backend.get("read_only_verification", {}).get(
                        "read_only_verified"
                    )
                    and selective_ok
                    and backend.get("selective_writeback", {}).get(
                        "unapproved_dirty_fields", 0
                    )
                    == 0
                )
            elif selective_ok:
                refreshed_ok = True
        else:
            field_contract_ok = True
            refreshed_ok = True

        output_fp = structure_content_fingerprint(args.output, structure_map)
        output_objects = protected_payload_manifest(args.output)
        output_font_failures = effective_font_failures(
            args.output, profile, structure_map
        )
        content_ok = baseline_fp == output_fp
        objects_ok = baseline_objects == output_objects
        fonts_ok = not output_font_failures
        if strict_backend and not (
            field_contract_ok and refreshed_ok and content_ok and objects_ok and fonts_ok
        ):
            if args.field_updater == "auto" and args.approve_deferred:
                backend = use_deferred_output(
                    args.input, args.output, "libreoffice_contract_or_integrity"
                )
                delivery_status = "deferred"
                output_fields = field_cache_inventory(args.output)
                output_fp = structure_content_fingerprint(args.output, structure_map)
                output_objects = protected_payload_manifest(args.output)
                output_font_failures = effective_font_failures(
                    args.output, profile, structure_map
                )
                content_ok = baseline_fp == output_fp
                objects_ok = baseline_objects == output_objects
                fonts_ok = not output_font_failures
            else:
                args.output.unlink(missing_ok=True)
                raise FormatMonographError(
                    "Field refresh did not preserve the editable-field "
                    "contract and document integrity."
                )
        if not content_ok or not objects_ok or not fonts_ok:
            args.output.unlink(missing_ok=True)
            raise FormatMonographError(
                "Finalization integrity failed "
                f"(content={'pass' if content_ok else 'fail'}, "
                f"protected_objects={'pass' if objects_ok else 'fail'}, "
                f"effective_fonts={'pass' if fonts_ok else 'fail'})."
            )

        result = {
            "status": "pass",
            "delivery_field_status": delivery_status,
            "input_field_cache": input_fields,
            "output_field_cache": output_fields,
            "field_backend": backend,
            "field_writeback_status": (
                backend.get("selective_writeback", {}).get("status")
                or ("deferred" if delivery_status == "deferred" else "not_needed")
            ),
            "content_integrity": "pass",
            "protected_object_integrity": "pass",
            "effective_font_integrity": "pass",
            "workflow_state": {
                "source_sha256": file_sha256(baseline),
                "input_sha256": file_sha256(args.input),
                "profile_sha256": file_sha256(args.profile),
                "structure_map_sha256": file_sha256(args.structure_map),
                "output_sha256": file_sha256(args.output),
                "stage": "finalized",
            },
            "target_pdf": (
                str(args.pdf_output)
                if args.pdf_output and args.pdf_output.is_file()
                else None
            ),
            "target_layout_status": (
                "target_pdf_ready_for_visual_qa"
                if args.pdf_output and args.pdf_output.is_file()
                else "not_verified"
            ),
            "output": str(args.output),
        }
        if args.status_output:
            write_json(args.status_output, result)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (
        FormatMonographError,
        OSError,
        ValueError,
        subprocess.SubprocessError,
        zipfile.BadZipFile,
    ) as exc:
        args.output.unlink(missing_ok=True)
        if args.pdf_output:
            args.pdf_output.unlink(missing_ok=True)
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
