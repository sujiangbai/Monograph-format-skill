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
from docx.oxml.ns import qn

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
)
from structure_map import (
    load_structure_map,
    structure_content_fingerprint,
    validate_structure_map_source,
)
from validate_profile import validate


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


FIELD_RESULT_PART = re.compile(
    r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$"
)


def _normalize_complex_toc_results(root: etree._Element) -> None:
    """Replace each complete TOC field span with one comparison-only marker."""
    stack: list[dict[str, Any]] = []
    toc_fields: list[dict[str, Any]] = []
    for element in root.iter():
        if element.tag == qn("w:fldChar"):
            kind = element.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"begin": element, "parts": [], "separate": None})
            elif kind == "separate" and stack:
                stack[-1]["separate"] = element
            elif kind == "end" and stack:
                field = stack.pop()
                instruction = " ".join("".join(field["parts"]).split()).upper()
                if instruction.startswith("TOC ") and field["separate"] is not None:
                    field.update({"end": element, "instruction": instruction})
                    toc_fields.append(field)
        elif element.tag == qn("w:instrText") and stack:
            if stack[-1]["separate"] is None:
                stack[-1]["parts"].append(element.text or "")

    bodies = root.xpath("/w:document/w:body", namespaces=NS)
    if not bodies:
        return
    body = bodies[0]

    def body_child(element: etree._Element) -> etree._Element | None:
        current = element
        while current.getparent() is not None and current.getparent() is not body:
            current = current.getparent()
        return current if current.getparent() is body else None

    spans: list[tuple[int, int, str, dict[str, Any]]] = []
    for field in toc_fields:
        start_block = body_child(field["begin"])
        end_block = body_child(field["end"])
        if start_block is None or end_block is None:
            raise FormatMonographError(
                "TOC field results outside the main document body are unsupported."
            )
        start = body.index(start_block)
        end = body.index(end_block)
        if start > end:
            raise FormatMonographError("TOC field boundaries are out of order.")

        inside = False
        outside_payload = False
        for block in list(body)[start : end + 1]:
            for element in block.iter():
                if element is field["begin"]:
                    inside = True
                    continue
                if element is field["end"]:
                    inside = False
                    continue
                if not inside and element.tag in {
                    qn("w:t"),
                    qn("w:delText"),
                    qn("w:drawing"),
                    qn("w:object"),
                    qn("w:pict"),
                }:
                    if element.tag not in {qn("w:t"), qn("w:delText")} or (
                        element.text or ""
                    ):
                        outside_payload = True
                        break
            if outside_payload:
                break
        if outside_payload:
            raise FormatMonographError(
                "A TOC field shares its boundary block with authored content."
            )
        spans.append((start, end, field["instruction"], field))

    for start, end, instruction, _ in sorted(
        spans, key=lambda item: item[0], reverse=True
    ):
        for index in range(end, start - 1, -1):
            body.remove(body[index])
        marker = etree.Element("{urn:format-monograph:audit}field-result")
        marker.set("instruction", instruction)
        body.insert(start, marker)


def _neutralized_field_result_sha256(data: bytes) -> str:
    """Hash OOXML while ignoring only cached field-result text."""
    root = etree.fromstring(data)
    _normalize_complex_toc_results(root)
    for field in root.xpath(".//w:fldSimple", namespaces=NS):
        for element in field.xpath(".//w:t | .//w:delText", namespaces=NS):
            element.text = ""

    stack: list[bool] = []
    for element in root.iter():
        if element.tag == qn("w:fldChar"):
            kind = element.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append(False)
            elif kind == "separate" and stack:
                stack[-1] = True
            elif kind == "end" and stack:
                stack.pop()
                continue
        if any(stack) and element.tag in {
            qn("w:t"),
            qn("w:delText"),
        }:
            element.text = ""
    canonical = etree.tostring(root, method="c14n", exclusive=True)
    return hashlib.sha256(canonical).hexdigest()


def controlled_field_result_writeback(
    baseline_path: Path, refreshed_path: Path, output_path: Path
) -> list[str]:
    """Copy verified field caches into the safe package, never refreshed payloads."""
    replaced: list[str] = []
    temp_output = output_path.with_name(
        f".{output_path.stem}.field-writeback.tmp.docx"
    )
    temp_output.unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(baseline_path) as baseline, zipfile.ZipFile(
            refreshed_path
        ) as refreshed, zipfile.ZipFile(
            temp_output, "w", compression=zipfile.ZIP_DEFLATED
        ) as output:
            refreshed_names = set(refreshed.namelist())
            for info in baseline.infolist():
                data = baseline.read(info.filename)
                if (
                    FIELD_RESULT_PART.match(info.filename)
                    and info.filename in refreshed_names
                ):
                    candidate = refreshed.read(info.filename)
                    if candidate != data:
                        if _neutralized_field_result_sha256(
                            data
                        ) != _neutralized_field_result_sha256(candidate):
                            raise FormatMonographError(
                                "Target application changed OOXML outside approved field results: "
                                + info.filename
                            )
                        data = candidate
                        replaced.append(info.filename)
                output.writestr(info, data)
        if protected_payload_manifest(temp_output) != protected_payload_manifest(
            baseline_path
        ):
            raise FormatMonographError(
                "Controlled field writeback changed a protected payload."
            )
        temp_output.replace(output_path)
    finally:
        temp_output.unlink(missing_ok=True)
    return replaced


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
) -> dict:
    request = {
        "protocol_version": "1.0",
        "input_path": str(input_path.resolve()),
        "output_path": str(output_path.resolve()),
        "profile_path": str(profile_path.resolve()),
        "structure_map_path": str(structure_map_path.resolve()),
        "allowed_field_types": ["TOC", "PAGE", "REF", "PAGEREF"],
        "target_software": target_software,
        "pdf_output_path": str(pdf_output.resolve()) if pdf_output else None,
    }
    completed = subprocess.run(
        _external_command(command),
        input=json.dumps(request, ensure_ascii=False),
        capture_output=True,
        text=True,
        timeout=600,
        check=False,
    )
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
                    refreshed = Path(refresh_name) / "refreshed.docx"
                    backend = external_refresh(
                        args.input,
                        refreshed,
                        args.field_updater_command,
                        args.profile,
                        args.structure_map,
                        args.pdf_output,
                        args.target_software or profile["target_applications"][0],
                    )
                    backend["controlled_writeback_parts"] = (
                        controlled_field_result_writeback(
                            args.input, refreshed, args.output
                        )
                    )
                delivery_status = (
                    "refreshed_target_word"
                    if "microsoft word" in str(backend.get("software", "")).casefold()
                    else "refreshed_external"
                )
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
            try:
                with tempfile.TemporaryDirectory(
                    prefix="format-monograph-libreoffice-fields-"
                ) as refresh_name:
                    refreshed = Path(refresh_name) / "refreshed.docx"
                    backend = libreoffice_refresh(
                        args.input, refreshed, args.renderer
                    )
                    backend["controlled_writeback_parts"] = (
                        controlled_field_result_writeback(
                            args.input, refreshed, args.output
                        )
                    )
            except FormatMonographError:
                if args.field_updater != "auto" or not args.approve_deferred:
                    raise
                backend = use_deferred_output(
                    args.input, args.output, "libreoffice_error"
                )
                delivery_status = "deferred"

        output_fields = field_cache_inventory(args.output)
        strict_backend = backend.get("backend") not in {
            "not_needed",
            "deferred_on_open",
        }
        if strict_backend:
            delivery_status = output_fields["status"]
            if backend.get("backend") not in {"libreoffice_uno"}:
                delivery_status = (
                    "refreshed_target_word"
                    if "microsoft word" in str(backend.get("software", "")).casefold()
                    else "refreshed_external"
                )
            field_contract_ok = field_contract_preserved(input_fields, output_fields)
            refreshed_ok = (
                not input_fields["main_toc_fields"] or delivery_status == "refreshed"
            )
            if backend.get("backend") not in {"libreoffice_uno"}:
                refreshed_ok = bool(backend.get("field_cache_verified"))
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
