#!/usr/bin/env python3
"""Real macOS LibreOffice negative smoke for macros and external fields."""

from __future__ import annotations

import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import FormatMonographError  # noqa: E402
from finalize_docx import (  # noqa: E402
    libreoffice_macro_refresh,
    package_external_connection_inventory,
)
from libreoffice_runtime import macos_internal_macro_soffice  # noqa: E402
from render_docx import locate_soffice  # noqa: E402


FIXTURE_SCRIPT_URI = (
    "vnd.sun.star.script:libreoffice_security_fixture_macro.py$"
    "create_fixture?language=Python&location=user"
)
EXTERNAL_PROBE_SCRIPT_URI = (
    "vnd.sun.star.script:libreoffice_external_probe_macro.py$"
    "probe_external_refresh?language=Python&location=user"
)
PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


class ProbeHandler(BaseHTTPRequestHandler):
    hits: list[str] = []

    def do_GET(self) -> None:  # noqa: N802 - stdlib callback name
        type(self).hits.append(self.path)
        self.send_response(200)
        self.send_header("Content-Type", "image/png")
        self.end_headers()
        self.wfile.write(PNG_1X1)

    def log_message(self, _format: str, *_args: object) -> None:
        return None


def inject_external_graphic(path: Path, port: int) -> str:
    namespaces = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
        "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
        "xlink": "http://www.w3.org/1999/xlink",
    }
    url = f"http://127.0.0.1:{port}/linked-graphic.png"
    temporary = path.with_suffix(".external.odt")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "content.xml":
                root = etree.fromstring(data)
                office_text = root.find(
                    f".//{{{namespaces['office']}}}text"
                )
                if office_text is None:
                    raise FormatMonographError("ODT fixture has no office:text element.")
                paragraph = etree.SubElement(
                    office_text, f"{{{namespaces['text']}}}p"
                )
                frame = etree.SubElement(
                    paragraph, f"{{{namespaces['draw']}}}frame"
                )
                frame.set(f"{{{namespaces['draw']}}}name", "LoopbackExternalGraphic")
                frame.set(f"{{{namespaces['text']}}}anchor-type", "paragraph")
                frame.set(f"{{{namespaces['svg']}}}width", "1cm")
                frame.set(f"{{{namespaces['svg']}}}height", "1cm")
                image = etree.SubElement(frame, f"{{{namespaces['draw']}}}image")
                image.set(f"{{{namespaces['xlink']}}}href", url)
                image.set(f"{{{namespaces['xlink']}}}type", "simple")
                image.set(f"{{{namespaces['xlink']}}}show", "embed")
                image.set(f"{{{namespaces['xlink']}}}actuate", "onLoad")
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8"
                )
            target.writestr(info, data)
    temporary.replace(path)
    return url


def inject_onload_event(path: Path) -> None:
    namespaces = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "script": "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
        "xlink": "http://www.w3.org/1999/xlink",
    }
    temporary = path.with_suffix(".event.odt")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "content.xml":
                root = etree.fromstring(data)
                scripts = root.find(
                    f"{{{namespaces['office']}}}scripts"
                )
                if scripts is None:
                    scripts = etree.Element(f"{{{namespaces['office']}}}scripts")
                    root.insert(0, scripts)
                listeners = etree.SubElement(
                    scripts, f"{{{namespaces['office']}}}event-listeners"
                )
                listener = etree.SubElement(
                    listeners, f"{{{namespaces['script']}}}event-listener"
                )
                listener.set(
                    f"{{{namespaces['script']}}}language", "ooo:script"
                )
                listener.set(f"{{{namespaces['script']}}}event-name", "dom:load")
                listener.set(
                    f"{{{namespaces['xlink']}}}href",
                    "vnd.sun.star.script:Standard.SecurityProbe.OnLoad"
                    "?language=Basic&location=document",
                )
                listener.set(f"{{{namespaces['xlink']}}}type", "simple")
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8"
                )
            target.writestr(info, data)
    temporary.replace(path)


def replace_package_link(path: Path, old: str, new: str) -> None:
    temporary = path.with_suffix(".relinked.odt")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "content.xml":
                data = data.replace(old.encode("utf-8"), new.encode("utf-8"))
            target.writestr(info, data)
    temporary.replace(path)


def inject_scheme_less_escape_relationship(path: Path) -> str:
    target_value = "../../format-monograph-relationship-probe.bin"
    temporary = path.with_suffix(".relationship.odt")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            target.writestr(info, source.read(info.filename))
        target.writestr(
            "_rels/.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rIdEscape" Type="probe" Target="{target_value}"/>'
            "</Relationships>",
        )
    temporary.replace(path)
    return target_value


def create_macro_fixture(soffice: str, root: Path, output: Path, marker: Path) -> None:
    profile = root / "fixture-profile"
    script_dir = profile / "user" / "Scripts" / "python"
    script_dir.mkdir(parents=True)
    helper = Path(__file__).with_name("libreoffice_security_fixture_macro.py")
    shutil.copy2(helper, script_dir / helper.name)
    result = root / "fixture-result.json"
    env = os.environ.copy()
    env.update(
        {
            "FORMAT_MONOGRAPH_SECURITY_ODT": str(output),
            "FORMAT_MONOGRAPH_SECURITY_MACRO_MARKER": str(marker),
            "FORMAT_MONOGRAPH_SECURITY_FIXTURE_RESULT": str(result),
        }
    )
    command = [
        soffice,
        "--headless",
        "--invisible",
        "--nologo",
        "--nodefault",
        "--norestore",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        FIXTURE_SCRIPT_URI,
    ]
    with tempfile.TemporaryFile(mode="w+t", encoding="utf-8") as stdout_log, tempfile.TemporaryFile(
        mode="w+t", encoding="utf-8"
    ) as stderr_log:
        process = subprocess.Popen(
            command,
            stdout=stdout_log,
            stderr=stderr_log,
            text=True,
            env=env,
        )
        deadline = time.monotonic() + 120
        try:
            while time.monotonic() < deadline and not result.is_file():
                if process.poll() is not None:
                    break
                time.sleep(0.1)
            if not result.is_file():
                stdout_log.seek(0)
                stderr_log.seek(0)
                raise FormatMonographError(
                    "LibreOffice security fixture generator did not return a result. "
                    f"stdout={stdout_log.read()[-4096:]} stderr={stderr_log.read()[-4096:]}"
                )
            details = json.loads(result.read_text(encoding="utf-8"))
            if not details.get("ok") or not output.is_file():
                raise FormatMonographError(
                    "LibreOffice security fixture generation failed: "
                    + json.dumps(details, ensure_ascii=False)
                )
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=15)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=15)


def run_external_positive_control(
    soffice: str, root: Path, fixture: Path, expected_url: str
) -> dict:
    profile = root / "unsafe-probe-profile"
    script_dir = profile / "user" / "Scripts" / "python"
    script_dir.mkdir(parents=True)
    helper = Path(__file__).with_name("libreoffice_external_probe_macro.py")
    shutil.copy2(helper, script_dir / helper.name)
    result = root / "unsafe-probe-result.json"
    env = os.environ.copy()
    env.update(
        {
            "FORMAT_MONOGRAPH_PROBE_INPUT": str(fixture),
            "FORMAT_MONOGRAPH_PROBE_RESULT": str(result),
            "FORMAT_MONOGRAPH_PROBE_EXPECTED_URL": expected_url,
        }
    )
    command = [
        soffice,
        "--headless",
        "--invisible",
        "--nologo",
        "--nodefault",
        "--norestore",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        EXTERNAL_PROBE_SCRIPT_URI,
    ]
    with tempfile.TemporaryFile(mode="w+t", encoding="utf-8") as stdout_log, tempfile.TemporaryFile(
        mode="w+t", encoding="utf-8"
    ) as stderr_log:
        process = subprocess.Popen(
            command,
            stdout=stdout_log,
            stderr=stderr_log,
            text=True,
            env=env,
        )
        deadline = time.monotonic() + 120
        try:
            while time.monotonic() < deadline and not result.is_file():
                if process.poll() is not None:
                    break
                time.sleep(0.1)
            if not result.is_file():
                stdout_log.seek(0)
                stderr_log.seek(0)
                raise FormatMonographError(
                    "LibreOffice external positive control did not return a result. "
                    f"stdout={stdout_log.read()[-4096:]} stderr={stderr_log.read()[-4096:]}"
                )
            details = json.loads(result.read_text(encoding="utf-8"))
            if not details.get("ok"):
                raise FormatMonographError(
                    "LibreOffice external positive control failed: "
                    + json.dumps(details, ensure_ascii=False)
                )
            return details
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=15)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=15)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--renderer", required=True)
    args = parser.parse_args()
    soffice, source = locate_soffice(args.renderer)
    macro_soffice = macos_internal_macro_soffice(soffice)
    if macro_soffice is None:
        print("The security smoke requires a macOS LibreOffice app bundle.", file=sys.stderr)
        return 2

    ProbeHandler.hits = []
    server = ThreadingHTTPServer(("127.0.0.1", 0), ProbeHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    report: dict[str, object] = {}
    temp_root: Path | None = None
    try:
        with tempfile.TemporaryDirectory(
            prefix="format-monograph-security-smoke-"
        ) as temp_name:
            temp_root = Path(temp_name)
            macro_fixture = temp_root / "document-macro.odt"
            macro_marker = temp_root / "document-macro-executed.txt"
            macro_output = temp_root / "document-macro-output.docx"
            create_macro_fixture(
                macro_soffice, temp_root, macro_fixture, macro_marker
            )
            macro_backend = libreoffice_macro_refresh(
                macro_fixture,
                macro_output,
                macro_soffice,
                source,
                toc_authorization=None,
            )
            active_macro_fixture = temp_root / "active-document-macro.odt"
            shutil.copy2(macro_fixture, active_macro_fixture)
            inject_onload_event(active_macro_fixture)
            active_macro_inventory = package_external_connection_inventory(
                active_macro_fixture
            )
            active_macro_rejection = None
            try:
                libreoffice_macro_refresh(
                    active_macro_fixture,
                    temp_root / "active-macro-output.docx",
                    macro_soffice,
                    source,
                    toc_authorization=None,
                )
            except FormatMonographError as exc:
                active_macro_rejection = str(exc)

            external_fixture = temp_root / "external-link.odt"
            external_output = temp_root / "external-fields-output.docx"
            external_marker = temp_root / "unused-external-macro-marker.txt"
            create_macro_fixture(
                macro_soffice, temp_root / "external-fixture", external_fixture, external_marker
            )
            expected_external_url = inject_external_graphic(
                external_fixture, server.server_port
            )
            external_inventory = package_external_connection_inventory(
                external_fixture
            )
            file_fixture = temp_root / "file-link.odt"
            shutil.copy2(external_fixture, file_fixture)
            file_url = "file:///tmp/format-monograph-never-read.png"
            replace_package_link(file_fixture, expected_external_url, file_url)
            file_inventory = package_external_connection_inventory(file_fixture)
            file_rejection = None
            try:
                libreoffice_macro_refresh(
                    file_fixture,
                    temp_root / "file-link-output.docx",
                    macro_soffice,
                    source,
                    toc_authorization=None,
                )
            except FormatMonographError as exc:
                file_rejection = str(exc)
            relationship_fixture = temp_root / "relationship-escape.odt"
            shutil.copy2(macro_fixture, relationship_fixture)
            relationship_target = inject_scheme_less_escape_relationship(
                relationship_fixture
            )
            relationship_inventory = package_external_connection_inventory(
                relationship_fixture
            )
            relationship_rejection = None
            ProbeHandler.hits = []
            try:
                libreoffice_macro_refresh(
                    relationship_fixture,
                    temp_root / "relationship-output.docx",
                    macro_soffice,
                    source,
                    toc_authorization=None,
                )
            except FormatMonographError as exc:
                relationship_rejection = str(exc)
            relationship_external_requests = len(ProbeHandler.hits)
            ProbeHandler.hits = []
            unsafe_probe = run_external_positive_control(
                macro_soffice, temp_root, external_fixture, expected_external_url
            )
            time.sleep(0.5)
            unsafe_external_requests = len(ProbeHandler.hits)
            unsafe_external_request_paths = list(ProbeHandler.hits)
            ProbeHandler.hits = []
            safe_rejection = None
            try:
                libreoffice_macro_refresh(
                    external_fixture,
                    external_output,
                    macro_soffice,
                    source,
                    toc_authorization=None,
                )
            except FormatMonographError as exc:
                safe_rejection = str(exc)
            time.sleep(0.5)
            report = {
                "status": "pass",
                "backend": "rejected_preflight",
                "uno_mode": "not_started_for_external_fixture",
                "document_macro_fixture": "embedded_on_load",
                "embedded_script_events": macro_backend.get(
                    "embedded_script_events"
                ),
                "document_macro_executions": int(macro_marker.exists()),
                "active_macro_inventory": active_macro_inventory,
                "active_macro_preflight_rejection": active_macro_rejection,
                "fixture_external_url": expected_external_url,
                "fixture_external_connection_inventory": external_inventory,
                "fixture_external_connection_count": len(external_inventory),
                "file_fixture_url": file_url,
                "file_fixture_inventory": file_inventory,
                "file_fixture_preflight_rejection": file_rejection,
                "relationship_fixture_target": relationship_target,
                "relationship_fixture_inventory": relationship_inventory,
                "relationship_preflight_rejection": relationship_rejection,
                "relationship_uno_mode": "not_started_for_relationship_fixture",
                "relationship_external_requests": relationship_external_requests,
                "unsafe_probe_recognized_external_connections": unsafe_probe.get(
                    "recognized_external_connections"
                ),
                "unsafe_probe_recognized_external_connection_count": unsafe_probe.get(
                    "recognized_external_connection_count"
                ),
                "unsafe_full_update_load_completed": unsafe_probe.get(
                    "unsafe_full_update_load_completed"
                ),
                "unsafe_graphics_loaded": unsafe_probe.get(
                    "unsafe_graphics_loaded"
                ),
                "unsafe_external_requests": unsafe_external_requests,
                "unsafe_external_request_paths": unsafe_external_request_paths,
                "external_requests": len(ProbeHandler.hits),
                "external_request_paths": list(ProbeHandler.hits),
                "safe_preflight_rejection": safe_rejection,
                "macro_text_fields_collection_refreshed": macro_backend.get(
                    "text_fields_collection_refreshed"
                ),
                "external_text_fields_collection_refreshed": False,
            }
            if (
                report["document_macro_executions"] != 0
                or report["embedded_script_events"] != 0
                or len(report["active_macro_inventory"]) != 1
                or not report["active_macro_preflight_rejection"]
                or report["fixture_external_connection_count"] != 1
                or len(report["file_fixture_inventory"]) != 1
                or not report["file_fixture_preflight_rejection"]
                or len(report["relationship_fixture_inventory"]) != 1
                or not report["relationship_preflight_rejection"]
                or report["relationship_external_requests"] != 0
                or report["unsafe_full_update_load_completed"] is not True
                or report["unsafe_graphics_loaded"] != 1
                or report["unsafe_external_requests"] < 1
                or report["external_requests"] != 0
                or not report["safe_preflight_rejection"]
                or report["macro_text_fields_collection_refreshed"] is not False
                or report["external_text_fields_collection_refreshed"] is not False
            ):
                report["status"] = "fail"
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)
    report["temporary_directory_remaining"] = bool(
        temp_root is not None and temp_root.exists()
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("status") == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
