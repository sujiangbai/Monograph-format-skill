#!/usr/bin/env python3
"""Real macOS smoke for exact one-index authorization and section stability."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import FormatMonographError  # noqa: E402
from finalize_docx import (  # noqa: E402
    libreoffice_macro_refresh,
    restore_known_libreoffice_toc_instruction_order,
    section_pagination_differences,
    toc_index_authorization,
)
from libreoffice_runtime import macos_internal_macro_soffice  # noqa: E402
from render_docx import locate_soffice  # noqa: E402
from toc_index_identity import ooxml_toc_identities  # noqa: E402


FIXTURE_SCRIPT_URI = (
    "vnd.sun.star.script:libreoffice_toc_fixture_macro.py$"
    "create_toc_fixture?language=Python&location=user"
)


def generate_fixture(soffice: str, root: Path, output: Path) -> None:
    profile = root / "generator-profile"
    script_dir = profile / "user" / "Scripts" / "python"
    script_dir.mkdir(parents=True)
    helper = Path(__file__).with_name("libreoffice_toc_fixture_macro.py")
    shutil.copy2(helper, script_dir / helper.name)
    result = root / "generator-result.json"
    env = os.environ.copy()
    env.update(
        {
            "FORMAT_MONOGRAPH_TOC_FIXTURE_OUTPUT": str(output),
            "FORMAT_MONOGRAPH_TOC_FIXTURE_RESULT": str(result),
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
            command, stdout=stdout_log, stderr=stderr_log, text=True, env=env
        )
        deadline = time.monotonic() + 120
        try:
            while time.monotonic() < deadline and not result.is_file():
                if process.poll() is not None:
                    break
                time.sleep(0.1)
            if not result.is_file():
                raise FormatMonographError("TOC fixture generator returned no result.")
            details = json.loads(result.read_text(encoding="utf-8"))
            if not details.get("ok") or not output.is_file():
                raise FormatMonographError(
                    "TOC fixture generation failed: "
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


def document_root(path: Path):
    with zipfile.ZipFile(path) as package:
        return etree.fromstring(package.read("word/document.xml"))


def add_exact_baseline_hide_web_switch(path: Path) -> None:
    temporary = path.with_suffix(".with-z.docx")
    namespace = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }
    changed = 0
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                for instruction in root.xpath(".//w:instrText", namespaces=namespace):
                    if 'TOC \\z \\o "1-3" \\h' in (instruction.text or ""):
                        instruction.text = (instruction.text or "").replace(
                            'TOC \\z \\o "1-3" \\h',
                            'TOC \\o "1-3" \\h \\z',
                        )
                        changed += 1
                    elif "TOC \\o \"1-3\" \\h" in (instruction.text or ""):
                        instruction.text = (instruction.text or "").replace(
                            'TOC \\o "1-3" \\h', 'TOC \\o "1-3" \\h \\z'
                        )
                        changed += 1
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
            target.writestr(info, data)
    if changed != 1:
        temporary.unlink(missing_ok=True)
        raise FormatMonographError("Synthetic fixture has no unique TOC instruction.")
    temporary.replace(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--renderer", required=True)
    args = parser.parse_args()
    soffice, source = locate_soffice(args.renderer)
    macro_soffice = macos_internal_macro_soffice(soffice)
    if macro_soffice is None:
        print("The index authorization smoke requires macOS LibreOffice.", file=sys.stderr)
        return 2
    temp_root: Path | None = None
    report: dict[str, object] = {"status": "fail"}
    with tempfile.TemporaryDirectory(
        prefix="format-monograph-index-smoke-"
    ) as temp_name:
        temp_root = Path(temp_name)
        seed = temp_root / "seed.docx"
        baseline = temp_root / "baseline.docx"
        refreshed = temp_root / "refreshed.docx"
        restored = temp_root / "restored.docx"
        generate_fixture(macro_soffice, temp_root, seed)
        add_exact_baseline_hide_web_switch(seed)
        seed_toc_identities = ooxml_toc_identities(seed)
        contract = [
            {"level": 1, "kind": "heading", "text_sha256": "a" * 64}
        ]
        warmup_authorization = toc_index_authorization(contract, seed)
        warmup_backend = libreoffice_macro_refresh(
            seed,
            baseline,
            macro_soffice,
            source,
            toc_authorization=warmup_authorization,
            toc_contract=contract,
        )
        add_exact_baseline_hide_web_switch(baseline)
        authorization = toc_index_authorization(contract, baseline)
        backend = libreoffice_macro_refresh(
            baseline,
            refreshed,
            macro_soffice,
            source,
            toc_authorization=authorization,
            toc_contract=contract,
        )
        section_differences = section_pagination_differences(
            document_root(baseline), document_root(refreshed)
        )
        restoration = restore_known_libreoffice_toc_instruction_order(
            baseline, refreshed, restored
        )
        report = {
            "status": "pass",
            "backend": backend.get("backend"),
            "warmup_approved_indexes_updated": warmup_backend.get(
                "approved_indexes_updated"
            ),
            "document_index_count": backend.get("document_index_count"),
            "approved_indexes_updated": backend.get("approved_indexes_updated"),
            "skipped_indexes": backend.get("skipped_indexes"),
            "expected_authorization": authorization,
            "reported_authorization_id": backend.get("toc_authorization_id"),
            "independent_contract_sha256": backend.get("toc_contract_sha256"),
            "observed_index_descriptor": backend.get(
                "observed_index_descriptor"
            ),
            "seed_toc_identities": seed_toc_identities,
            "section_differences": section_differences,
            "pagination_semantics_identical": not section_differences,
            "instruction_restoration": restoration,
        }
        if (
            report["backend"] != "libreoffice_uno"
            or report["document_index_count"] != 1
            or report["approved_indexes_updated"] != 1
            or report["skipped_indexes"] != 0
            or report["reported_authorization_id"]
            != authorization["authorization_id"]
            or report["independent_contract_sha256"]
            != authorization["structure_contract_sha256"]
            or report["observed_index_descriptor"]
            != {
                key: value
                for key, value in authorization.items()
                if key != "authorization_id"
            }
            or section_differences
            or restoration.get("field_contract_identical") is not True
        ):
            report["status"] = "fail"
    report["temporary_directory_remaining"] = bool(
        temp_root is not None and temp_root.exists()
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
