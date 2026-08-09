from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document

from test_v024_finalization import V024FinalizationTests

REPO = Path(__file__).resolve().parents[1]
SCRIPT = REPO / "format-monograph" / "scripts" / "render_docx.py"


def main() -> int:
    with tempfile.TemporaryDirectory() as temp_name:
        root = Path(temp_name)
        source = root / "render-smoke.docx"
        output = root / "pages"
        document = Document()
        document.add_heading("Render Smoke Test", level=1)
        document.add_paragraph("This synthetic page verifies LibreOffice and PyMuPDF.")
        document.save(source)

        completed = subprocess.run(
            [
                sys.executable,
                str(SCRIPT),
                str(source),
                "--output-dir",
                str(output),
            ],
            cwd=REPO / "format-monograph",
            capture_output=True,
            text=True,
            check=False,
        )
        if completed.returncode != 0:
            print(completed.stdout)
            print(completed.stderr, file=sys.stderr)
            return completed.returncode

        manifest = json.loads((output / "render-manifest.json").read_text(encoding="utf-8"))
        pages = [Path(path) for path in manifest["pages"]]
        if manifest["page_count"] < 1 or not pages:
            print("Renderer produced no pages.", file=sys.stderr)
            return 1
        if any(not page.is_file() or page.stat().st_size == 0 for page in pages):
            print("Renderer produced an empty or missing PNG.", file=sys.stderr)
            return 1

        case = V024FinalizationTests(
            methodName="test_deferred_finalization_requires_explicit_qa"
        )
        case.setUp()
        try:
            formatted = case.apply()
            finalized = case.root / "field-finalized.docx"
            status_path = case.root / "field-finalization.json"
            finalized_result = subprocess.run(
                [
                    sys.executable,
                    str(REPO / "format-monograph" / "scripts" / "finalize_docx.py"),
                    str(formatted),
                    "--source",
                    str(case.source),
                    "--profile",
                    str(case.profile),
                    "--structure-map",
                    str(case.structure),
                    "--output",
                    str(finalized),
                    "--status-output",
                    str(status_path),
                    "--field-updater",
                    "auto",
                    "--approve-deferred",
                ],
                cwd=REPO / "format-monograph",
                capture_output=True,
                text=True,
                check=False,
            )
            if finalized_result.returncode != 0:
                print(finalized_result.stdout)
                print(finalized_result.stderr, file=sys.stderr)
                return finalized_result.returncode
            finalization = json.loads(status_path.read_text(encoding="utf-8"))
            backend = finalization["field_backend"]
            accepted_backend = backend["backend"] == "libreoffice_uno" or (
                backend["backend"] == "deferred_on_open"
                and backend.get("fallback_from")
                == "libreoffice_contract_or_integrity"
            )
            if not accepted_backend:
                print(f"LibreOffice field backend was not exercised: {backend}", file=sys.stderr)
                return 1
            if (
                finalization["content_integrity"] != "pass"
                or finalization["protected_object_integrity"] != "pass"
            ):
                print("Finalization integrity did not pass.", file=sys.stderr)
                return 1
        finally:
            case.tearDown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
