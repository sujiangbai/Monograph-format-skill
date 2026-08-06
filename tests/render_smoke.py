from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document

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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
