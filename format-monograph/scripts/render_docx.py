#!/usr/bin/env python3
"""Render a DOCX to per-page PNG images using LibreOffice and PyMuPDF."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from _common import FormatMonographError, ensure_docx, write_json
from target_software import LIBREOFFICE, resolve_target_id
from libreoffice_runtime import default_macos_soffice


def locate_soffice(requested: str | None = None) -> tuple[str, str]:
    source = "argument"
    executable = requested
    if not executable:
        source = "environment"
        executable = os.environ.get("FORMAT_MONOGRAPH_RENDERER")
    if not executable:
        source = "application"
        executable = default_macos_soffice()
    if not executable:
        source = "path"
        executable = shutil.which("soffice") or shutil.which("libreoffice")
    if not executable:
        raise FormatMonographError("LibreOffice soffice was not found.")
    path = Path(executable).expanduser()
    if not path.is_file() and source != "path":
        raise FormatMonographError(f"Configured renderer does not exist: {path}")
    return str(path.resolve()) if path.is_file() else str(executable), source


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--keep-pdf", action="store_true")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--renderer", help="Explicit LibreOffice-compatible renderer path.")
    parser.add_argument(
        "--target-pdf",
        type=Path,
        help="Render an already exported target-software PDF instead of converting with LibreOffice.",
    )
    parser.add_argument(
        "--target-software",
        help="Publisher-designated target application used to qualify visual QA.",
    )
    args = parser.parse_args()

    try:
        ensure_docx(args.input)
        if args.dpi < 72 or args.dpi > 300:
            raise FormatMonographError("--dpi must be between 72 and 300.")
        if args.target_pdf:
            if not args.target_pdf.is_file() or args.target_pdf.suffix.lower() != ".pdf":
                raise FormatMonographError("--target-pdf must name an existing PDF file.")
            soffice, renderer_source, renderer_version = None, "target_pdf", None
        else:
            soffice, renderer_source = locate_soffice(args.renderer)
            version_result = subprocess.run(
                [soffice, "--version"],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
            renderer_version = (version_result.stdout or version_result.stderr).strip() or None
        try:
            import fitz
        except ImportError as exc:
            raise FormatMonographError("PyMuPDF is required for page rendering.") from exc

        args.output_dir.mkdir(parents=True, exist_ok=True)
        manifest_path = args.output_dir / "render-manifest.json"
        existing_pages = sorted(args.output_dir.glob("page-*.png"))
        if (existing_pages or manifest_path.exists()) and not args.force:
            raise FormatMonographError(
                "Render output already exists; use --force or choose a new output directory."
            )
        if args.force:
            for page in existing_pages:
                page.unlink()
            manifest_path.unlink(missing_ok=True)

        with tempfile.TemporaryDirectory(prefix="format-monograph-render-") as temp_name:
            temp_dir = Path(temp_name)
            profile_dir = temp_dir / "lo-profile"
            conversion_dir = temp_dir / "converted"
            profile_dir.mkdir()
            conversion_dir.mkdir()
            env = os.environ.copy()
            env["HOME"] = str(temp_dir / "home")
            Path(env["HOME"]).mkdir()

            if args.target_pdf:
                pdf_path = args.target_pdf.resolve()
            else:
                command = [
                    soffice,
                    "--headless",
                    f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(conversion_dir),
                    str(args.input.resolve()),
                ]
                completed = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    env=env,
                    timeout=300,
                    check=False,
                )
                pdf_path = conversion_dir / f"{args.input.stem}.pdf"
                if completed.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size == 0:
                    raise FormatMonographError(
                        "LibreOffice conversion failed. "
                        f"stdout={completed.stdout.strip()} stderr={completed.stderr.strip()}"
                    )

            pdf = fitz.open(pdf_path)
            page_paths = []
            scale = args.dpi / 72.0
            matrix = fitz.Matrix(scale, scale)
            for index, page in enumerate(pdf):
                output = args.output_dir / f"page-{index + 1:04d}.png"
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                pixmap.save(output)
                page_paths.append(str(output))
            pdf.close()

            kept_pdf = None
            if args.keep_pdf:
                kept_pdf_path = args.output_dir / f"{args.input.stem}.pdf"
                if pdf_path.resolve() != kept_pdf_path.resolve():
                    shutil.copy2(pdf_path, kept_pdf_path)
                kept_pdf = str(kept_pdf_path)

        target_id = resolve_target_id(args.target_software)
        result = {
            "input": str(args.input.resolve()),
            "page_count": len(page_paths),
            "dpi": args.dpi,
            "pages": page_paths,
            "pdf": kept_pdf,
            "renderer": soffice,
            "renderer_version": renderer_version,
            "renderer_source": renderer_source,
            "target_software": target_id,
            "target_pdf_source": str(args.target_pdf.resolve()) if args.target_pdf else None,
            "target_layout_unverified": bool(
                args.target_software
                and not args.target_pdf
                and target_id != LIBREOFFICE
            ),
            "visual_review": "pending",
        }
        write_json(manifest_path, result)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (FormatMonographError, OSError, subprocess.SubprocessError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
