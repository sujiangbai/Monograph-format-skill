#!/usr/bin/env python3
"""Opt-in read-only/PDF check of only the accepted 010 live-05 candidate."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run_field_probe as field
from pypdf import PdfReader

EVIDENCE = field.ROOT / "artifacts/probe-evidence"
CALCULATION_RESULT = EVIDENCE / "field-010-live-05/result.json"
CALCULATION_RESULT_SHA256 = "70eb8c837e082ab260b520fef7736b84c001e5508e2973cbe09601d1940761b0"
CALCULATION_RESULT_SIZE = 7269
CANDIDATE = EVIDENCE / "field-010-live-05/v051-synthetic-selective-candidate.docx"
CANDIDATE_SHA256 = "338ffb3dedc293f516bf28e088cdd6f0b864bb324254412e50e904ea33795ff6"


def build_applescript() -> str:
    ownership = Path(__file__).with_name("readonly_settings_probe.applescript").read_text()
    helpers, body = ownership.split("on run argv\n", 1)
    readers = Path(__file__).with_name("field_probe.applescript").read_text().split("on calculateFields(doc)\n", 1)[0]
    readers = readers.replace("__BODY_CODES__", field.apple_literal(field.EXPECTED_BODY_CODES))
    body = body.replace('(count of argv) is not 1', '(count of argv) is not 2')
    body = body.replace('Exactly one generated synthetic DOCX path required', 'Exactly one candidate path and new PDF path required')
    body = body.replace('set inputPath to item 1 of argv', 'set inputPath to item 1 of argv\n    set verificationPdfPath to item 2 of argv')
    verification = '''                my requireTrue(savedState, "candidate_saved_before")
                set savedBeforeSnapshot to savedState
                repeat with printOption in {printFields, printLinks, printCodes}
                    if class of printOption is not boolean or (contents of printOption) is not false then error "unsafe_print_option" number 7159
                end repeat
                set beforeSnapshot to my captureSnapshot(ownedDocument, my fieldInventory(ownedDocument))
                set savedBeforePdf to get saved of ownedDocument
                my requireTrue(savedBeforePdf, "candidate_saved_before_pdf")
                log {"readonly_snapshot_before_pdf_complete"}
                -- PDF export only. No DOCX save, field update or repaginate.
                save as ownedDocument file name verificationPdfPath file format format PDF add to recent files false
                log {"word_pdf_export_returned"}
                set ownedDocument to my exactDocument(inputPath, false)
                my requireReadOnly((get read only of ownedDocument))
                set savedAfterPdf to get saved of ownedDocument
                if class of savedAfterPdf is not boolean then error "candidate_saved_after_pdf_not_boolean" number 7152
                set afterSnapshot to my captureSnapshot(ownedDocument, my fieldInventory(ownedDocument))
                my requireTrue(my sameTuple(beforeSnapshot, afterSnapshot), "readonly_snapshot_unchanged")
                my requireReadOnly((get read only of ownedDocument))
                set savedAfterSnapshot to get saved of ownedDocument
                if class of savedAfterSnapshot is not boolean then error "candidate_saved_after_snapshot_not_boolean" number 7152
                if (get update fields at print of settings) is not false or (get update links at print of settings) is not false or (get print field codes of settings) is not false then error "print_option_changed" number 7159
                log {"readonly_snapshot_after_pdf_complete", true}
'''
    body = body.replace('                -- End raw getter diagnostics.\n', verification)
    old_return = next(line for line in body.splitlines() if 'return {"readonly_probe_complete"' in line)
    body = body.replace(old_return, '            log {"readonly_pdf_complete", closeOutcome, (get automation security), (get update links at open of settings), count of documents, (get background printing status)}\n            return my jsonValue({beforeSnapshot, afterSnapshot, savedBeforeSnapshot, savedBeforePdf, savedAfterPdf, savedAfterSnapshot, closeOutcome, "", restoreErrors})')
    return helpers + readers + "\non run argv\n" + body


def compare_snapshots(raw: list, calculation: dict) -> list[dict]:
    observed = field.sanitized_snapshots(raw)
    expected = {k: v for k, v in calculation["convergence"][-1].items() if k != "round"}
    if len(observed) != 2 or any({k: v for k, v in row.items() if k != "round"} != expected for row in observed):
        raise ValueError("Read-only candidate differs from accepted calculation evidence")
    return observed


def pdf_binding(pdf: Path, expected_pages: int) -> dict:
    pages = len(PdfReader(pdf).pages)
    if pages != expected_pages:
        raise ValueError("Word PDF page count differs from verified Word snapshot")
    return {"path": str(pdf), "sha256": field.sha256(pdf), "size_bytes": pdf.stat().st_size,
            "page_count": pages, "source": "microsoft_word_save_as_pdf"}


def run(output_dir: Path) -> int:
    source = EVIDENCE / "safe-open-004/v051-synthetic.docx"
    field.validate_input(source)
    expected_calculation = EVIDENCE / "field-010-live-05/result.json"
    if (CALCULATION_RESULT != expected_calculation
            or CALCULATION_RESULT.is_symlink()
            or not CALCULATION_RESULT.is_file()
            or CALCULATION_RESULT.resolve() != expected_calculation
            or CALCULATION_RESULT.stat().st_size != CALCULATION_RESULT_SIZE
            or field.sha256(CALCULATION_RESULT) != CALCULATION_RESULT_SHA256):
        raise ValueError("Only the complete pinned live-05 calculation evidence may be used")
    calculation = json.loads(CALCULATION_RESULT.read_text())
    if (CANDIDATE.is_symlink() or field.sha256(CANDIDATE) != CANDIDATE_SHA256
            or calculation.get("status") != "candidate_ready_for_readonly_verification"
            or calculation.get("content_integrity", {}).get("status") != "pass"
            or calculation.get("selective_writeback", {}).get("status") != "selective_verified"
            or calculation.get("candidate", {}).get("sha256") != CANDIDATE_SHA256
            or calculation.get("candidate", {}).get("path") != str(CANDIDATE)):
        raise ValueError("Only the accepted live-05 synthetic candidate may be verified")
    output_dir = output_dir.resolve()
    if output_dir.parent != EVIDENCE:
        raise ValueError("Verification output must be a new synthetic evidence directory")
    output_dir.mkdir(exist_ok=False)
    pdf = output_dir / "word-verification.pdf"
    script = build_applescript()
    script_path = output_dir / "readonly-pdf.applescript"
    script_path.write_text(script)
    report = {"status": "blocked", "final_ready": False,
              "source_before": {"path": str(source), "sha256": field.sha256(source), "size_bytes": source.stat().st_size},
              "candidate_before": {"path": str(CANDIDATE), "sha256": field.sha256(CANDIDATE), "size_bytes": CANDIDATE.stat().st_size},
              "calculation_before": {"path": str(CALCULATION_RESULT), "sha256": field.sha256(CALCULATION_RESULT), "size_bytes": CALCULATION_RESULT.stat().st_size},
              "manual_control": {
                  "source": "user_reported_word_gui_export_of_same_fixed_candidate_copy",
                  "observation": "Word prompted to save; user chose not to save",
                  "candidate_sha256_after_decline": CANDIDATE_SHA256,
                  "scope": "saved=false alone does not prove disk DOCX changed and does not prove PDF or in-memory content correct"},
              "visual_qa": "not_run", "docx_save_requested": False,
              "field_refresh_requested": False, "repaginate_requested": False}
    try:
        try:
            result = subprocess.run(["/usr/bin/osascript", str(script_path), str(CANDIDATE), str(pdf)],
                                    text=True, capture_output=True, timeout=360, check=False)
        except subprocess.TimeoutExpired as exc:
            stderr = exc.stderr or ""
            if isinstance(stderr, bytes):
                stderr = stderr.decode("utf-8", errors="replace")
            (output_dir / "word-live.stderr.log").write_text(stderr)
            report["timeout_cleanup"] = field.timeout_cleanup(script, CANDIDATE, stderr)
            raise RuntimeError("read_only_pdf_timeout; see exact cleanup outcome") from exc
        (output_dir / "word-live.stderr.log").write_text(result.stderr)
        report["exit_code"] = result.returncode
        if result.returncode:
            raise RuntimeError(result.stderr.strip().splitlines()[-1] if result.stderr.strip() else "Word verification failed")
        evidence = json.loads(result.stdout)
        if not isinstance(evidence, list) or len(evidence) != 9:
            raise ValueError("Word returned incomplete read-only/PDF evidence")
        before, after = evidence[:2]
        saved_values = evidence[2:6]
        if any(type(value) is not bool for value in saved_values):
            raise ValueError("Word returned non-Boolean saved evidence")
        if saved_values[0] is not True or saved_values[1] is not True:
            raise ValueError("Word did not confirm saved=true before PDF export")
        close_outcome, close_errors, restore_errors = evidence[6:9]
        if (close_outcome != "exact_document_closed_without_save"
                or close_errors != "" or restore_errors != ""):
            raise ValueError("Exact no-save close or control restoration was not confirmed")
        report["saved_observations"] = {"before_snapshot": saved_values[0],
                                          "before_pdf": saved_values[1],
                                          "after_pdf": saved_values[2],
                                          "after_snapshot": saved_values[3]}
        report["cleanup"] = {"close_outcome": close_outcome,
                               "close_errors": close_errors,
                               "restore_errors": restore_errors}
        report["read_only_snapshots"] = compare_snapshots([before, after], calculation)
        report["pdf"] = pdf_binding(pdf, calculation["convergence"][-1]["document_pages"])
        report["candidate_content_integrity"] = field.validate_refreshed(source, CANDIDATE)
        report["status"] = "synthetic_readonly_pdf_verified_pending_visual"
    except Exception as exc:
        report["failure"] = {"type": type(exc).__name__, "message": str(exc)}
    finally:
        report["candidate_after"] = {"path": str(CANDIDATE), "sha256": field.sha256(CANDIDATE), "size_bytes": CANDIDATE.stat().st_size}
        report["source_after"] = {"path": str(source), "sha256": field.sha256(source), "size_bytes": source.stat().st_size}
        report["calculation_after"] = {"path": str(CALCULATION_RESULT), "sha256": field.sha256(CALCULATION_RESULT), "size_bytes": CALCULATION_RESULT.stat().st_size}
        report["candidate_bytes_unchanged"] = report["candidate_before"] == report["candidate_after"]
        report["source_unchanged"] = report["source_before"] == report["source_after"]
        report["calculation_unchanged"] = report["calculation_before"] == report["calculation_after"]
        if (not report["candidate_bytes_unchanged"] or not report["source_unchanged"]
                or not report["calculation_unchanged"]):
            report["status"] = "blocked"
            report["failure"] = {"type": "IntegrityError", "message": "Candidate, original or calculation evidence changed"}
        (output_dir / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "synthetic_readonly_pdf_verified_pending_visual" else 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    raise SystemExit(run(parser.parse_args().output_dir))
