#!/usr/bin/env python3
"""One opt-in PDF-state observation, never a replacement acceptance gate."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run_readonly_pdf_probe as normal

field = normal.field
OUTPUT = normal.EVIDENCE / "field-010-pdf-state-diagnostic-01"
CALCULATION_SHA256 = "70eb8c837e082ab260b520fef7736b84c001e5508e2973cbe09601d1940761b0"
PHASES = ("before_snapshot", "after_before_snapshot", "after_pdf", "after_after_snapshot")
STATE_KEYS = ("phase", "path", "saved", "read_only", "document_count", "printing_count",
              "security_force_disable", "links_at_open", "print_fields", "print_links", "print_codes")
SLOTS = ("toc_raw_text", "toc_pages", "body_physical", "body_logical", "sections",
         "document_pages", "approved_field_results", "heading_pages")

DIAGNOSTIC_HANDLERS = '''
on requireDiagnosticState(stateRow, inputPath, allowDirty)
    if class of stateRow is not list or (count of stateRow) is not 11 then error "diagnostic_state_incomplete" number 7160
    my exactText(item 2 of stateRow, inputPath, "diagnostic_path")
    set observedSaved to item 3 of stateRow
    if class of observedSaved is not boolean then error "diagnostic_saved_unknown" number 7160
    if allowDirty is not true and observedSaved is not true then error "diagnostic_dirty_before_pdf" number 7160
    my requireReadOnly(item 4 of stateRow)
    if class of item 5 of stateRow is not integer or item 5 of stateRow is not 1 then error "diagnostic_document_conflict" number 7160
    if class of item 6 of stateRow is not integer or item 6 of stateRow is not 0 then error "diagnostic_printing_conflict" number 7160
    my requireTrue(item 7 of stateRow, "diagnostic_security")
    repeat with stateIndex from 8 to 11
        set observedOption to item stateIndex of stateRow
        if class of observedOption is not boolean or observedOption is not false then error "diagnostic_unsafe_option" number 7160
    end repeat
    return stateRow
end requireDiagnosticState

on readDiagnosticState(inputPath, phaseLabel, allowDirty)
    set doc to my exactDocument(inputPath, false)
    tell application "Microsoft Word"
        set currentPath to get posix full name of doc
        log {"pdf_state_path", phaseLabel, currentPath, class of currentPath}
        set currentSaved to get saved of doc
        log {"pdf_state_saved", phaseLabel, currentSaved, class of currentSaved}
        set currentReadOnly to get read only of doc
        log {"pdf_state_read_only", phaseLabel, currentReadOnly, class of currentReadOnly}
        set currentCount to count of documents
        set currentPrinting to get background printing status
        set currentSecurity to get automation security
        set currentLinks to get update links at open of settings
        set currentPrintFields to get update fields at print of settings
        set currentPrintLinks to get update links at print of settings
        set currentPrintCodes to get print field codes of settings
        log {"pdf_state_controls", phaseLabel, currentCount, currentPrinting, currentSecurity, currentLinks, currentPrintFields, currentPrintLinks, currentPrintCodes}
        set securitySafe to currentSecurity is msoAutomationSecurityForceDisable
    end tell
    return my requireDiagnosticState({phaseLabel, currentPath, currentSaved, currentReadOnly, currentCount, currentPrinting, securitySafe, currentLinks, currentPrintFields, currentPrintLinks, currentPrintCodes}, inputPath, allowDirty)
end readDiagnosticState
'''


def replace_once(source: str, old: str, new: str) -> str:
    if source.count(old) != 1:
        raise ValueError("Validated read-only source changed; diagnostic requires review")
    return source.replace(old, new, 1)


def build_applescript() -> str:
    """Copy the normal script in memory; never change its source or saved gate."""
    source = normal.build_applescript()
    source = replace_once(source, "on run argv\n", DIAGNOSTIC_HANDLERS + "\non run argv\n")
    source = replace_once(source, "                set beforeSnapshot to my captureSnapshot", '                set stateRows to {my readDiagnosticState(inputPath, "before_snapshot", false)}\n                set beforeSnapshot to my captureSnapshot')
    source = replace_once(source, '                my requireTrue(savedBeforePdf, "candidate_saved_before_pdf")', '                set end of stateRows to my readDiagnosticState(inputPath, "after_before_snapshot", false)')
    source = replace_once(source, '                if class of savedAfterPdf is not boolean then error "candidate_saved_after_pdf_not_boolean" number 7152', '                set end of stateRows to my readDiagnosticState(inputPath, "after_pdf", true)')
    source = replace_once(source, '                my requireTrue(my sameTuple(beforeSnapshot, afterSnapshot), "readonly_snapshot_unchanged")', '                set observedTupleEqual to my sameTuple(beforeSnapshot, afterSnapshot)')
    source = replace_once(source, '                if class of savedAfterSnapshot is not boolean then error "candidate_saved_after_snapshot_not_boolean" number 7152', '                set end of stateRows to my readDiagnosticState(inputPath, "after_after_snapshot", true)')
    source = replace_once(source, 'log {"readonly_snapshot_after_pdf_complete", true}', 'log {"diagnostic_snapshot_after_pdf_complete", observedTupleEqual}')
    source = replace_once(source, 'log {"readonly_pdf_complete",', 'log {"diagnostic_pdf_complete_not_accepted",')
    source = replace_once(source, 'return my jsonValue({beforeSnapshot, afterSnapshot, savedBeforeSnapshot, savedBeforePdf, savedAfterPdf, savedAfterSnapshot, closeOutcome, "", restoreErrors})', "return my jsonValue({stateRows, beforeSnapshot, afterSnapshot, observedTupleEqual})")
    return source


def digest_value(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()


def snapshot_view(row: list) -> dict:
    """Redact one actual observation, without fabricating a convergence pair."""
    if not isinstance(row, list) or len(row) != 8:
        raise ValueError("incomplete eight-slot snapshot")
    if not isinstance(row[0], str) or not row[0]:
        raise ValueError("missing TOC text")
    entries = [line for line in row[0].replace("\r", "\n").split("\n") if line]
    if len(entries) != 2 or not isinstance(row[6], list) or len(row[6]) != 6 or not isinstance(row[7], list) or len(row[7]) != 2:
        raise ValueError("incomplete TOC/scalar/heading entries")
    if any(type(value) is not int or value < 1 for value in row[1:4] + [row[5]] + row[7]):
        raise ValueError("unknown page evidence")
    if row[4] != [[1, 1, "lowerRoman"], [2, 1, "decimal"]]:
        raise ValueError("section scheme changed")
    for line, title, page in zip(entries, ("Chapter Alpha", "Chapter Beta"), row[7]):
        if line not in (f"{title}\t{page}", f"\t{title}\t{page}"):
            raise ValueError("TOC title/order/page differs from approved sources")
    codes = list(field.EXPECTED_BODY_CODES[1:5]) + [" PAGE ", " PAGE "]
    for value, ordinal, code in zip(row[6], range(2, 8), codes):
        if not isinstance(value, list) or len(value) != 3 or type(value[0]) is not int or value[:2] != [ordinal, code] or not isinstance(value[2], str) or not value[2]:
            raise ValueError("missing or changed scalar instruction/identity/result")
    return {"toc_entry_count": len(entries),
            "toc_entry_sha256": [hashlib.sha256(line.encode()).hexdigest() for line in entries],
            "toc_entry_pages": row[7], "toc_pages": row[1],
            "body_start_physical_page": row[2], "body_start_logical_page": row[3],
            "sections": row[4], "document_pages": row[5],
            "approved_field_results": [{"ordinal": item[0], "code": item[1], "result_sha256": hashlib.sha256(item[2].encode()).hexdigest()} for item in row[6]],
            "heading_logical_pages": row[7]}


def compare_observations(before, after, calculation: dict) -> dict:
    expected = {k: v for k, v in calculation["convergence"][-1].items() if k != "round"}
    result = {"raw_before_after_equal": before == after,
              "changed_slots": [name for index, name in enumerate(SLOTS) if not isinstance(before, list) or not isinstance(after, list) or len(before) != 8 or len(after) != 8 or before[index] != after[index]],
              "reference_scope": "live05 persisted complete redacted evidence; original raw calculation text was not persisted"}
    for name, row in (("before", before), ("after", after)):
        observed = {"raw_snapshot_sha256": digest_value(row)}
        try:
            view = snapshot_view(row)
            observed.update({"complete": True, "view": view, "matches_live05": view == expected,
                             "reference_differences": [key for key in sorted(set(view) | set(expected)) if view.get(key) != expected.get(key)]})
        except (ValueError, TypeError, IndexError) as exc:
            observed.update({"complete": False, "matches_live05": False, "failure": str(exc)})
        result[name] = observed
    return result


def binding(path: Path) -> dict:
    return {"path": str(path), "sha256": field.sha256(path), "size_bytes": path.stat().st_size}


def validate_states(rows) -> list[dict]:
    if not isinstance(rows, list) or len(rows) != 4:
        raise ValueError("Missing diagnostic state points")
    for index, row in enumerate(rows):
        if not isinstance(row, list) or len(row) != len(STATE_KEYS) or row[0] != PHASES[index] or row[1] != str(normal.CANDIDATE):
            raise ValueError("Diagnostic state identity mismatch")
        if (any(type(row[k]) is not bool for k in (2, 3, 6, 7, 8, 9, 10))
                or row[3] is not True or row[6] is not True or any(row[7:])
                or type(row[4]) is not int or row[4] != 1 or type(row[5]) is not int or row[5] != 0
                or (index < 2 and row[2] is not True)):
            raise ValueError("Unsafe diagnostic state")
    return [dict(zip(STATE_KEYS, row)) for row in rows]


def run(output_dir: Path) -> int:
    source = normal.EVIDENCE / "safe-open-004/v051-synthetic.docx"
    field.validate_input(source)
    if (normal.CANDIDATE.is_symlink() or normal.CANDIDATE.resolve() != normal.CANDIDATE
            or field.sha256(normal.CANDIDATE) != normal.CANDIDATE_SHA256
            or normal.CANDIDATE.stat().st_size != 39546
            or field.sha256(normal.CALCULATION_RESULT) != CALCULATION_SHA256):
        raise ValueError("Only the fixed live05 candidate and accepted calculation may be observed")
    calculation = json.loads(normal.CALCULATION_RESULT.read_text())
    if output_dir.resolve() != OUTPUT or output_dir.is_symlink():
        raise ValueError("Only the new authorized diagnostic-01 output is allowed")
    output_dir.mkdir(exist_ok=False)
    pdf = output_dir / "word-verification.pdf"
    script = build_applescript()
    script_path = output_dir / "pdf-state-diagnostic.applescript"
    script_path.write_text(script)
    report = {"status": "diagnostic_blocked", "accepted": False, "final_ready": False,
              "source_before": binding(source), "candidate_before": binding(normal.CANDIDATE),
              "calculation_result": binding(normal.CALCULATION_RESULT),
              "docx_save_requested": False, "field_refresh_requested": False, "repaginate_requested": False,
              "normal_acceptance_entry_modified": False, "visual_qa": "not_run",
              "limitations": "Finite field/TOC/page observations do not prove all Word memory unchanged or explain the saved flag."}
    try:
        try:
            done = subprocess.run(["/usr/bin/osascript", str(script_path), str(normal.CANDIDATE), str(pdf)], text=True, capture_output=True, timeout=360, check=False)
        except subprocess.TimeoutExpired as exc:
            stderr = exc.stderr or ""
            if isinstance(stderr, bytes):
                stderr = stderr.decode("utf-8", errors="replace")
            (output_dir / "word-live.stderr.log").write_text(stderr)
            report["timeout_cleanup"] = field.timeout_cleanup(script, normal.CANDIDATE, stderr)
            raise RuntimeError("diagnostic timeout; structured snapshots missing; inspect raw state logs") from exc
        (output_dir / "word-live.stderr.log").write_text(done.stderr)
        report["exit_code"] = done.returncode
        if done.returncode:
            report["snapshot_gap"] = "No complete structured return; inspect phase-tagged stderr, do not infer missing values"
            raise RuntimeError(done.stderr.strip().splitlines()[-1] if done.stderr.strip() else "Word diagnostic failed")
        rows, before, after, apple_equal = json.loads(done.stdout)
        report["states"] = validate_states(rows)
        report["snapshots"] = compare_observations(before, after, calculation)
        if type(apple_equal) is not bool or apple_equal != report["snapshots"]["raw_before_after_equal"]:
            raise ValueError("AppleScript/Python tuple comparison disagreement")
        report["normal_saved_gate_would_reject"] = any(not row["saved"] for row in report["states"])
        report["candidate_content_integrity"] = field.validate_refreshed(source, normal.CANDIDATE)
        report["status"] = "diagnostic_complete_not_accepted"
    except Exception as exc:
        report["failure"] = {"type": type(exc).__name__, "message": str(exc)}
    finally:
        report["source_after"] = binding(source)
        report["candidate_after"] = binding(normal.CANDIDATE)
        report["source_unchanged"] = report["source_after"] == report["source_before"]
        report["candidate_unchanged"] = report["candidate_after"] == report["candidate_before"]
        if not report["source_unchanged"] or not report["candidate_unchanged"]:
            report["status"] = "diagnostic_blocked"
            report["integrity_failure"] = "Original or candidate bytes changed"
        if pdf.exists():
            report["pdf"] = binding(pdf)
            try:
                report["pdf"]["page_count"] = len(normal.PdfReader(pdf).pages)
                report["pdf"]["matches_live05_page_count"] = report["pdf"]["page_count"] == calculation["convergence"][-1]["document_pages"]
            except Exception as exc:
                report["pdf"]["read_failure"] = str(exc)
            report["pdf"]["source"] = "microsoft_word_save_as_pdf_diagnostic_only"
        (output_dir / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "diagnostic_complete_not_accepted" else 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", required=True, type=Path)
    raise SystemExit(run(parser.parse_args().output_dir))
