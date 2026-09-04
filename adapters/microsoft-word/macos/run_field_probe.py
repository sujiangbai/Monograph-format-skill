#!/usr/bin/env python3
"""Run the V0.5.1 synthetic Word field/convergence probe."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[3]
SCRIPTS = ROOT / "format-monograph/scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import FormatMonographError, protected_payload_manifest  # noqa: E402
from field_writeback import (  # noqa: E402
    DEFAULT_ALLOWED_FIELD_TYPES,
    _approved_result_text_ids,
    _semantic_part_sources,
    _toc_span,
    _validate_backend_part,
    parse_fields,
    selective_field_result_writeback,
)

EXPECTED_SHA256 = "cff6c281c16320b2a42a565d73b355868c6b1393dac5bf40513e9894edf0105a"
EXPECTED_SIZE = 39369
EXPECTED_BODY_CODES = (
    ' TOC \\o "1-1" ',
    " NUMPAGES ",
    " SECTIONPAGES ",
    " PAGEREF probe_alpha ",
    " REF probe_alpha ",
    ' QUOTE "unapproved_constant" ',
)
FIELD_PARTS = ("word/document.xml", "word/footer1.xml", "word/footer2.xml")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def sha256(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def validate_input(source: Path) -> dict:
    if not source.is_file() or source.stat().st_size != EXPECTED_SIZE:
        raise ValueError("010 requires the unchanged safe-open-004 sample")
    digest = sha256(source)
    if digest != EXPECTED_SHA256:
        raise ValueError("010 sample SHA-256 does not match the approved baseline")
    manifest = json.loads(source.with_name("fixture.json").read_text(encoding="utf-8"))
    if manifest.get("sha256") != digest or manifest.get("size_bytes") != EXPECTED_SIZE:
        raise ValueError("010 fixture manifest does not match the approved sample")
    fields = manifest.get("package_preflight", {}).get("fields", [])
    actual = tuple(item["instruction"] for item in fields if item["part"] == "word/document.xml")
    if actual != EXPECTED_BODY_CODES:
        raise ValueError("010 body-field allowlist differs from the approved manifest")
    if [item["instruction"] for item in fields if item["part"].startswith("word/footer")] != [" PAGE ", " PAGE "]:
        raise ValueError("010 footer-field allowlist differs from the approved manifest")
    if manifest.get("current_step_approved_updates") != []:
        raise ValueError("fixture records unexpected prior updates")
    expected_sources = [{"level": 1, "kind": "heading", "text_sha256": hashlib.sha256(text.encode()).hexdigest()}
                        for text in ("Chapter Alpha", "Chapter Beta")]
    if manifest.get("later_toc_sources") != expected_sources:
        raise ValueError("010 TOC source contract changed")
    return manifest


def apple_literal(values: tuple[str, ...]) -> str:
    return "{" + ", ".join(json.dumps(value) for value in values) + "}"


def approved_reuse_copy(source: Path) -> Path:
    """Only the exact 010 path that the user has now granted Word access to."""
    expected_source = ROOT / "artifacts/probe-evidence/safe-open-004/v051-synthetic.docx"
    calculation = ROOT / "artifacts/probe-evidence/field-010-live-01/v051-synthetic-word-refreshed.docx"
    if source != expected_source or calculation.is_symlink() or calculation.resolve() != calculation:
        raise ValueError("Reuse is limited to the exact approved source/calculation paths")
    if not calculation.is_file() or calculation.stat().st_size != EXPECTED_SIZE or sha256(calculation) != EXPECTED_SHA256:
        raise ValueError("Approved calculation copy changed; refusing reuse")
    return calculation


def build_applescript() -> str:
    ownership = (Path(__file__).with_name("readonly_settings_probe.applescript").read_text(encoding="utf-8"))
    helpers, old_run = ownership.split("on run argv\n", 1)
    calculation = Path(__file__).with_name("field_probe.applescript").read_text(encoding="utf-8")
    calculation = calculation.replace("__BODY_CODES__", apple_literal(EXPECTED_BODY_CODES))
    run_body = """on run argv
    if (count of argv) is not 1 then error "Exactly one calculation-copy DOCX path required" number 7110
    set inputPath to item 1 of argv
    set openAttempted to false
    tell application "Microsoft Word"
        with timeout of 120 seconds
            if (count of documents) is not 0 then error "user_documents_present" number 7101
            if (get background printing status) is not 0 then error "printing_active" number 7102
            set originalSecurity to get automation security
            set originalLinks to get update links at open of settings
            if originalSecurity is missing value then error "security_original_unknown" number 7103
            if class of originalLinks is not boolean then error "links_original_unknown" number 7104
            log {"original_controls", originalSecurity, originalLinks}
            try
                set automation security to msoAutomationSecurityForceDisable
                set update links at open of settings to false
                if (get automation security) is not msoAutomationSecurityForceDisable then error "security_readback_failed" number 7105
                if (get update links at open of settings) is not false then error "links_readback_failed" number 7106
                if (count of documents) is not 0 then error "user_activity_conflict" number 7107
                if (get background printing status) is not 0 then error "printing_conflict" number 7108
                log {"safe_open_controls", (get automation security), (get update links at open of settings)}
                set openAttempted to true
                open file name inputPath read only false add to recent files false
                set ownedDocument to my exactDocument(inputPath, false)
                if (count of documents) is not 1 then error "user_activity_conflict" number 7107
                if (get background printing status) is not 0 then error "printing_conflict" number 7108
                set observedReadOnly to get read only of ownedDocument
                if class of observedReadOnly is not boolean or observedReadOnly is not false then error "calculation_copy_not_editable" number 7157
                log {"calculation_copy_owned", observedReadOnly, class of observedReadOnly}
                set snapshots to my calculateFields(ownedDocument)
                set closeOutcome to my closeExactDocument(inputPath)
                if closeOutcome is not "exact_document_closed_without_save" then error "document_disappeared_before_close" number 7115
                if (count of documents) is not 0 then error "documents_remain_or_user_activity" number 7115
            on error failureMessage number failureNumber
                set recoveryDetails to my recoverAfterFailure(inputPath, openAttempted, originalSecurity, originalLinks)
                error (failureMessage & recoveryDetails) number failureNumber
            end try
            set restoreErrors to my restoreControls(originalSecurity, originalLinks)
            if restoreErrors is not "" then error ("restore_failed:" & restoreErrors) number 7109
            log {"calculation_complete", count of snapshots, closeOutcome, (get automation security), (get update links at open of settings), count of documents, (get background printing status)}
            return my jsonValue(snapshots)
        end timeout
    end tell
end run
"""
    if "open file name inputPath read only true add to recent files false" not in old_run:
        raise ValueError("ownership helper is not the validated 009 form")
    return helpers + calculation + "\n" + run_body


def fixture_position_manifest(root, records) -> list:
    """Small fixture-specific order/whitespace assertion, not an object registry."""
    ignored = _approved_result_text_ids(root, records, set(DEFAULT_ALLOWED_FIELD_TYPES))
    toc_blocks = set()
    for record in records:
        if record.field_type == "TOC":
            container, first, last = _toc_span(root, record)
            toc_blocks.update(list(container)[first:last + 1])
    events = []

    def walk(node):
        if node in ignored or node in toc_blocks:
            return
        name = etree.QName(node).localname
        if name == "t":
            if events and events[-1][0] == "text":
                events[-1][1] += node.text or ""
            else:
                events.append(["text", node.text or ""])
            return
        if name in {"tab", "br", "cr", "gridSpan", "vMerge", "gridCol"}:
            events.append([name, sorted(node.attrib.items())])
        if name == "sectPr":
            # Pagination attributes use the unchanged core contract below.
            # Here only the boundary position amongst authored blocks matters.
            events.append(["section_boundary"])
            return
        if name == "drawing":
            # This fixture has one inline image; its position is represented by
            # the surrounding paragraph/text/table events, not a body index.
            kind = etree.QName(node[0]).localname
            extents = [(child.get("cx"), child.get("cy")) for child in node.iter() if etree.QName(child).localname == "extent"]
            events.append(["drawing", kind, extents])
            return
        boundary = name in {"p", "tbl", "tr", "tc"}
        if boundary:
            events.append(["start", name])
        for child in node:
            walk(child)
        if boundary:
            events.append(["end", name])

    walk(root)
    return events


def validate_refreshed(baseline: Path, refreshed: Path) -> dict:
    if protected_payload_manifest(baseline) != protected_payload_manifest(refreshed):
        raise FormatMonographError("Word changed protected media/formula/embedded payload")
    with zipfile.ZipFile(baseline) as left, zipfile.ZipFile(refreshed) as right:
        story_sources, discarded = _semantic_part_sources(left, right)
        for name in FIELD_PARTS:
            a = etree.fromstring(left.read(name))
            ar = parse_fields(a)
            source_names = story_sources.get(name, [name] if name in right.namelist() else [])
            if not source_names:
                raise FormatMonographError("Word removed a required semantic story")
            for source_name in source_names:
                b = etree.fromstring(right.read(source_name))
                br = parse_fields(b)
                _validate_backend_part(a, b, ar, br, set(DEFAULT_ALLOWED_FIELD_TYPES))
                if fixture_position_manifest(a, ar) != fixture_position_manifest(b, br):
                    raise FormatMonographError(f"Word changed authored whitespace, image/table order, container, adjacency or section boundary in {name}")
    return {"status": "pass", "baseline_sha256": sha256(baseline), "refreshed_sha256": sha256(refreshed),
            "story_mapping": story_sources, "core_discarded_serialization": discarded,
            "fixture_order_container_adjacency_and_whitespace": "pass"}


def timeout_cleanup(source_text: str, calculation_path: Path, stderr: str) -> dict:
    """Only the owned calculation copy; never quit Word or guess original controls."""
    match = re.search(r"(?m)^original_controls, (msoAutomationSecurity(?:Low|ForceDisable|ByUI)), (true|false)$", stderr)
    if not match:
        return {"status": "unconfirmed", "reason": "original_controls_not_recorded; no guessed restoration"}
    handlers = source_text.split("on run argv\n", 1)[0]
    cleanup = handlers + f'''
on run argv
    tell application "Microsoft Word"
        with timeout of 30 seconds
            set details to my recoverAfterFailure(item 1 of argv, true, {match[1]}, {match[2]})
            return {{details, automation security, update links at open of settings, count of documents, background printing status}}
        end timeout
    end tell
end run
'''
    try:
        done = subprocess.run(["/usr/bin/osascript", "-", str(calculation_path)], input=cleanup, text=True, capture_output=True, timeout=180, check=False)
        return {"status": "recorded" if done.returncode == 0 else "unconfirmed", "exit_code": done.returncode,
                "stdout": done.stdout, "stderr": done.stderr}
    except subprocess.TimeoutExpired:
        return {"status": "unconfirmed", "reason": "bounded_exact_cleanup_timed_out"}


def sanitized_snapshots(raw: list) -> list[dict]:
    cleaned = []
    for index, row in enumerate(raw, 1):
        if not isinstance(row, list) or len(row) != 8:
            raise ValueError("Word returned an incomplete convergence tuple")
        toc_text = row[0]
        if not isinstance(toc_text, str) or not toc_text:
            raise ValueError("Word returned an empty TOC")
        entries = [line for line in toc_text.replace("\r", "\n").split("\n") if line]
        if len(entries) != 2 or len(row[6]) != 6 or len(row[7]) != 2:
            raise ValueError("Incomplete TOC/approved fields/heading-page tuple")
        for line, title, page in zip(entries, ("Chapter Alpha", "Chapter Beta"), row[7]):
            # Word's observed TOC envelope may contain one initial layout tab.
            # Do not strip arbitrary whitespace or normalize the title/page.
            # Raw complete tuples below must still be exactly equal.
            if line not in (f"{title}\t{page}", f"\t{title}\t{page}"):
                raise ValueError("TOC text/order/pages do not match approved heading sources")
        if any(type(value) is not int or value < 1 for value in row[1:4] + [row[5]] + row[7]):
            raise ValueError("Unknown or invalid page evidence")
        if row[4] != [[1, 1, "lowerRoman"], [2, 1, "decimal"]]:
            raise ValueError("Section page-number scheme changed")
        for field, code, ordinal in zip(row[6], list(EXPECTED_BODY_CODES[1:5]) + [" PAGE ", " PAGE "], range(2, 8)):
            if not isinstance(field, list) or len(field) != 3 or field[:2] != [ordinal, code] or not isinstance(field[2], str) or not field[2]:
                raise ValueError("Missing or mismatched approved field result")
        cleaned.append({
            "round": index,
            "toc_entry_count": len(entries),
            "toc_entry_sha256": [hashlib.sha256(line.encode()).hexdigest() for line in entries],
            "toc_entry_pages": list(row[7]),
            "toc_pages": row[1],
            "body_start_physical_page": row[2],
            "body_start_logical_page": row[3],
            "sections": row[4],
            "document_pages": row[5],
            "approved_field_results": [{"ordinal": item[0], "code": item[1], "result_sha256": hashlib.sha256(item[2].encode()).hexdigest()} for item in row[6]],
            "heading_logical_pages": row[7],
        })
    if len(cleaned) not in {2, 3} or raw[-1] != raw[-2]:
        raise ValueError("Word did not return two consecutive equal complete tuples")
    return cleaned


def run(args: argparse.Namespace) -> int:
    source, output_dir = args.source.resolve(), args.output_dir.resolve()
    manifest = validate_input(source)
    reuse = bool(getattr(args, "reuse_calculation_copy", False))
    reused_path = approved_reuse_copy(source) if reuse else None
    output_dir.mkdir(parents=True, exist_ok=False)
    refreshed = reused_path or output_dir / "v051-synthetic-word-refreshed.docx"
    candidate = output_dir / "v051-synthetic-selective-candidate.docx"
    if reuse:
        # Preserve the precise pre-run bytes without replacing the granted file
        # or overwriting any old log/result. The original 01 result is historical.
        snapshot = output_dir / "calculation-before.docx"
        shutil.copyfile(refreshed, snapshot)
        if sha256(snapshot) != EXPECTED_SHA256 or sha256(refreshed) != EXPECTED_SHA256:
            raise ValueError("Calculation-copy pre-run snapshot mismatch")
    else:
        shutil.copyfile(source, refreshed)
    source_text = build_applescript()
    (output_dir / "field-probe-010.applescript").write_text(source_text, encoding="utf-8")
    command = ["/usr/bin/osascript", str(output_dir / "field-probe-010.applescript"), str(refreshed)]
    report = {
        "status": "blocked", "final_ready": False,
        "source": {"path": str(source), "sha256": sha256(source), "size_bytes": source.stat().st_size},
        "calculation_copy": {"path": str(refreshed), "sha256": sha256(refreshed), "size_bytes": refreshed.stat().st_size},
        "reused_exact_user_authorized_path": reuse,
        "calculation_before_snapshot": str(snapshot) if reuse else None,
        "command": command,
        "stderr_log": "word-live.stderr.log", "approved_field_types": sorted(DEFAULT_ALLOWED_FIELD_TYPES),
        "toc_contract": manifest["later_toc_sources"],
    }
    try:
        try:
            result = subprocess.run(command, text=True, capture_output=True, timeout=600, check=False)
        except subprocess.TimeoutExpired as exc:
            stderr = exc.stderr or ""
            if isinstance(stderr, bytes):
                stderr = stderr.decode("utf-8", errors="replace")
            (output_dir / "word-live.stderr.log").write_text(stderr, encoding="utf-8")
            report["timeout_cleanup"] = timeout_cleanup(source_text, refreshed, stderr)
            raise RuntimeError("calculation_process_timeout; see exact cleanup outcome") from exc
        report["exit_code"] = result.returncode
        (output_dir / "word-live.stderr.log").write_text(result.stderr, encoding="utf-8")
        report["calculation_copy"].update(sha256=sha256(refreshed), size_bytes=refreshed.stat().st_size)
        if result.returncode:
            raise RuntimeError(result.stderr.strip().splitlines()[-1] if result.stderr.strip() else "Word probe failed without diagnostic")
        snapshots = json.loads(result.stdout)
        report["convergence"] = sanitized_snapshots(snapshots)
        report["content_integrity"] = validate_refreshed(source, refreshed)
        writeback = selective_field_result_writeback(
            source, refreshed, candidate,
            allowed_field_types=DEFAULT_ALLOWED_FIELD_TYPES,
            toc_contract=manifest["later_toc_sources"],
        )
        report["selective_writeback"] = writeback
        report["candidate"] = {"path": str(candidate), "sha256": sha256(candidate), "size_bytes": candidate.stat().st_size}
        report["status"] = "candidate_ready_for_readonly_verification"
    except Exception as exc:
        report["failure"] = {"type": type(exc).__name__, "message": str(exc)}
    finally:
        report["source_after"] = {"sha256": sha256(source), "size_bytes": source.stat().st_size}
        report["source_unchanged"] = report["source_after"] == {"sha256": EXPECTED_SHA256, "size_bytes": EXPECTED_SIZE}
        if not report["source_unchanged"]:
            report["status"] = "blocked"
            report["failure"] = {"type": "SourceIntegrityError", "message": "Pinned source changed; no candidate acceptance"}
        (output_dir / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "candidate_ready_for_readonly_verification" else 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--reuse-calculation-copy", action="store_true",
                        help="Reuse only the unchanged, explicitly authorized field-010-live-01 copy; new output directory still required")
    raise SystemExit(run(parser.parse_args()))
