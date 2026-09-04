"""Offline tests only; never start Word."""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfWriter

ROOT = Path(__file__).resolve().parents[1]
path = ROOT / "adapters/microsoft-word/macos/run_readonly_pdf_probe.py"
spec = importlib.util.spec_from_file_location("v051_readonly_pdf", path)
verify = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify)


class ReadonlyPdfTests(unittest.TestCase):
    def snapshot(self):
        codes = list(verify.field.EXPECTED_BODY_CODES[1:5]) + [" PAGE ", " PAGE "]
        values = ["3", "2", "1", "Chapter Alpha", "i", "1"]
        return ["\tChapter Alpha\t1\r\tChapter Beta\t2\r", 1, 2, 1,
                [[1, 1, "lowerRoman"], [2, 1, "decimal"]], 3,
                [[i, code, value] for i, (code, value) in enumerate(zip(codes, values), 2)], [1, 2]]

    def calculation(self, candidate: Path, pages: int = 3) -> dict:
        snapshot = self.snapshot()
        if pages != 3:
            snapshot[5] = pages
            snapshot[6][0][2] = str(pages)
        return {"status": "candidate_ready_for_readonly_verification",
                "content_integrity": {"status": "pass"},
                "selective_writeback": {"status": "selective_verified"},
                "candidate": {"path": str(candidate), "sha256": verify.field.sha256(candidate)},
                "convergence": verify.field.sanitized_snapshots([snapshot, copy.deepcopy(snapshot)])}

    def test_only_pdf_export_is_reachable_after_original_readonly_controls(self):
        source = verify.build_applescript()
        ownership = path.with_name("readonly_settings_probe.applescript").read_text().split("on run argv\n", 1)[0]
        self.assertTrue(source.startswith(ownership))
        self.assertIn('open file name inputPath read only true add to recent files false', source)
        self.assertNotRegex(source, r'(?m)^\s*(?:my requireTrue\(\()?update fields?\b')
        self.assertNotIn('(update field ', source)
        self.assertNotIn('repaginate doc', source)
        self.assertNotIn('save doc', source)
        self.assertNotIn('on calculateFields', source)
        self.assertEqual(1, source.count('save as ownedDocument'))
        self.assertIn('file format format PDF add to recent files false', source)
        self.assertLess(source.index('my requireReadOnly(wasReadOnly)'), source.index('save as ownedDocument'))
        self.assertLess(source.index('"unsafe_print_option"'), source.index('save as ownedDocument'))
        self.assertIn('my sameTuple(beforeSnapshot, afterSnapshot)', source)
        self.assertIn('my exactDocument(inputPath, false)', source)
        self.assertEqual(2, source.count('my requireReadOnly((get read only of ownedDocument))'))
        self.assertIn('class of savedAfterPdf is not boolean', source)
        self.assertIn('class of savedAfterSnapshot is not boolean', source)
        self.assertNotIn('"candidate_saved_after_pdf")', source)
        self.assertNotIn('"candidate_saved_after_snapshot")', source)
        self.assertNotIn('set update fields at print', source)
        self.assertNotRegex(source, r'(?m)^\s*(?:update field|repaginate doc|save doc|set saved of)\b')

    def test_both_readonly_snapshots_must_equal_accepted_calculation(self):
        raw = [self.snapshot(), self.snapshot()]
        calculation = {"convergence": verify.field.sanitized_snapshots(raw)}
        self.assertEqual(2, len(verify.compare_snapshots(raw, calculation)))
        for slot, value in ((1, 2), (2, 3), (3, 2), (5, 4), (7, [1, 1])):
            changed = copy.deepcopy(raw)
            changed[0][slot] = changed[1][slot] = value
            with self.subTest(slot=slot), self.assertRaises(ValueError):
                verify.compare_snapshots(changed, calculation)
        changed = copy.deepcopy(raw)
        changed[1][6][0][2] = "4"
        with self.assertRaises(ValueError):
            verify.compare_snapshots(changed, calculation)

    def test_pdf_page_count_hash_and_size_bind_actual_pdf(self):
        with tempfile.TemporaryDirectory(prefix="v051-pdf-unit-") as directory:
            pdf = Path(directory) / "synthetic.pdf"
            writer = PdfWriter()
            for _ in range(3):
                writer.add_blank_page(width=100, height=100)
            writer.write(pdf)
            result = verify.pdf_binding(pdf, 3)
            self.assertEqual(3, result["page_count"])
            self.assertEqual(verify.field.sha256(pdf), result["sha256"])
            self.assertEqual(pdf.stat().st_size, result["size_bytes"])
            with self.assertRaises(ValueError):
                verify.pdf_binding(pdf, 2)

    def test_timeout_remains_blocked_and_preserves_both_docx_bindings(self):
        with tempfile.TemporaryDirectory(prefix="v051-pdf-timeout-") as directory:
            root = Path(directory).resolve()
            source = root / "safe-open-004/v051-synthetic.docx"
            source.parent.mkdir()
            source.write_bytes(b"original")
            candidate = root / "candidate.docx"
            candidate.write_bytes(b"candidate")
            digest = verify.field.sha256(candidate)
            result_path = root / "field-010-live-05/result.json"
            result_path.parent.mkdir()
            result_path.write_text(json.dumps(self.calculation(candidate)))
            with patch.object(verify, "EVIDENCE", root), patch.object(verify, "CANDIDATE", candidate), \
                 patch.object(verify, "CANDIDATE_SHA256", digest), patch.object(verify, "CALCULATION_RESULT", result_path), \
                 patch.object(verify, "CALCULATION_RESULT_SHA256", verify.field.sha256(result_path)), \
                 patch.object(verify, "CALCULATION_RESULT_SIZE", result_path.stat().st_size), \
                 patch.object(verify.field, "validate_input"), \
                 patch.object(verify.field, "EXPECTED_SHA256", verify.field.sha256(source)), \
                 patch.object(verify.field, "EXPECTED_SIZE", source.stat().st_size), \
                 patch.object(verify.subprocess, "run", side_effect=subprocess.TimeoutExpired([], 360, stderr=b"original_controls, msoAutomationSecurityLow, true\n")), \
                 patch.object(verify.field, "timeout_cleanup", return_value={"status": "unconfirmed"}) as cleanup, \
                 patch("builtins.print"):
                self.assertEqual(2, verify.run(root / "output"))
                cleanup.assert_called_once()
            report = json.loads((root / "output/result.json").read_text())
            self.assertEqual("blocked", report["status"])
            self.assertFalse(report["final_ready"])
            self.assertTrue(report["source_unchanged"])
            self.assertTrue(report["candidate_bytes_unchanged"])
            self.assertEqual("unconfirmed", report["timeout_cleanup"]["status"])

    def test_calculation_bytes_convergence_symlink_and_path_replacement_fail_before_word(self):
        with tempfile.TemporaryDirectory(prefix="v051-calculation-binding-") as directory:
            root = Path(directory).resolve()
            source = root / "safe-open-004/v051-synthetic.docx"
            source.parent.mkdir()
            source.write_bytes(b"original")
            candidate = root / "candidate.docx"
            candidate.write_bytes(b"candidate")
            fixed = root / "field-010-live-05/result.json"
            fixed.parent.mkdir()
            original = json.dumps(self.calculation(candidate), sort_keys=True)
            fixed.write_text(original)
            expected_hash, expected_size = verify.field.sha256(fixed), fixed.stat().st_size

            cases = []
            fixed.write_text(original + " ")
            cases.append(("json_bytes", fixed))
            forged = self.calculation(candidate, pages=4)
            fixed.write_text(json.dumps(forged, sort_keys=True))
            cases.append(("self_consistent_four_pages", fixed))
            target = root / "actual-result.json"
            target.write_text(original)
            fixed.unlink()
            fixed.symlink_to(target)
            cases.append(("symlink", fixed))
            replacement = root / "replacement.json"
            replacement.write_text(original)
            cases.append(("path_replacement", replacement))

            for name, result_path in cases:
                if name == "json_bytes":
                    fixed.unlink(missing_ok=True)
                    fixed.write_text(original + " ")
                elif name == "self_consistent_four_pages":
                    fixed.unlink(missing_ok=True)
                    fixed.write_text(json.dumps(forged, sort_keys=True))
                elif name == "symlink":
                    fixed.unlink(missing_ok=True)
                    fixed.symlink_to(target)
                with self.subTest(case=name), \
                     patch.object(verify, "EVIDENCE", root), patch.object(verify, "CANDIDATE", candidate), \
                     patch.object(verify, "CANDIDATE_SHA256", verify.field.sha256(candidate)), \
                     patch.object(verify, "CALCULATION_RESULT", result_path), \
                     patch.object(verify, "CALCULATION_RESULT_SHA256", expected_hash), \
                     patch.object(verify, "CALCULATION_RESULT_SIZE", expected_size), \
                     patch.object(verify.field, "validate_input"), patch.object(verify.subprocess, "run") as word:
                    with self.assertRaisesRegex(ValueError, "complete pinned live-05"):
                        verify.run(root / f"out-{name}")
                    word.assert_not_called()

    def test_saved_false_or_true_with_complete_direct_evidence_stops_pending_visual(self):
        with tempfile.TemporaryDirectory(prefix="v051-readonly-success-") as directory:
            root = Path(directory).resolve()
            source = root / "safe-open-004/v051-synthetic.docx"
            source.parent.mkdir()
            source.write_bytes(b"original")
            candidate = root / "candidate.docx"
            candidate.write_bytes(b"candidate")
            result_path = root / "field-010-live-05/result.json"
            result_path.parent.mkdir()
            result_path.write_text(json.dumps(self.calculation(candidate)))
            snapshots = [self.snapshot(), copy.deepcopy(self.snapshot())]

            for saved_after in (False, True):
                output = root / f"output-{str(saved_after).lower()}"

                def word_result(command, **kwargs):
                    pdf = Path(command[-1])
                    writer = PdfWriter()
                    for _ in range(3):
                        writer.add_blank_page(width=612, height=792)
                    writer.write(pdf)
                    direct = snapshots + [True, True, saved_after, saved_after,
                                          "exact_document_closed_without_save", "", ""]
                    return subprocess.CompletedProcess(command, 0, stdout=json.dumps(direct), stderr="")

                with self.subTest(saved_after=saved_after), \
                     patch.object(verify, "EVIDENCE", root), patch.object(verify, "CANDIDATE", candidate), \
                     patch.object(verify, "CANDIDATE_SHA256", verify.field.sha256(candidate)), \
                     patch.object(verify, "CALCULATION_RESULT", result_path), \
                     patch.object(verify, "CALCULATION_RESULT_SHA256", verify.field.sha256(result_path)), \
                     patch.object(verify, "CALCULATION_RESULT_SIZE", result_path.stat().st_size), \
                     patch.object(verify.field, "validate_input"), \
                     patch.object(verify.field, "validate_refreshed", return_value={"status": "pass"}), \
                     patch.object(verify.subprocess, "run", side_effect=word_result) as word, patch("builtins.print"):
                    self.assertEqual(0, verify.run(output))
                    word.assert_called_once()
                report = json.loads((output / "result.json").read_text())
                self.assertEqual("synthetic_readonly_pdf_verified_pending_visual", report["status"])
                self.assertFalse(report["final_ready"])
                self.assertEqual("not_run", report["visual_qa"])
                self.assertEqual({"before_snapshot": True, "before_pdf": True,
                                  "after_pdf": saved_after, "after_snapshot": saved_after},
                                 report["saved_observations"])
                self.assertEqual({"close_outcome": "exact_document_closed_without_save",
                                  "close_errors": "", "restore_errors": ""}, report["cleanup"])
                self.assertEqual(2, len(report["read_only_snapshots"]))
                self.assertEqual(3, report["pdf"]["page_count"])
                self.assertTrue(report["candidate_bytes_unchanged"])
                self.assertTrue(report["source_unchanged"])
                self.assertTrue(report["calculation_unchanged"])
                self.assertIn("user_reported_word_gui_export", report["manual_control"]["source"])

    def test_nonboolean_drift_cleanup_readonly_and_entity_changes_stay_blocked(self):
        cases = ("nonboolean_saved", "pre_saved_false", "snapshot_drift", "close_uncertain",
                 "restore_error", "read_only_lost", "candidate_changed", "calculation_changed")
        for case in cases:
            with self.subTest(case=case), tempfile.TemporaryDirectory(prefix=f"v051-{case}-") as directory:
                root = Path(directory).resolve()
                source = root / "safe-open-004/v051-synthetic.docx"
                source.parent.mkdir()
                source.write_bytes(b"original")
                candidate = root / "candidate.docx"
                candidate.write_bytes(b"candidate")
                result_path = root / "field-010-live-05/result.json"
                result_path.parent.mkdir()
                result_path.write_text(json.dumps(self.calculation(candidate)))
                before, after = self.snapshot(), copy.deepcopy(self.snapshot())
                saved = [True, True, False, False]
                close = ["exact_document_closed_without_save", "", ""]
                if case == "nonboolean_saved":
                    saved[2] = "false"
                elif case == "pre_saved_false":
                    saved[1] = False
                elif case == "snapshot_drift":
                    after[5] = 4
                elif case == "close_uncertain":
                    close[0] = "close_not_verified"
                    close[1] = "close failed"
                elif case == "restore_error":
                    close[2] = "security restore failed"

                def word_result(command, **kwargs):
                    if case == "read_only_lost":
                        return subprocess.CompletedProcess(command, 1, stdout="", stderr="not_confirmed_readonly (7114)")
                    pdf = Path(command[-1])
                    writer = PdfWriter()
                    for _ in range(3):
                        writer.add_blank_page(width=612, height=792)
                    writer.write(pdf)
                    if case == "candidate_changed":
                        candidate.write_bytes(candidate.read_bytes() + b"changed")
                    if case == "calculation_changed":
                        result_path.write_text(result_path.read_text() + " ")
                    return subprocess.CompletedProcess(command, 0,
                                                       stdout=json.dumps([before, after] + saved + close), stderr="")

                with patch.object(verify, "EVIDENCE", root), patch.object(verify, "CANDIDATE", candidate), \
                     patch.object(verify, "CANDIDATE_SHA256", verify.field.sha256(candidate)), \
                     patch.object(verify, "CALCULATION_RESULT", result_path), \
                     patch.object(verify, "CALCULATION_RESULT_SHA256", verify.field.sha256(result_path)), \
                     patch.object(verify, "CALCULATION_RESULT_SIZE", result_path.stat().st_size), \
                     patch.object(verify.field, "validate_input"), \
                     patch.object(verify.field, "validate_refreshed", return_value={"status": "pass"}), \
                     patch.object(verify.subprocess, "run", side_effect=word_result), patch("builtins.print"):
                    self.assertEqual(2, verify.run(root / "output"))
                report = json.loads((root / "output/result.json").read_text())
                self.assertEqual("blocked", report["status"])
                self.assertFalse(report["final_ready"])


if __name__ == "__main__":
    unittest.main()
