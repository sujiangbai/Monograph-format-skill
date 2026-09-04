"""Direct diagnostic tests: pure AppleScript or mocked edges, never Word."""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfWriter

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "adapters/microsoft-word/macos/run_pdf_state_diagnostic.py"
spec = importlib.util.spec_from_file_location("v051_pdf_state", PATH)
diag = importlib.util.module_from_spec(spec)
spec.loader.exec_module(diag)


class PdfStateDiagnosticTests(unittest.TestCase):
    def snapshot(self):
        codes = list(diag.field.EXPECTED_BODY_CODES[1:5]) + [" PAGE ", " PAGE "]
        values = ["3", "2", "1", "Chapter Alpha", "i", "1"]
        return ["\tChapter Alpha\t1\r\tChapter Beta\t2\r", 1, 2, 1,
                [[1, 1, "lowerRoman"], [2, 1, "decimal"]], 3,
                [[i, code, value] for i, (code, value) in enumerate(zip(codes, values), 2)], [1, 2]]

    def calculation(self):
        return {"convergence": diag.field.sanitized_snapshots([self.snapshot(), self.snapshot()])}

    def states(self):
        return [[phase, str(diag.normal.CANDIDATE), index < 2, True, 1, 0, True, False, False, False, False]
                for index, phase in enumerate(diag.PHASES)]

    def test_normal_gate_is_byte_identical_and_diagnostic_only_exports_pdf(self):
        self.assertEqual("ba0e734177c13b93a84451faeb0723443cdc573a043ca1b1dc75ceeb5a753bfc",
                         diag.field.sha256(PATH.with_name("run_readonly_pdf_probe.py")))
        source = diag.build_applescript()
        ownership = PATH.with_name("readonly_settings_probe.applescript").read_text().split("on run argv\n", 1)[0]
        self.assertTrue(source.startswith(ownership))
        self.assertEqual(1, source.count("save as ownedDocument"))
        self.assertIn("file format format PDF add to recent files false", source)
        for forbidden in (r"(?m)^\s*on calculateFields", r"\(update field ", r"(?m)^\s*repaginate doc",
                          r"(?m)^\s*save doc", r"(?m)^\s*set saved of", r"(?m)^\s*set update fields at print",
                          r"(?m)^\s*set update links at print", r"(?m)^\s*set print field codes"):
            self.assertNotRegex(source, forbidden)
        self.assertIn('open file name inputPath read only true add to recent files false', source)
        self.assertIn('my fieldInventory(ownedDocument)', source)
        self.assertIn('"unapproved_quote"', source)
        self.assertIn('my requireTrue(savedState, "candidate_saved_before")', source)
        self.assertEqual(4, source.count('my readDiagnosticState(inputPath, "'))
        self.assertLess(source.index('"after_before_snapshot", false)'), source.index("save as ownedDocument"))
        self.assertLess(source.index('"after_pdf", true)'), source.index("set afterSnapshot"))
        self.assertLess(source.index('"after_after_snapshot", true)'), source.index("set closeOutcome to my closeExactDocument", source.index("on run argv")))

    def test_unexpected_source_drift_refuses_generation(self):
        with self.assertRaises(ValueError):
            diag.replace_once("same same", "same", "new")

    def test_two_actual_snapshots_compare_independently_without_plaintext(self):
        result = diag.compare_observations(self.snapshot(), self.snapshot(), self.calculation())
        self.assertTrue(result["raw_before_after_equal"])
        self.assertEqual([], result["changed_slots"])
        for name in ("before", "after"):
            self.assertTrue(result[name]["complete"])
            self.assertTrue(result[name]["matches_live05"])
        self.assertNotIn("Chapter Alpha", json.dumps(result))
        self.assertNotIn("round", result["before"]["view"])

    def test_actual_difference_does_not_hide_either_reference_result(self):
        for slot, replacement in ((1, 2), (2, 3), (3, 2), (5, 4),
                                  (0, "Chapter Alpha\t1\rChapter Beta\t2\r")):
            changed = self.snapshot()
            changed[slot] = replacement
            with self.subTest(slot=slot):
                result = diag.compare_observations(self.snapshot(), changed, self.calculation())
                self.assertFalse(result["raw_before_after_equal"])
                self.assertEqual([diag.SLOTS[slot]], result["changed_slots"])
                self.assertTrue(result["before"]["matches_live05"])
                self.assertFalse(result["after"]["matches_live05"])
        changed = self.snapshot()
        changed[6][0][2] = "4"
        result = diag.compare_observations(changed, self.snapshot(), self.calculation())
        self.assertFalse(result["before"]["matches_live05"])
        self.assertTrue(result["after"]["matches_live05"])

    def test_missing_instruction_or_tuple_is_explicitly_incomplete(self):
        bad = self.snapshot()
        bad[6][0][1] = " INCLUDETEXT other "
        for row in (None, [], bad, self.snapshot()[:7]):
            with self.subTest(row_type=type(row).__name__):
                result = diag.compare_observations(self.snapshot(), row, self.calculation())
                self.assertFalse(result["after"]["complete"])
                self.assertFalse(result["after"]["matches_live05"])
                self.assertTrue(result["before"]["matches_live05"])

    def test_state_matrix_only_boolean_dirty_after_pdf_is_diagnostic_allowed(self):
        rows = self.states()
        self.assertEqual([True, True, False, False], [row["saved"] for row in diag.validate_states(rows)])
        for phase, slot, value in ((0, 2, False), (1, 2, False), (2, 2, None), (3, 2, "false"),
                                   (2, 3, False), (2, 4, 2), (2, 5, 1), (2, 6, False),
                                   (2, 7, True), (2, 8, True), (2, 9, True), (2, 10, None), (2, 1, "other.docx")):
            changed = copy.deepcopy(rows)
            changed[phase][slot] = value
            with self.subTest(phase=phase, slot=slot), self.assertRaises(ValueError):
                diag.validate_states(changed)

    @unittest.skipUnless(shutil.which("osascript"), "pure AppleScript required")
    def test_actual_applescript_state_guards_offline(self):
        field_text = PATH.with_name("field_probe.applescript").read_text()
        ownership = PATH.with_name("readonly_settings_probe.applescript").read_text()
        handlers = "\n".join(re.search(rf"(?ms)^on {name}\(.*?^end {name}$", text).group(0)
                             for name, text in (("exactText", field_text), ("requireTrue", field_text),
                                                ("requireReadOnly", ownership), ("requireDiagnosticState", diag.DIAGNOSTIC_HANDLERS)))
        base = '{"phase", "/candidate.docx", true, true, 1, 0, true, false, false, false, false}'
        cases = [(base, "false", True), (base.replace('true, true, 1', 'false, true, 1'), "true", True),
                 (base.replace('true, true, 1', 'false, true, 1'), "false", False)]
        values = ['"phase"', '"/candidate.docx"', "true", "true", "1", "0", "true", "false", "false", "false", "false"]
        for slot, value in ((1, '"/other.docx"'), (2, 'missing value'), (2, '"false"'), (3, 'false'), (4, '2'),
                            (5, '1'), (6, 'false'), (7, 'true'), (8, 'missing value'), (9, 'true'), (10, 'true')):
            changed = values.copy()
            changed[slot] = value
            cases.append(("{" + ", ".join(changed) + "}", "true", False))
        for row, allowed, succeeds in cases:
            with self.subTest(row=row, allowed=allowed):
                done = subprocess.run(["/usr/bin/osascript", "-"], input=handlers + f'\nreturn my requireDiagnosticState({row}, "/candidate.docx", {allowed})', text=True, capture_output=True, timeout=10)
                self.assertEqual(succeeds, done.returncode == 0, done.stderr)

    def test_timeout_never_accepts_and_preserves_fixed_artifacts(self):
        with tempfile.TemporaryDirectory(prefix="v051-pdf-diagnostic-unit-") as directory:
            root = Path(directory).resolve()
            source = root / "safe-open-004/v051-synthetic.docx"
            source.parent.mkdir()
            source.write_bytes(b"original")
            candidate = root / "candidate.docx"
            candidate.write_bytes(b"a" * 39546)
            calculation = root / "calculation.json"
            calculation.write_text(json.dumps(self.calculation()))
            with patch.object(diag.normal, "EVIDENCE", root), patch.object(diag.normal, "CANDIDATE", candidate), \
                 patch.object(diag.normal, "CANDIDATE_SHA256", diag.field.sha256(candidate)), \
                 patch.object(diag.normal, "CALCULATION_RESULT", calculation), \
                 patch.object(diag, "CALCULATION_SHA256", diag.field.sha256(calculation)), patch.object(diag, "OUTPUT", root / "out"), \
                 patch.object(diag.field, "validate_input"), \
                 patch.object(diag.subprocess, "run", side_effect=subprocess.TimeoutExpired([], 360, stderr=b"original_controls, msoAutomationSecurityLow, true\n")), \
                 patch.object(diag.field, "timeout_cleanup", return_value={"status": "unconfirmed"}) as cleanup, patch("builtins.print"):
                self.assertEqual(2, diag.run(root / "out"))
                cleanup.assert_called_once()
            report = json.loads((root / "out/result.json").read_text())
            self.assertFalse(report["final_ready"])
            self.assertFalse(report["accepted"])
            self.assertTrue(report["source_unchanged"])
            self.assertTrue(report["candidate_unchanged"])
            self.assertEqual("diagnostic_blocked", report["status"])

    def test_mocked_complete_diagnostic_orchestration_remains_unaccepted(self):
        with tempfile.TemporaryDirectory(prefix="v051-pdf-diagnostic-success-") as directory:
            root = Path(directory).resolve()
            source = root / "safe-open-004/v051-synthetic.docx"
            source.parent.mkdir()
            source.write_bytes(b"original")
            candidate = root / "candidate.docx"
            candidate.write_bytes(b"a" * 39546)
            calculation = root / "calculation.json"
            calculation.write_text(json.dumps(self.calculation()))
            output = root / "out"

            def word_result(command, **kwargs):
                pdf = Path(command[-1])
                writer = PdfWriter()
                for _ in range(3):
                    writer.add_blank_page(width=612, height=792)
                writer.write(pdf)
                raw = [self.states(), self.snapshot(), copy.deepcopy(self.snapshot()), True]
                return subprocess.CompletedProcess(command, 0, stdout=json.dumps(raw), stderr="")

            with patch.object(diag.normal, "EVIDENCE", root), patch.object(diag.normal, "CANDIDATE", candidate), \
                 patch.object(diag.normal, "CANDIDATE_SHA256", diag.field.sha256(candidate)), \
                 patch.object(diag.normal, "CALCULATION_RESULT", calculation), \
                 patch.object(diag, "CALCULATION_SHA256", diag.field.sha256(calculation)), patch.object(diag, "OUTPUT", output), \
                 patch.object(diag.field, "validate_input"), \
                 patch.object(diag.field, "validate_refreshed", return_value={"status": "pass"}), \
                 patch.object(diag.subprocess, "run", side_effect=word_result) as word, patch("builtins.print"):
                self.assertEqual(0, diag.run(output))
                word.assert_called_once()
            report = json.loads((output / "result.json").read_text())
            self.assertEqual("diagnostic_complete_not_accepted", report["status"])
            self.assertFalse(report["accepted"])
            self.assertFalse(report["final_ready"])
            self.assertEqual("not_run", report["visual_qa"])
            self.assertEqual(list(diag.PHASES), [row["phase"] for row in report["states"]])
            self.assertTrue(report["snapshots"]["raw_before_after_equal"])
            self.assertTrue(report["snapshots"]["before"]["matches_live05"])
            self.assertTrue(report["snapshots"]["after"]["matches_live05"])
            self.assertEqual(3, report["pdf"]["page_count"])
            self.assertTrue(report["pdf"]["matches_live05_page_count"])
            self.assertEqual("microsoft_word_save_as_pdf_diagnostic_only", report["pdf"]["source"])


if __name__ == "__main__":
    unittest.main()
