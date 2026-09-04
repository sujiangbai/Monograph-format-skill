"""Offline 010 tests. Never open Word or regenerate the fixed live sample."""
from __future__ import annotations

import copy
import importlib.util
import json
import re
import subprocess
import shutil
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "adapters/microsoft-word/macos/run_field_probe.py"
spec = importlib.util.spec_from_file_location("v051_field_probe", MODULE_PATH)
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)


class FieldProbeTests(unittest.TestCase):
    def snapshot(self):
        codes = list(probe.EXPECTED_BODY_CODES[1:5]) + [" PAGE ", " PAGE "]
        values = ["3", "2", "1", "Chapter Alpha", "i", "1"]
        return ["Chapter Alpha\t1\rChapter Beta\t2\r", 1, 2, 1,
                [[1, 1, "lowerRoman"], [2, 1, "decimal"]], 3,
                [[i, code, value] for i, (code, value) in enumerate(zip(codes, values), 2)], [1, 2]]

    def test_complete_tuple_converges_without_persisting_toc_text(self):
        snapshot = self.snapshot()
        result = probe.sanitized_snapshots([snapshot, copy.deepcopy(snapshot)])
        self.assertEqual(2, result[-1]["toc_entry_count"])
        self.assertEqual([1, 2], result[-1]["toc_entry_pages"])
        self.assertNotIn("Chapter Alpha", json.dumps(result))
        self.assertEqual(6, len(result[-1]["approved_field_results"]))

    def test_partial_or_drifting_tuples_are_not_convergence(self):
        base = self.snapshot()
        for slot, replacement in ((0, "Chapter Beta\t2\rChapter Alpha\t1\r"),
                                  (1, None), (2, 3), (3, 2), (4, []),
                                  (5, 4), (6, []), (7, [2, 3])):
            changed = copy.deepcopy(base)
            changed[slot] = replacement
            with self.subTest(slot=slot), self.assertRaises(ValueError):
                probe.sanitized_snapshots([base, changed])
        for rows in ([base], [base] * 4):
            with self.assertRaises(ValueError):
                probe.sanitized_snapshots(rows)

    def test_word_toc_single_layout_tab_does_not_relax_title_or_page(self):
        snapshot = self.snapshot()
        snapshot[0] = "\tChapter Alpha\t1\r\tChapter Beta\t2\r"
        self.assertEqual([1, 2], probe.sanitized_snapshots([snapshot, copy.deepcopy(snapshot)])[-1]["toc_entry_pages"])
        for text in ("\t\tChapter Alpha\t1\r\tChapter Beta\t2\r",
                     " Chapter Alpha\t1\r\tChapter Beta\t2\r",
                     "\tChapter Alpha \t1\r\tChapter Beta\t2\r",
                     "\tChapter Alpha\t1\r\tChapter Beta\t1\r",
                     "\tChapter Alpha\t01\r\tChapter Beta\t2\r",
                     "\tchapter Alpha\t1\r\tChapter Beta\t2\r"):
            wrong = copy.deepcopy(snapshot)
            wrong[0] = text
            with self.subTest(text_hash=probe.hashlib.sha256(text.encode()).hexdigest()), self.assertRaises(ValueError):
                probe.sanitized_snapshots([wrong, copy.deepcopy(wrong)])
        # A layout change between rounds is still not convergence.
        with self.assertRaises(ValueError):
            probe.sanitized_snapshots([self.snapshot(), snapshot])

    @unittest.skipUnless(shutil.which("osascript"), "macOS pure AppleScript required")
    def test_page_at_uses_integer_offset_and_exact_returned_range(self):
        source = MODULE_PATH.with_name("field_probe.applescript").read_text()
        handler = re.search(r"(?ms)^on pageAt\(.*?^end pageAt$", source).group(0)
        self.assertIn("pageAt(doc, snapshotCharacterOffset, adjusted)", handler)
        self.assertNotIn("start position", handler)
        # Only Word command/property edges are mocked; actual guards execute.
        handler = handler.replace('tell application "Microsoft Word"', '').replace('end tell', '')
        handler = handler.replace('create range doc start snapshotCharacterOffset end snapshotCharacterOffset', '{snapshotCharacterOffset, snapshotCharacterOffset}')
        handler = handler.replace('get start of content of pointRange', 'item 1 of pointRange')
        handler = handler.replace('get end of content of pointRange', 'item 2 of pointRange')
        handler = re.sub(r'get range information pointRange information type active end (?:adjusted )?page number', '"2"', handler)
        for offset, succeeds in ((123, True), ('"123"', False), ('missing value', False), ('true', False), (-1, False)):
            result = subprocess.run(["/usr/bin/osascript", "-"], input=handler + f'\nreturn pageAt(missing value, {offset}, true)',
                                    text=True, capture_output=True, timeout=10)
            self.assertEqual(succeeds, result.returncode == 0, result.stderr)
            if succeeds:
                self.assertEqual("2", result.stdout.strip())
        wrong = handler.replace('{snapshotCharacterOffset, snapshotCharacterOffset}', '{snapshotCharacterOffset, snapshotCharacterOffset + 1}')
        result = subprocess.run(["/usr/bin/osascript", "-"], input=wrong + '\nreturn pageAt(missing value, 123, false)',
                                text=True, capture_output=True, timeout=10)
        self.assertNotEqual(0, result.returncode)
        self.assertIn("snapshot_range_mismatch", result.stderr)

    def test_generated_source_keeps_009_ownership_and_no_collection_update(self):
        source = probe.build_applescript()
        original = MODULE_PATH.with_name("readonly_settings_probe.applescript").read_text()
        self.assertTrue(source.startswith(original.split("on run argv\n", 1)[0]))
        self.assertIn("open file name inputPath read only false add to recent files false", source)
        self.assertIn("my fieldInventory(doc)", source)
        self.assertIn("repeat with roundIndex from 1 to 3", source)
        self.assertIn("save doc", source)
        self.assertNotIn("update fields", source)
        self.assertNotIn("flatten", source)
        self.assertNotIn("set content", source)
        self.assertLess(source.index("set approvedFields to my fieldInventory(doc)"), source.index("update field (item 1 of approvedFields)"))
        self.assertIn("set selectedFields to items 1 thru 5 of bodyFields", source)
        self.assertIn("unapproved_quote", source)

    def test_style_observer_logs_immediately_around_declared_paragraph_getter(self):
        source = MODULE_PATH.with_name("style_observation.applescript").read_text()
        self.assertIn("repeat with paragraphObject in paragraphSnapshot", source)
        self.assertIn("set observedStyle to get style of paragraphObject", source)
        self.assertNotIn("paragraphRange", source)
        self.assertLess(source.index('"direct_paragraph_style_before"'), source.index("set observedStyle to get style"))
        self.assertLess(source.index('"paragraph_style_raw"'), source.index('"nested_start_before"'))
        self.assertNotIn("update field", source)
        self.assertNotIn("save ", source)

    @unittest.skipUnless(shutil.which("osascript"), "macOS pure AppleScript required")
    def test_heading_identity_and_existing_content_guards_offline(self):
        source = MODULE_PATH.with_name("field_probe.applescript").read_text()
        block = source[source.index("        set headingStyle to"):source.index('        my exactText((get content of text object of bookmark')]
        self.assertIn("set observedStyle to get style of paragraphObject", block)
        self.assertIn("if observedStyle is headingStyle then", block)
        self.assertNotIn("name local", block)
        # Execute the actual guard/selection block with only Word API edges
        # substituted. Script objects model identity, not localized name text.
        replacements = {
            "get Word style style heading1 of doc": "mockTarget",
            "is not Word style": "is not script",
            "get built in of headingStyle": "mockBuiltin",
            "get style type of headingStyle": "mockType",
            "style type paragraph": '"paragraph"',
            "get paragraphs of doc": "mockParagraphs",
            "get style of paragraphObject": "get styleRef of paragraphObject",
            "get content of text object of paragraphObject": "get textValue of paragraphObject",
        }
        for original, replacement in replacements.items():
            self.assertIn(original, block)
            block = block.replace(original, replacement)
        handlers = "\n".join(re.search(rf"(?ms)^on {name}\(.*?^end {name}$", source).group(0)
                             for name in ("exactText", "requireTrue"))
        setup = '''
script targetObject
    property labelText : "same localized name"
end script
script otherObject
    property labelText : "same localized name"
end script
script alpha
    property styleRef : missing value
    property textValue : "Chapter Alpha" & return
end script
script beta
    property styleRef : missing value
    property textValue : "Chapter Beta" & return
end script
set styleRef of alpha to targetObject
set styleRef of beta to targetObject
set mockTarget to targetObject
set mockBuiltin to true
set mockType to "paragraph"
set mockParagraphs to {alpha, beta}
'''
        cases = {
            "valid_object_identity": ("", True),
            "target_text_not_object": ('set mockTarget to "same localized name"', False),
            "not_builtin": ("set mockBuiltin to false", False),
            "builtin_not_boolean": ('set mockBuiltin to "true"', False),
            "wrong_type": ('set mockType to "character"', False),
            "same_name_wrong_object": ("set styleRef of beta to otherObject", False),
            "style_name_not_identity": ('set styleRef of beta to "same localized name"', False),
            "extra_heading": ("set mockParagraphs to {alpha, beta, beta}", False),
            "missing_heading": ("set mockParagraphs to {alpha}", False),
            "wrong_order": ("set mockParagraphs to {beta, alpha}", False),
            "changed_full_text": ('set textValue of beta to "Chapter Beta " & return', False),
        }
        for label, (mutation, succeeds) in cases.items():
            with self.subTest(label=label):
                result = subprocess.run(["/usr/bin/osascript", "-"],
                                        input=handlers + setup + mutation + "\n" + block + '\nreturn "passed"',
                                        text=True, capture_output=True, timeout=10)
                if succeeds:
                    self.assertEqual(0, result.returncode, result.stderr)
                    self.assertEqual("passed", result.stdout.strip())
                else:
                    self.assertNotEqual(0, result.returncode)
                    self.assertRegex(result.stderr, r"heading_(style|count|1|2)")

    def test_fixture_position_checks_image_table_neighbors_and_whitespace(self):
        xml = '<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:p><w:r><w:t>before  image</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="20"/></wp:inline></w:drawing></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p></w:body>'
        baseline = etree.fromstring(xml)
        expected = probe.fixture_position_manifest(baseline, [])
        for mutation in ("image_move", "table_move", "whitespace", "container"):
            changed = etree.fromstring(xml)
            if mutation == "image_move":
                changed.append(changed[1])
            elif mutation == "table_move":
                changed.insert(0, changed[2])
            elif mutation == "whitespace":
                changed[0][0][0].text = "before image"
            else:
                changed[2][0][0].append(changed[1])
            with self.subTest(mutation=mutation):
                self.assertNotEqual(expected, probe.fixture_position_manifest(changed, []))

    @unittest.skipUnless(shutil.which("osascript"), "macOS pure AppleScript required")
    def test_toc_descriptor_partition_preserves_all_top_level_fields(self):
        source = MODULE_PATH.with_name("field_probe.applescript").read_text()
        handlers = "\n".join(re.search(rf"(?ms)^on {name}\(.*?^end {name}$", source).group(0)
                             for name in ("exactText", "sameTuple", "jsonValue", "fixtureTopLevelOrdinals"))
        codes = list(probe.EXPECTED_BODY_CODES)
        types = ["TOC", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "QUOTE"]
        ranges = [[90, 104, 105, 196], [438, 448, 449, 450], [465, 479, 480, 481],
                  [496, 517, 518, 519], [537, 554, 555, 568], [589, 618, 619, 633]]
        top = [[code, kind, *bounds] for code, kind, bounds in zip(codes, types, ranges)]
        children = [[' PAGEREF _Toc239373550 \\h ', 'PAGEREF', 120, 147, 148, 149],
                    [' PAGEREF _Toc239373551 \\h ', 'PAGEREF', 165, 192, 193, 194]]
        body = [top[0], *children, *top[1:]]

        def literal(value):
            if isinstance(value, list):
                return "{" + ",".join(literal(item) for item in value) + "}"
            return json.dumps(value)

        def execute(actual_body, scoped):
            call = '\nreturn jsonValue(fixtureTopLevelOrdinals(' + ','.join(
                literal(value) for value in (actual_body, scoped, codes, types)) + '))'
            return subprocess.run(["/usr/bin/osascript", "-"], input=handlers + call,
                                  text=True, capture_output=True, timeout=10)

        for actual_body, scoped, expected in ((top, [], [1, 2, 3, 4, 5, 6]),
                                              (body, children, [1, 4, 5, 6, 7, 8])):
            result = execute(actual_body, scoped)
            self.assertEqual(0, result.returncode, result.stderr)
            self.assertEqual(expected, json.loads(result.stdout))
            # Only first five top-level ordinals are eligible: not QUOTE or children.
            selected_codes = [actual_body[i - 1][0] for i in expected[:5]]
            self.assertEqual(codes[:5], selected_codes)

        for mutation in ("missing_child", "extra_child", "scoped_mismatch", "duplicate_child",
                         "scope_order", "child_wrong_type", "child_outside", "cross_boundary",
                         "touch_outer_marker", "overlapping_children", "body_duplicate",
                         "missing_non_toc", "extra_non_toc", "top_code", "top_type", "top_order",
                         "bad_range_type", "bad_range_order", "non_toc_overlaps", "outer_changed"):
            # The two API collections must not share mutable test records.
            actual_body, scoped = copy.deepcopy(body), copy.deepcopy(children)
            if mutation == "missing_child":
                scoped.pop()
            elif mutation == "extra_child":
                scoped.append(copy.deepcopy(scoped[0]))
            elif mutation == "scoped_mismatch":
                scoped[0][0] += " "
            elif mutation == "duplicate_child":
                scoped[1] = copy.deepcopy(scoped[0])
            elif mutation == "scope_order":
                scoped.reverse()
            elif mutation in ("child_wrong_type", "child_outside", "cross_boundary", "touch_outer_marker", "overlapping_children"):
                if mutation == "child_wrong_type":
                    actual_body[1][1] = "REF"
                elif mutation == "child_outside":
                    actual_body[1][2:] = [220, 247, 248, 249]
                elif mutation == "cross_boundary":
                    actual_body[1][2:] = [100, 147, 148, 149]
                elif mutation == "touch_outer_marker":
                    actual_body[1][2] = 105
                else:
                    actual_body[1][5] = 166
                scoped[0] = copy.deepcopy(actual_body[1])
            elif mutation == "body_duplicate":
                actual_body[2] = copy.deepcopy(actual_body[1])
            elif mutation == "missing_non_toc":
                actual_body.pop()
            elif mutation == "extra_non_toc":
                actual_body.append(copy.deepcopy(actual_body[-1]))
            elif mutation == "top_code":
                actual_body[3][0] = " numpages "
            elif mutation == "top_type":
                actual_body[3][1] = "PAGE"
            elif mutation == "top_order":
                actual_body[3], actual_body[4] = actual_body[4], actual_body[3]
            elif mutation == "bad_range_type":
                actual_body[3][2] = "438"
            elif mutation == "bad_range_order":
                actual_body[3][3] = 450
            elif mutation == "non_toc_overlaps":
                actual_body[4][2:] = [440, 454, 455, 456]
            else:
                actual_body[0][0] = ' TOC \\o "1-2" '
            with self.subTest(mutation=mutation):
                result = execute(actual_body, scoped)
                self.assertNotEqual(0, result.returncode)
                self.assertRegex(result.stderr, r"\(715[14]\)")

    def test_timeout_records_cleanup_and_source_integrity_without_word(self):
        with tempfile.TemporaryDirectory(prefix="v051-field-offline-") as directory:
            root = Path(directory)
            source = root / "synthetic.docx"
            source.write_bytes(b"offline sentinel")
            original_hash = probe.sha256(source)
            args = type("Args", (), {"source": source, "output_dir": root / "output"})()
            manifest = {"later_toc_sources": []}
            with patch.object(probe, "validate_input", return_value=manifest), \
                 patch.object(probe, "EXPECTED_SHA256", original_hash), \
                 patch.object(probe, "EXPECTED_SIZE", source.stat().st_size), \
                 patch.object(probe.subprocess, "run", side_effect=subprocess.TimeoutExpired(["osascript"], 600, stderr=b"original_controls, msoAutomationSecurityLow, true\n")), \
                 patch.object(probe, "timeout_cleanup", return_value={"status": "unconfirmed"}) as cleanup, \
                 patch("builtins.print"):
                self.assertEqual(2, probe.run(args))
                cleanup.assert_called_once()
            report = json.loads((args.output_dir / "result.json").read_text())
            self.assertEqual("blocked", report["status"])
            self.assertTrue(report["source_unchanged"])
            self.assertEqual("unconfirmed", report["timeout_cleanup"]["status"])
            self.assertFalse(report["final_ready"])
            self.assertEqual({"status": "unconfirmed", "reason": "original_controls_not_recorded; no guessed restoration"},
                             probe.timeout_cleanup("", source, ""))

    def test_reuse_is_exact_path_unchanged_and_never_overwrites_old_logs(self):
        with tempfile.TemporaryDirectory(prefix="v051-reuse-offline-") as directory:
            root = Path(directory).resolve()
            source = root / "artifacts/probe-evidence/safe-open-004/v051-synthetic.docx"
            calculation = root / "artifacts/probe-evidence/field-010-live-01/v051-synthetic-word-refreshed.docx"
            source.parent.mkdir(parents=True)
            calculation.parent.mkdir(parents=True)
            source.write_bytes(b"offline reuse sentinel")
            calculation.write_bytes(source.read_bytes())
            old_log = calculation.with_name("result.json")
            old_log.write_text("historical result")
            args = type("Args", (), {"source": source, "output_dir": root / "new-run", "reuse_calculation_copy": True})()
            with patch.object(probe, "ROOT", root), \
                 patch.object(probe, "EXPECTED_SHA256", probe.sha256(source)), \
                 patch.object(probe, "EXPECTED_SIZE", source.stat().st_size), \
                 patch.object(probe, "validate_input", return_value={"later_toc_sources": []}), \
                 patch.object(probe.subprocess, "run", return_value=subprocess.CompletedProcess([], 1, "", "bounded synthetic failure")) as execute, \
                 patch("builtins.print"):
                self.assertEqual(calculation, probe.approved_reuse_copy(source))
                with self.assertRaises(ValueError):
                    probe.approved_reuse_copy(calculation)
                self.assertEqual(2, probe.run(args))
                self.assertEqual(str(calculation), execute.call_args.args[0][-1])
                self.assertEqual(source.read_bytes(), (args.output_dir / "calculation-before.docx").read_bytes())
                self.assertEqual("historical result", old_log.read_text())
                with self.assertRaises(FileExistsError):
                    probe.run(args)
                calculation.write_bytes(b"changed")
                with self.assertRaises(ValueError):
                    probe.approved_reuse_copy(source)

    def test_reuse_refuses_symlink_even_to_identical_bytes(self):
        with tempfile.TemporaryDirectory(prefix="v051-reuse-link-") as directory:
            root = Path(directory).resolve()
            source = root / "artifacts/probe-evidence/safe-open-004/v051-synthetic.docx"
            calculation = root / "artifacts/probe-evidence/field-010-live-01/v051-synthetic-word-refreshed.docx"
            source.parent.mkdir(parents=True)
            calculation.parent.mkdir(parents=True)
            source.write_bytes(b"identical")
            calculation.symlink_to(source)
            with patch.object(probe, "ROOT", root), self.assertRaises(ValueError):
                probe.approved_reuse_copy(source)

    @unittest.skipUnless(shutil.which("osascript"), "macOS pure AppleScript required")
    def test_json_encoder_and_exact_instruction_check_are_offline(self):
        source = MODULE_PATH.with_name("field_probe.applescript").read_text()
        handlers = "\n".join(re.search(rf"(?ms)^on {name}\(.*?^end {name}$", source).group(0) for name in ("jsonValue", "exactText", "sameTuple"))
        pure = handlers + '\nreturn jsonValue({"quoted \\\"text\\\"" & tab & "tail" & return, 3, true, {1, 2}})'
        result = subprocess.run(["/usr/bin/osascript", "-"], input=pure, text=True, capture_output=True, timeout=10)
        self.assertEqual(0, result.returncode, result.stderr)
        self.assertEqual(['quoted "text"\ttail\r', 3, True, [1, 2]], json.loads(result.stdout))
        for value in ('"PAGE"', '" page "', "missing value"):
            bad = subprocess.run(["/usr/bin/osascript", "-"], input=handlers + f'\nexactText({value}, " PAGE ", "field")', text=True, capture_output=True, timeout=10)
            self.assertNotEqual(0, bad.returncode)
        comparison = subprocess.run(["/usr/bin/osascript", "-"], input=handlers + '\nreturn {sameTuple({{"Alpha", 1}}, {{"Alpha", 1}}), sameTuple({{"Alpha", 1}}, {{"alpha", 1}}), sameTuple({{"Alpha", 1}}, {{"Alpha ", 1}})}', text=True, capture_output=True, timeout=10)
        self.assertEqual("true, false, false", comparison.stdout.strip(), comparison.stderr)


if __name__ == "__main__":
    unittest.main()
