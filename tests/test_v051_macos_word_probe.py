"""Synthetic fixture checks only; ordinary unittest never launches Word."""
from __future__ import annotations

import tempfile
import unittest
import zipfile
import json
import re
import shutil
import subprocess
from pathlib import Path

from lxml import etree

from tests.v051_macos_word_fixture import create_fixture, inspect_synthetic_package


PROBE = Path(__file__).resolve().parents[1] / "adapters/microsoft-word/macos/readonly_settings_probe.applescript"


@unittest.skipUnless(shutil.which("osascript"), "offline AppleScript tests need macOS osascript")
class OfflineAppleScriptOwnershipTests(unittest.TestCase):
    """Execute only pure handlers/test doubles, with no application connection."""
    def handler(self, name):
        source = PROBE.read_text(encoding="utf-8")
        return re.search(rf"(?ms)^on {name}\(.*?^end {name}$", source).group(0)

    def run_pure(self, script):
        self.assertNotIn('application "Microsoft Word"', script)
        self.assertNotIn("POSIX file", script)
        return subprocess.run(["/usr/bin/osascript", "-"], input=script, text=True,
                              capture_output=True, timeout=10, check=False)

    def test_no_result_assignment_reproduces_2753_without_any_application(self):
        for initial in ("", "set observed to missing value\n"):
            with self.subTest(initialized=bool(initial)):
                result = self.run_pure("on noResult()\nreturn\nend noResult\n" + initial +
                                       "set observed to noResult()\nreturn observed")
                self.assertNotEqual(0, result.returncode)
                self.assertIn("(-2753)", result.stderr)

    def test_exact_path_selection_matrix(self):
        target = "/synthetic/v051-synthetic.docx"
        cases = (
            ([], False, None, 7112),
            (["/other/v051-synthetic.docx"], False, None, 7112),
            ([target], False, "1", None),
            (["/other/document.docx", target], False, "2", None),
            ([target, target], False, None, 7111),
            ([target.upper()], False, None, 7112),
            ([target + " "], False, None, 7112),
            ([target.replace("v051-", "v051")], False, None, 7112),
            ([target.replace("synthetic/", "synthétic/")], False, None, 7112),
            ([None], False, None, 7113),
            ([], True, "0", None),
        )
        for paths, allow_absent, expected, error in cases:
            values = ", ".join("missing value" if p is None else json.dumps(p, ensure_ascii=False) for p in paths)
            script = self.handler("exactPathIndex") + f'\nreturn exactPathIndex({{{values}}}, "{target}", {str(allow_absent).lower()})'
            with self.subTest(paths=paths, allow_absent=allow_absent):
                result = self.run_pure(script)
                if error:
                    self.assertNotEqual(0, result.returncode)
                    self.assertIn(f"({error})", result.stderr)
                else:
                    self.assertEqual(0, result.returncode, result.stderr)
                    self.assertEqual(expected, result.stdout.strip())

    def test_readonly_must_be_confirmed_boolean_true(self):
        for value in ("true", "false", "missing value", '"true"', "1"):
            script = self.handler("requireReadOnly") + f'\nrequireReadOnly({value})\nreturn "pass"'
            with self.subTest(value=value):
                result = self.run_pure(script)
                if value == "true":
                    self.assertEqual(0, result.returncode, result.stderr)
                else:
                    self.assertNotEqual(0, result.returncode)
                    self.assertIn("(7114)", result.stderr)

    def test_no_result_without_assignment_still_requires_exact_document(self):
        for paths, expected in (("{}", "7112"), ('{"/synthetic/file.docx"}', "1")):
            script = self.handler("exactPathIndex") + f'''
on noResult()
    return
end noResult
noResult()
try
    return exactPathIndex({paths}, "/synthetic/file.docx", false)
on error failure number code
    return code
end try
'''
            with self.subTest(paths=paths):
                result = self.run_pure(script)
                self.assertEqual(0, result.returncode, result.stderr)
                self.assertEqual(expected, result.stdout.strip())

    def test_cleanup_zero_multiple_unique_and_failed_close(self):
        # Substitute only the app boundary and close event; run the actual cleanup
        # handler and exact-path selector, not a separate Python implementation.
        cleanup = self.handler("closeExactDocument").replace('tell application "Microsoft Word"', "tell me").replace(
            "close closeTarget saving no", "my recordClose(closeTarget)")
        for scenario, expected in (("absent", "no_exact_document_to_close, 0"),
                                   ("multiple", "7111, 0"),
                                   ("unique", "exact_document_closed_without_save, 1"),
                                   ("remains", "7115, 1"),
                                   ("close_error", "7120, 1"),
                                   ("path_error", "7113, 0")):
            doubles = f'''
property scenario : "{scenario}"
property closeCalls : 0
on exactDocument(inputPath, allowAbsent)
    if scenario is "path_error" then error "path_unavailable" number 7113
    set paths to {{inputPath}}
    if scenario is "multiple" then set paths to {{inputPath, inputPath}}
    if scenario is "absent" or (scenario is "unique" and closeCalls is 1) then set paths to {{}}
    set picked to my exactPathIndex(paths, inputPath, allowAbsent)
    if picked is 0 then return missing value
    return "synthetic_document_reference"
end exactDocument
on recordClose(documentReference)
    if documentReference is not "synthetic_document_reference" then error "wrong close target"
    set closeCalls to closeCalls + 1
    if scenario is "close_error" then error "close_failed" number 7120
end recordClose
'''
            invocation = '''
try
    set outcome to closeExactDocument("/synthetic/v051-synthetic.docx")
    return {outcome, closeCalls}
on error failure number code
    return {code, closeCalls}
end try
'''
            with self.subTest(scenario=scenario):
                result = self.run_pure(doubles + self.handler("exactPathIndex") + "\n" + cleanup + invocation)
                self.assertEqual(0, result.returncode, result.stderr)
                self.assertEqual(expected, result.stdout.strip())

    def test_after_open_exception_restores_even_if_cleanup_is_ambiguous(self):
        for attempted, close_raises, expected_trace in ((False, False, "restore"),
                                                        (True, False, "close, restore"),
                                                        (True, True, "close, restore")):
            doubles = f'''
property calls : {{}}
on closeExactDocument(inputPath)
    set end of calls to "close"
    if {str(close_raises).lower()} then error "multiple_exact_path_matches" number 7111
    return "exact_document_closed_without_save"
end closeExactDocument
on restoreControls(originalSecurity, originalLinks)
    set end of calls to "restore"
    return ""
end restoreControls
'''
            script = doubles + self.handler("recoverAfterFailure") + f'''
set details to recoverAfterFailure("/synthetic/file.docx", {str(attempted).lower()}, "original", true)
return {{details, calls}}
'''
            with self.subTest(attempted=attempted, ambiguous=close_raises):
                result = self.run_pure(script)
                self.assertEqual(0, result.returncode, result.stderr)
                self.assertTrue(result.stdout.strip().endswith(expected_trace), result.stdout)
                if close_raises:
                    self.assertIn("close_not_verified", result.stdout)
                    self.assertIn("multiple_exact_path_matches", result.stdout)

    def test_main_keeps_open_errors_and_restores_after_exact_cleanup(self):
        source = PROBE.read_text(encoding="utf-8")
        self.assertNotIn("active document", source)
        self.assertNotIn("set openResult to open", source)
        self.assertNotIn('document "v051-synthetic.docx"', source)
        open_calls = re.findall(r"(?m)^\s*open .+$", source)
        self.assertEqual(["open file name inputPath read only true add to recent files false"],
                         [line.strip() for line in open_calls])
        self.assertNotIn("open (POSIX file inputPath)", source)
        self.assertIn("with timeout of 120 seconds", source)
        self.assertIn("set ownedDocument to my exactDocument(inputPath, false)", source)
        failure = source.split("on error failureMessage number failureNumber", 1)[1]
        recovery = self.handler("recoverAfterFailure")
        self.assertLess(recovery.index("closeExactDocument(inputPath)"), recovery.index("restoreControls(originalSecurity, originalLinks)"))
        self.assertIn("recoverAfterFailure(inputPath, openAttempted, originalSecurity, originalLinks)", failure)
        self.assertIn("number failureNumber", failure)
        self.assertIn("close_outcome=", recovery)

    def diagnostic_block(self, readonly_value):
        source = PROBE.read_text(encoding="utf-8")
        block = source.split("-- Begin raw getter diagnostics:", 1)[1]
        block = block.split("\n", 1)[1].split("-- End raw getter diagnostics.", 1)[0]
        for getter, value in (("get read only of ownedDocument", readonly_value),
                              ("get saved of ownedDocument", "true"),
                              ("get update fields at print of settings", "false"),
                              ("get update links at print of settings", "missing value"),
                              ("get print field codes of settings", "false")):
            block = block.replace(getter, value)
        return block

    def test_raw_diagnostics_record_all_values_before_unchanged_readonly_guard(self):
        for value in ("true", "false", "missing value", '"true"', "1"):
            script = self.handler("requireReadOnly") + "\n" + self.diagnostic_block(value) + '\nreturn "pass"'
            with self.subTest(readonly=value):
                result = self.run_pure(script)
                for label in ("read_only", "saved", "print_fields", "print_links", "print_codes"):
                    self.assertIn(f"diagnostic_{label},", result.stderr)
                if value in {"true", "false"}:
                    self.assertIn(f"diagnostic_read_only, {value}, boolean", result.stderr)
                if value == "true":
                    self.assertEqual(0, result.returncode, result.stderr)
                else:
                    self.assertNotEqual(0, result.returncode)
                    self.assertIn("(7114)", result.stderr)

    def test_diagnostic_query_errors_are_not_swallowed(self):
        block = self.diagnostic_block("false").replace("set printFields to false", "set printFields to failedGetter()")
        script = self.handler("requireReadOnly") + '''
on failedGetter()
    error "synthetic_getter_failed" number 7140
end failedGetter
''' + block + '\nreturn "unexpected_success"'
        result = self.run_pure(script)
        self.assertNotEqual(0, result.returncode)
        self.assertIn("(7140)", result.stderr)
        self.assertIn("diagnostic_read_only, false, boolean", result.stderr)
        self.assertNotIn("diagnostic_print_links", result.stderr)


class MacWordSyntheticProbeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="v051-word-unit-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.manifest = create_fixture(self.root / "generated")
        self.docx = self.root / "generated/v051-synthetic.docx"

    def test_required_synthetic_objects_and_explicit_field_scope(self):
        check = self.manifest["package_preflight"]
        self.assertEqual((2, 1, 1), (check["section_count"], check["table_count"], check["image_count"]))
        self.assertEqual((0, 0), (check["macro_payloads"], check["external_connections"]))
        self.assertEqual([], self.manifest["current_step_approved_updates"])
        self.assertEqual(2, len(self.manifest["later_toc_sources"]))
        self.assertEqual(7, sum(item["approved_for_later_refresh"] for item in check["fields"]))
        unapproved = [item for item in check["fields"] if not item["approved_for_later_refresh"]]
        self.assertEqual([' QUOTE "unapproved_constant" '], [item["instruction"] for item in unapproved])
        with zipfile.ZipFile(self.docx) as package:
            root = etree.fromstring(package.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            self.assertEqual(["lowerRoman", "decimal"], root.xpath("//w:pgNumType/@w:fmt", namespaces=ns))
            self.assertEqual(["1", "1"], root.xpath("//w:pgNumType/@w:start", namespaces=ns))
            self.assertEqual(["probe_alpha", "probe_beta"], root.xpath("//w:bookmarkStart/@w:name", namespaces=ns))
            self.assertEqual(["2700", "6660"], root.xpath("//w:tblGrid/w:gridCol/@w:w", namespaces=ns))

    def test_fixture_never_overwrites_existing_output(self):
        original = self.docx.read_bytes()
        with self.assertRaises(FileExistsError):
            create_fixture(self.root / "generated")
        self.assertEqual(original, self.docx.read_bytes())

    def test_fixture_audit_rejects_inserted_external_relationship(self):
        changed = self.root / "external.docx"
        with zipfile.ZipFile(self.docx) as source, zipfile.ZipFile(changed, "w") as output:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename == "word/_rels/document.xml.rels":
                    root = etree.fromstring(data)
                    node = etree.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                    node.set("Id", "negative-synthetic")
                    node.set("TargetMode", "External")
                    node.set("Target", "https://invalid.example/synthetic")
                    data = etree.tostring(root)
                output.writestr(member, data)
        with self.assertRaisesRegex(ValueError, "external relationship"):
            inspect_synthetic_package(changed)

    def test_fixture_audit_rejects_macro_payload(self):
        changed = self.root / "macro.docx"
        changed.write_bytes(self.docx.read_bytes())
        with zipfile.ZipFile(changed, "a") as package:
            package.writestr("word/vbaProject.bin", b"synthetic marker only")
        with self.assertRaisesRegex(ValueError, "macro/active payload"):
            inspect_synthetic_package(changed)


if __name__ == "__main__":
    unittest.main()
