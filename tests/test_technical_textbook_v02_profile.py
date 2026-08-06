from __future__ import annotations

import json
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
PROFILE_PATH = (
    ROOT / "format-monograph" / "examples" / "profiles"
    / "technical-textbook-layout.v0.2.draft.json"
)
SCHEMA_PATH = (
    ROOT / "format-monograph" / "references" / "format-profile.schema.json"
)


class TechnicalTextbookV02ProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.profile = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        cls.schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))

    def test_profile_matches_schema_11(self) -> None:
        errors = sorted(
            Draft202012Validator(self.schema).iter_errors(self.profile),
            key=lambda error: list(error.absolute_path),
        )
        self.assertEqual([], [error.message for error in errors])
        self.assertEqual("1.1", self.profile["schema_version"])

    def test_runtime_authority_and_equation_policy_are_explicit(self) -> None:
        self.assertEqual("user_requirement", self.profile["source_precedence"][0])
        policy = self.profile["runtime_policy"]
        self.assertTrue(policy["caller_requirements_highest"])
        self.assertTrue(policy["editable_equations_required"])
        self.assertEqual("block", policy["formula_image_policy"])
        self.assertEqual("qa", policy["legacy_equation_policy"])

    def test_profile_is_candidate_not_publication_standard(self) -> None:
        self.assertEqual("draft", self.profile["approval"]["status"])
        self.assertGreaterEqual(len(self.profile["rules"]), 18)
        self.assertTrue(any(q["blocking"] for q in self.profile["open_questions"]))
        for rule in self.profile["rules"]:
            self.assertEqual("draft", rule["status"], rule["id"])

    def test_old_sample_profile_is_not_overwritten(self) -> None:
        old = PROFILE_PATH.with_name("technical-textbook-layout.draft.json")
        self.assertTrue(old.exists())
        parsed = json.loads(old.read_text(encoding="utf-8"))
        self.assertEqual("1.0", parsed["schema_version"])
        self.assertGreaterEqual(len(parsed["rules"]), 20)

    def test_private_sources_have_no_files_or_local_paths(self) -> None:
        for source in self.profile["sources"]:
            locator = source.get("locator", "").lower()
            self.assertNotIn(".pdf", locator)
            self.assertNotIn(":\\", locator)
            self.assertNotIn("d:/", locator)
            if source["type"] == "sample_book":
                self.assertFalse(source["public"])

    def test_automatic_rules_use_supported_v11_properties(self) -> None:
        sys_path = str(ROOT / "format-monograph" / "scripts")
        import sys
        if sys_path not in sys.path:
            sys.path.insert(0, sys_path)
        from validate_profile import validate

        errors, _ = validate(PROFILE_PATH)
        self.assertEqual([], errors)


if __name__ == "__main__":
    unittest.main()
