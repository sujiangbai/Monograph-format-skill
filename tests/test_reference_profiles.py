from __future__ import annotations

import json
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator


ROOT = Path(__file__).resolve().parents[1]
PROFILE_PATH = (
    ROOT
    / "format-monograph"
    / "examples"
    / "profiles"
    / "gbt-2011-editorial-baseline.draft.json"
)
SCHEMA_PATH = (
    ROOT
    / "format-monograph"
    / "references"
    / "format-profile.schema.json"
)


class ReferenceProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.profile = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        cls.schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))

    def test_draft_profile_matches_schema(self) -> None:
        Draft202012Validator.check_schema(self.schema)
        errors = sorted(
            Draft202012Validator(self.schema).iter_errors(self.profile),
            key=lambda error: list(error.absolute_path),
        )
        self.assertEqual([], [error.message for error in errors])

    def test_reference_rules_cannot_modify_authored_content(self) -> None:
        self.assertEqual("draft", self.profile["approval"]["status"])
        self.assertTrue(self.profile["open_questions"])
        self.assertTrue(
            any(question["blocking"] for question in self.profile["open_questions"])
        )
        for rule in self.profile["rules"]:
            self.assertEqual("draft", rule["status"], rule["id"])
            self.assertEqual("manual_review", rule["application"], rule["id"])

    def test_repository_does_not_record_private_pdf_paths(self) -> None:
        for source in self.profile["sources"]:
            self.assertFalse(source["public"])
            locator = source.get("locator", "").lower()
            self.assertNotIn(".pdf", locator)
            self.assertNotIn(":\\", locator)


if __name__ == "__main__":
    unittest.main()
