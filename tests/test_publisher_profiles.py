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
    / "commercial-press-academic-trial.draft.json"
)
SCHEMA_PATH = (
    ROOT
    / "format-monograph"
    / "references"
    / "format-profile.schema.json"
)


class PublisherProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.profile = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
        cls.schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))

    def test_profile_matches_schema(self) -> None:
        errors = sorted(
            Draft202012Validator(self.schema).iter_errors(self.profile),
            key=lambda error: list(error.absolute_path),
        )
        self.assertEqual([], [error.message for error in errors])

    def test_profile_remains_unapproved_until_sample_qa(self) -> None:
        self.assertEqual("draft", self.profile["approval"]["status"])
        blocking = [
            question
            for question in self.profile["open_questions"]
            if question["blocking"] and not question.get("answer", "").strip()
        ]
        self.assertGreaterEqual(len(blocking), 3)
        for rule in self.profile["rules"]:
            self.assertEqual("draft", rule["status"], rule["id"])
            self.assertEqual("manual_review", rule["application"], rule["id"])

    def test_content_checks_are_report_only(self) -> None:
        content_rules = [
            rule
            for rule in self.profile["rules"]
            if rule["properties"].get("validation") == "content_audit"
        ]
        self.assertTrue(content_rules)
        for rule in content_rules:
            self.assertEqual("report_only", rule["properties"].get("policy"), rule["id"])
            self.assertEqual("manual_review", rule["application"], rule["id"])

    def test_public_source_uses_web_locator_only(self) -> None:
        source = self.profile["sources"][0]
        self.assertTrue(source["public"])
        self.assertTrue(source["locator"].startswith("https://"))
        self.assertNotIn(":\\", source["locator"])
        self.assertNotIn(".doc", source["locator"].lower())


if __name__ == "__main__":
    unittest.main()
