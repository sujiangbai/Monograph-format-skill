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
    / "technical-textbook-layout.draft.json"
)
SCHEMA_PATH = (
    ROOT
    / "format-monograph"
    / "references"
    / "format-profile.schema.json"
)


class TextbookSampleProfileTests(unittest.TestCase):
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

    def test_sample_stays_unapproved_and_manual(self) -> None:
        self.assertEqual("draft", self.profile["approval"]["status"])
        self.assertGreaterEqual(len(self.profile["rules"]), 20)
        for rule in self.profile["rules"]:
            self.assertEqual("draft", rule["status"], rule["id"])
            self.assertEqual("manual_review", rule["application"], rule["id"])

    def test_private_scan_is_not_committed_or_addressed(self) -> None:
        source = self.profile["sources"][0]
        self.assertEqual("sample_book", source["type"])
        self.assertFalse(source["public"])
        locator = source["locator"].lower()
        self.assertNotIn(":\\", locator)
        self.assertNotIn("d:/", locator)
        self.assertNotIn(".pdf", locator)
        self.assertIn("not committed", locator)

    def test_scan_geometry_is_never_authoritative(self) -> None:
        page_rule = next(
            rule for rule in self.profile["rules"] if rule["id"] == "FMT-PAGE-201"
        )
        properties = page_rule["properties"]
        self.assertFalse(properties["scan_geometry_authoritative"])
        self.assertEqual("pending_qa", properties["finished_page_size"])
        self.assertEqual("pending_qa", properties["margins_and_gutter"])

    def test_exact_layout_values_remain_blocked(self) -> None:
        unanswered_blocking = [
            question
            for question in self.profile["open_questions"]
            if question["blocking"] and not question.get("answer", "").strip()
        ]
        self.assertGreaterEqual(len(unanswered_blocking), 6)

    def test_sample_specific_differences_require_override(self) -> None:
        guarded_rules = {
            rule["id"]: rule
            for rule in self.profile["rules"]
            if rule["properties"].get("publisher_specific_override_required")
        }
        self.assertEqual(
            {"FMT-SECTION-203", "FMT-BIB-201"},
            set(guarded_rules),
        )
        for rule in guarded_rules.values():
            note = rule["notes"]
            self.assertTrue("QA" in note or "等待决定" in note, rule["id"])


if __name__ == "__main__":
    unittest.main()
