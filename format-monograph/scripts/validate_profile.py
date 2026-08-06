#!/usr/bin/env python3
"""Validate a format profile against schema and semantic safety rules."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker

from _common import (
    FormatMonographError,
    load_json,
    profile_schema_path,
    unsupported_properties,
)


def semantic_errors(profile: dict) -> list[str]:
    errors: list[str] = []
    source_ids = [source["id"] for source in profile.get("sources", [])]
    rule_ids = [rule["id"] for rule in profile.get("rules", [])]

    if len(source_ids) != len(set(source_ids)):
        errors.append("Source IDs must be unique.")
    if len(rule_ids) != len(set(rule_ids)):
        errors.append("Rule IDs must be unique.")

    known_sources = set(source_ids)
    known_rules = set(rule_ids)
    for rule in profile.get("rules", []):
        missing = sorted(set(rule["source_ids"]) - known_sources)
        if missing:
            errors.append(f"Rule {rule['id']} references unknown sources: {', '.join(missing)}")
        unsupported = unsupported_properties(rule)
        if rule["application"] == "automatic" and unsupported:
            errors.append(
                f"Automatic rule {rule['id']} has unsupported properties: {', '.join(unsupported)}"
            )

    for conflict in profile.get("conflicts", []):
        missing = sorted(set(conflict["rule_ids"]) - known_rules)
        if missing:
            errors.append(
                f"Conflict {conflict['id']} references unknown rules: {', '.join(missing)}"
            )
        if conflict["status"] == "resolved" and not conflict.get("resolution", "").strip():
            errors.append(f"Resolved conflict {conflict['id']} requires a resolution.")

    approval = profile.get("approval", {})
    if approval.get("status") == "approved":
        if not approval.get("approved_by"):
            errors.append("Approved profile requires approval.approved_by.")
        if not approval.get("approved_at"):
            errors.append("Approved profile requires approval.approved_at.")
        for conflict in profile.get("conflicts", []):
            if conflict["status"] == "open":
                errors.append(f"Approved profile cannot contain open conflict {conflict['id']}.")
        for question in profile.get("open_questions", []):
            if question["blocking"] and not question.get("answer", "").strip():
                errors.append(
                    f"Approved profile has unanswered blocking question {question['id']}."
                )
        for rule in profile.get("rules", []):
            if rule["status"] == "draft":
                errors.append(f"Approved profile cannot contain draft rule {rule['id']}.")

    return errors


def validate(profile_path: Path) -> tuple[list[str], dict]:
    profile = load_json(profile_path)
    schema = load_json(profile_schema_path())
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = [
        f"{'/'.join(str(item) for item in error.absolute_path) or '<root>'}: {error.message}"
        for error in sorted(validator.iter_errors(profile), key=lambda item: list(item.path))
    ]
    if not errors:
        errors.extend(semantic_errors(profile))
    return errors, profile


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("profile", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        errors, profile = validate(args.profile)
    except FormatMonographError as exc:
        errors, profile = [str(exc)], {}

    result = {
        "valid": not errors,
        "profile_id": profile.get("profile_id"),
        "approval_status": profile.get("approval", {}).get("status"),
        "errors": errors,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif errors:
        print("Profile validation failed:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
    else:
        print(
            f"Profile {result['profile_id']} is valid "
            f"(approval: {result['approval_status']})."
        )
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
