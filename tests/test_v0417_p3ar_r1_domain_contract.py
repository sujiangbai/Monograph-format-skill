from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import subprocess
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from typing import Any
from unittest import mock


REPO = Path(__file__).resolve().parents[1]
BASE_COMMIT = "b4a6d19ecfa71e2844782812d4011cb7ce5ff3be"
CONTRACT_PATH = REPO / "tests/fixtures/v0417/p3a_r/r1-domain-contract.json"
REPORT_PATH = REPO / "tests/fixtures/v0417/p3a_r/r1-domain-report.md"
TEST_PATH = REPO / "tests/test_v0417_p3ar_r1_domain_contract.py"

SEMANTIC_PATH = "docs/plans/v0.4.0-profile-system-and-long-document-reliability.md"
AUTHORITY_INPUTS = [
    (SEMANTIC_PATH, "1d5b43979987744b57a7b6b686f2f3cbbffc3f0b", "semantic_authority"),
    ("docs/plans/v0.4.1-profile-foundation-and-grouped-execution.md", "4cb9119ccc91ec39876590a26ca0d7e8745c6d3b", "lifecycle_proof"),
    ("docs/plans/v0.4.1.1-p2-contract-correction.md", "5ccfa31964672c4409caa7f0c4a7b50e6fd060a4", "lifecycle_proof"),
    ("docs/plans/v0.4.1.2-p3-capability-and-asset-freeze.md", "e6716765534384f2e91922907b160166971cb7f7", "lifecycle_proof"),
    ("docs/plans/v0.4.1.3-p3a-c2-performance-contract-correction.md", "56d389d1d2c5ccfa5457fd12f11e00f1260d28cd", "lifecycle_proof"),
    ("docs/plans/v0.4.1.4-horizontal-foundation-and-staged-vertical-validation.md", "28b3c8658203050fec97201b565fbc5ecae1b0aa", "lifecycle_proof"),
    ("docs/plans/v0.4.1.5-append-only-production-registry-succession.md", "3649f216dcd8af9deb0009697b0cd4defdd37e2b", "lifecycle_proof"),
    ("docs/plans/v0.4.1.6-registry-semantic-closure-and-matrix-succession.md", "7e3c5503fb73cb9e57b6f2ef3e65fadc7d56f4fa", "lifecycle_proof"),
]
V0417_PATH = "docs/plans/v0.4.1.7-r1-semantic-domain-completion-and-r2-restart.md"
V0417_OID = "6037484a7c8fb711ba4d27deb46f3790a49e1b81"
V0418_PATH = "docs/plans/v0.4.1.8-atomic-ownership-authority-correction.md"
V0418_OID = "abe29ab6b12126fa85f839471a35a4d1bb506eb7"
V0418_RAW_SHA256 = "3d0e63ded3f64eecc279056d1a6f2d193cc452c38cc6143d1c7cc9c22f309cfc"
V0418_CANONICAL_SHA256 = "26c04207fce751b600609e550b7369ed46b48bf99e90f157b3dcf975802653e6"

ALLOWED_KINDS = {
    "property",
    "constraint",
    "capability",
    "later_rule",
    "non_registry_contract",
}
DOMAIN_SCOPES = {
    "project_closed",
    "source_preserved",
    "external_versioned",
    "project_open_constrained",
}
ALLOWED_PATHS = {
    "tests/fixtures/v0417/p3a_r/r1-domain-contract.json",
    "tests/fixtures/v0417/p3a_r/r1-domain-report.md",
    "tests/test_v0417_p3ar_r1_domain_contract.py",
}
FIXED_CHECKOUT_BLOB_OIDS = {
    "tests/fixtures/v0417/p3a_r/r1-domain-contract.json": (
        "1eef91123ecc8bfaa95851d88bfdabd27d8470bc"
    ),
    "tests/fixtures/v0417/p3a_r/r1-domain-report.md": (
        "aa4ff0f0ad6d65cd43a0c9105dc63f8bbceab82c"
    ),
}
FROZEN_AUTHORITY_BLOB_OIDS = {
    **{path: oid for path, oid, _role in AUTHORITY_INPUTS},
    V0417_PATH: V0417_OID,
    V0418_PATH: V0418_OID,
}
GITHUB_EVENT_MAX_BYTES = 2 * 1024 * 1024
CANONICAL_REPOSITORY = "sujiangbai/Monograph-format-skill"
COMMIT_SINGLE_LINE_HEADERS = {
    b"author",
    b"committer",
    b"encoding",
}
COMMIT_MULTILINE_HEADERS = {
    b"gpgsig",
    b"gpgsig-sha256",
    b"mergetag",
}
GENERIC_PLACEHOLDERS = {
    "atomic_typed_authority_contract",
    "typed_authority_clause_set",
    "authority-stated numeric bounds only",
    "selected by explicit decision-level semantic review",
    "exact_authority_clause",
    "typed-value-plus-unit-and-provenance",
    "typed-field-equality",
    "typed-source-document-property",
}
REFERENCE_RELATIONS = {
    "capability_refs": "executed_or_audited_by",
    "registry_property_ids": "constrains",
    "capability_dependencies": "gates",
    "external_obligations": "preserves_for_or_routes_to",
}
NA_REASON_EXPLANATIONS = {
    "capability_refs": (
        "no_separate_capability_dependency",
        "The frozen atomic property has no separate executable or auditor capability.",
    ),
    "registry_property_ids": (
        "no_registry_property_dependency",
        "The frozen atomic constraint does not constrain a registry property entity.",
    ),
    "capability_dependencies": (
        "no_capability_dependency",
        "The frozen atomic constraint has no executable or auditor capability dependency.",
    ),
    "external_obligations": (
        "no_external_obligation_dependency",
        "The frozen atomic constraint has no later-rule or non-registry obligation dependency.",
    ),
}


def _property_target(decision_id: str, fields: list[str]) -> dict[str, Any]:
    return {
        "entity_id": f"r1-domain/{decision_id}:registry",
        "contract_kind": "property",
        "field_paths": [f"definition.value_schema.fields.{field}" for field in fields],
    }


BODY_FONT_TARGET = _property_target(
    "V040-F-001", ["cjk_font", "latin_font", "font_size", "bold"]
)
BODY_STYLE_BUNDLE = [
    BODY_FONT_TARGET,
    _property_target(
        "V040-F-002", ["alignment", "last_line_alignment", "distributed_alignment"]
    ),
    _property_target("V040-F-003", ["minimum_line_spacing", "fixed_line_spacing"]),
    _property_target(
        "V040-F-004",
        [
            "first_line_indent",
            "left_indent",
            "right_indent",
            "space_before",
            "space_after",
            "synthetic_space_or_tab_indent",
        ],
    ),
    _property_target(
        "V040-F-005", ["widow_orphan_control", "keep_with_next", "keep_together"]
    ),
]
HEADING_HIERARCHY_BUNDLE = [
    _property_target("V040-E-001", ["cjk_font", "latin_and_digit_font"]),
    _property_target(
        "V040-E-002",
        [
            "font_size",
            "bold",
            "alignment",
            "minimum_line_spacing",
            "space_before",
            "space_after",
        ],
    ),
    _property_target(
        "V040-E-003",
        [
            "font_size",
            "bold",
            "alignment",
            "minimum_line_spacing",
            "space_before",
            "space_after",
        ],
    ),
    _property_target(
        "V040-E-004",
        [
            "level_3_font_size",
            "level_3_bold",
            "level_3_minimum_line_spacing",
            "level_3_space_before",
            "level_3_space_after",
            "level_4_font_size",
            "level_4_bold",
            "level_4_minimum_line_spacing",
            "level_4_space_before",
            "level_4_space_after",
            "alignment",
        ],
    ),
    _property_target(
        "V040-E-005",
        [
            "first_line_indent",
            "left_indent",
            "right_indent",
            "hanging_indent",
            "keep_with_next",
            "keep_together",
            "insert_blank_paragraph",
        ],
    ),
    {
        "entity_id": "r1-domain/V040-E-007:registry",
        "contract_kind": "constraint",
        "field_paths": ["definition.decidable_invariants"],
    },
]


def _reference_domain(relation: str, targets: list[dict[str, Any]]) -> dict[str, Any]:
    return {"kind": "reference", "relation": relation, "targets": targets}


SEMANTIC_HEADING_LEVEL_FIELDS = {
    1: (
        "V040-E-002",
        [
            "font_size",
            "bold",
            "alignment",
            "minimum_line_spacing",
            "space_before",
            "space_after",
        ],
    ),
    2: (
        "V040-E-003",
        [
            "font_size",
            "bold",
            "alignment",
            "minimum_line_spacing",
            "space_before",
            "space_after",
        ],
    ),
    3: (
        "V040-E-004",
        [
            "level_3_font_size",
            "level_3_bold",
            "level_3_minimum_line_spacing",
            "level_3_space_before",
            "level_3_space_after",
            "alignment",
        ],
    ),
    4: (
        "V040-E-004",
        [
            "level_4_font_size",
            "level_4_bold",
            "level_4_minimum_line_spacing",
            "level_4_space_before",
            "level_4_space_after",
            "alignment",
        ],
    ),
}
SEMANTIC_HEADING_REFERENCE_KEYS = {
    ("V040-T-001", "heading_level"),
    ("V040-U-001", "heading_level"),
    ("V040-V-002", "heading_level"),
    ("V040-W-002", "title_heading_level"),
    ("V040-X-001", "title_heading_level"),
}
SEMANTIC_HEADING_APPROVAL_SOURCES = [
    "frozen_semantic_artifact",
    "user_approved_artifact",
    "publisher_approved_artifact",
]


def _semantic_level_mappings() -> list[dict[str, Any]]:
    return [
        {
            "semantic_level": level,
            **_property_target(decision_id, fields),
        }
        for level, (decision_id, fields) in SEMANTIC_HEADING_LEVEL_FIELDS.items()
    ]


def _semantic_level_targets(
    mappings: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    targets: dict[str, dict[str, Any]] = {}
    for mapping in mappings:
        target_id = mapping["entity_id"]
        target = targets.setdefault(
            target_id,
            {
                "entity_id": target_id,
                "contract_kind": mapping["contract_kind"],
                "field_paths": [],
            },
        )
        for path in mapping["field_paths"]:
            if path not in target["field_paths"]:
                target["field_paths"].append(path)
    return list(targets.values())


def _semantic_level_reference(decision_id: str) -> dict[str, Any]:
    mappings = _semantic_level_mappings()
    return {
        "kind": "reference",
        "relation": "maps_approved_semantic_level_to_heading_style",
        "targets": _semantic_level_targets(mappings),
        "value_contract": {
            "version": 1,
            "kind": "approved_semantic_heading_level_mapping",
            "approved_input": {
                "value_shape": {
                    "type": "object",
                    "additional_fields": False,
                    "required_fields": [
                        "semantic_level",
                        "approval_source",
                        "approval_status",
                        "approval_artifact_sha256",
                    ],
                    "properties": [
                        {
                            "name": "semantic_level",
                            "type": "integer",
                            "enum": [1, 2, 3, 4],
                        },
                        {
                            "name": "approval_source",
                            "type": "string",
                            "enum": SEMANTIC_HEADING_APPROVAL_SOURCES,
                        },
                        {
                            "name": "approval_status",
                            "type": "string",
                            "enum": ["approved"],
                        },
                        {
                            "name": "approval_artifact_sha256",
                            "type": "string",
                            "pattern": "^[0-9a-f]{64}$",
                        },
                    ],
                },
                "approved_behavior": "map_exact_declared_level",
                "ambiguous_behavior": "blocked_qa",
                "missing_behavior": "reject",
                "unknown_level_behavior": "reject",
            },
            "level_mappings": mappings,
            "authority_boundary": {
                "semantic_authority": "V0.4.0-frozen-semantic-authority",
                "source_decision_id": decision_id,
                "selection_must_come_from": SEMANTIC_HEADING_APPROVAL_SOURCES,
                "module_may_fix_role_level": False,
                "module_may_invent_or_expand_level": False,
            },
        },
        "consumer_gate": {
            "relation": "consumes_approved_semantic_heading_level",
            "entity_id": f"r1-domain/{decision_id}:execution-p5b",
            "contract_kind": "capability",
            "field_paths": ["definition.input_boundary"],
        },
    }


N007_COMPATIBLE_BORDER_FIELDS = {
    "V040-N-003": [
        "top_border",
        "bottom_border",
        "header_separator",
        "vertical_borders",
    ],
    "V040-N-004": ["outer_border", "inner_border", "hierarchy_lines"],
    "V040-N-006": ["outer_border", "inner_border"],
}
N007_VALUE_BRANCHES = [
    {
        "owning_table_semantic": "simple_data_table",
        "source_label": "简单数据表",
        **_property_target("V040-N-003", N007_COMPATIBLE_BORDER_FIELDS["V040-N-003"]),
    },
    {
        "owning_table_semantic": "multilevel_header_table",
        "source_label": "多级表头表",
        **_property_target("V040-N-004", N007_COMPATIBLE_BORDER_FIELDS["V040-N-004"]),
    },
    {
        "owning_table_semantic": "text_comparison_table",
        "source_label": "文字对照表",
        **_property_target("V040-N-006", N007_COMPATIBLE_BORDER_FIELDS["V040-N-006"]),
    },
]


def _n007_value_targets(branches: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "entity_id": branch["entity_id"],
            "contract_kind": branch["contract_kind"],
            "field_paths": branch["field_paths"],
        }
        for branch in branches
    ]


def _n007_border_reference() -> dict[str, Any]:
    return {
        "kind": "reference",
        "relation": "selects_approved_table_style_from",
        "targets": _n007_value_targets(N007_VALUE_BRANCHES),
        "selection_contract": {
            "version": 1,
            "kind": "approved_owning_table_semantic_selection",
            "approved_input": {
                "value_shape": {
                    "type": "object",
                    "additional_fields": False,
                    "required_fields": [
                        "owning_table_semantic",
                        "approval_source",
                        "approval_status",
                        "approval_artifact_sha256",
                    ],
                    "properties": [
                        {
                            "name": "owning_table_semantic",
                            "type": "string",
                            "enum": [
                                "simple_data_table",
                                "multilevel_header_table",
                                "engineering_parameter_table",
                                "text_comparison_table",
                            ],
                        },
                        {
                            "name": "approval_source",
                            "type": "string",
                            "enum": SEMANTIC_HEADING_APPROVAL_SOURCES,
                        },
                        {
                            "name": "approval_status",
                            "type": "string",
                            "enum": ["approved"],
                        },
                        {
                            "name": "approval_artifact_sha256",
                            "type": "string",
                            "pattern": "^[0-9a-f]{64}$",
                        },
                    ],
                },
                "matched_value_behavior": "emit_exact_compatible_property_value",
                "ambiguous_behavior": "blocked_qa",
                "missing_behavior": "reject",
                "unknown_behavior": "reject",
            },
            "value_branches": N007_VALUE_BRANCHES,
            "authority_boundary": {
                "semantic_authority": "V0.4.0-frozen-semantic-authority",
                "source_decision_id": "V040-N-007",
                "selection_must_come_from": SEMANTIC_HEADING_APPROVAL_SOURCES,
                "user_or_publisher_approval_required": True,
                "module_may_classify_without_approval": False,
                "content_type_may_override_owning_table_semantic": False,
            },
        },
        "deferred_non_value_branch": {
            "when": {
                "owning_table_semantic": "engineering_parameter_table",
                "source_label": "工程参数表",
            },
            "target": {
                "entity_id": "r1-domain/V040-N-005:future-primary",
                "contract_kind": "later_rule",
                "field_paths": [
                    "definition.input_boundary",
                    "definition.output_boundary",
                    "definition.preserved_original_obligation",
                    "definition.target_stage",
                ],
                "source_authority": {
                    "semantic_authority": "V0.4.0-frozen-semantic-authority",
                    "source_decision_id": "V040-N-005",
                    "path": SEMANTIC_PATH,
                    "git_blob_oid": AUTHORITY_INPUTS[0][1],
                    "literal_sha256": "a8b33d6a142d6cfb4a1883776f729b1ff8d1be229407bbbf18865c1682f9a2ea",
                },
            },
            "current_value_status": "deferred",
            "blocked_to": "V0.4.2",
            "may_emit_value": False,
            "may_invent_default": False,
            "may_execute": False,
            "approval_boundary": {
                "user_or_publisher_approval_required": True,
                "confirmed_domain_classification_required": True,
                "module_may_self_authorize": False,
            },
        },
    }


FIELD_REFERENCE_EXPECTATIONS: dict[tuple[str, str], dict[str, Any]] = {
    ("V040-I-002", "font_and_size"): _reference_domain(
        "inherits_declared_value_from", [BODY_FONT_TARGET]
    ),
    ("V040-I-002", "minimum_line_spacing"): _reference_domain(
        "inherits_declared_value_from",
        [_property_target("V040-F-003", ["minimum_line_spacing"])],
    ),
    ("V040-J-001", "complete_paragraph_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
    ("V040-J-003", "font"): _reference_domain(
        "inherits_declared_value_from",
        [_property_target("V040-F-001", ["cjk_font", "latin_font", "bold"])],
    ),
    ("V040-J-003", "font_size"): _reference_domain(
        "inherits_declared_value_from",
        [_property_target("V040-F-001", ["font_size"])],
    ),
    ("V040-J-003", "line_spacing"): _reference_domain(
        "inherits_declared_value_from",
        [_property_target("V040-F-003", ["minimum_line_spacing", "fixed_line_spacing"])],
    ),
    ("V040-K-006", "paragraph_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
    ("V040-L-003", "font_context"): _reference_domain(
        "inherits_declared_value_from", [BODY_FONT_TARGET]
    ),
    ("V040-N-007", "border_style"): _n007_border_reference(),
    ("V040-N-014", "trailing_spacing"): _reference_domain(
        "inherits_declared_value_from",
        [_property_target("V040-M-007", ["space_after", "insert_blank_paragraph"])],
    ),
    ("V040-O-003", "later_paragraph_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
    ("V040-P-002", "content_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
    ("V040-R-003", "font"): _reference_domain(
        "inherits_declared_value_from", [BODY_FONT_TARGET]
    ),
    ("V040-T-001", "heading_level"): _semantic_level_reference("V040-T-001"),
    ("V040-U-001", "heading_level"): _semantic_level_reference("V040-U-001"),
    ("V040-U-002", "internal_heading_levels"): _reference_domain(
        "inherits_approved_heading_hierarchy_from", HEADING_HIERARCHY_BUNDLE
    ),
    ("V040-U-003", "paragraph_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
    ("V040-V-002", "heading_level"): _semantic_level_reference("V040-V-002"),
    ("V040-W-002", "title_heading_level"): _semantic_level_reference("V040-W-002"),
    ("V040-X-001", "title_heading_level"): _semantic_level_reference("V040-X-001"),
    ("V040-X-001", "body_style"): _reference_domain(
        "inherits_declared_value_from", BODY_STYLE_BUNDLE
    ),
}

FIELD_PRESERVATION_SELECTORS: dict[tuple[str, str, str], str] = {
    ("V040-A-001", "paper_width", "direct"): "w:pgSz/@w:w",
    ("V040-A-001", "paper_height", "direct"): "w:pgSz/@w:h",
    ("V040-A-001", "orientation", "direct"): "w:pgSz/@w:orient",
    ("V040-A-001", "section_boundary", "direct"): "w:sectPr position",
    ("V040-B-005", "vertical_alignment", "case:complex_layout"): "source title-page vertical alignment",
    ("V040-B-005", "complex_relative_layout", "direct"): "source title-page object positions",
    ("V040-N-001", "intentional_floating_or_layout_table", "direct"): "source table placement",
    ("V040-N-009", "page_layout_table", "direct"): "complete source table structure",
    ("V040-N-012", "column_width_ratios", "direct"): "w:tblGrid and w:tcW",
    ("V040-N-012", "natural_table_width", "direct"): "source table width",
    ("V040-O-004", "override", "case:otherwise"): "source theorem numbering",
    ("V040-V-004", "level_indents", "case:not_confirmed"): "source index hierarchy indentation",
    ("V040-V-007", "column_count", "direct"): "w:cols/@w:num",
    ("V040-V-007", "column_spacing", "direct"): "w:cols/@w:space",
    ("V040-V-007", "section_boundaries", "direct"): "index section w:sectPr",
    ("V040-W-002", "container_kind", "direct"): "source paragraph-or-table structure",
    ("V040-X-002", "alignment", "case:distinct_source_layout"): "source signature/date alignment",
    ("V040-Y-001", "presentation", "case:inline_label"): "source abstract inline label presentation",
    ("V040-Y-002", "english_indent", "direct"): "source English abstract indentation",
}
OPEN_NESTED_PRESERVATION_DECISIONS = {
    "V040-B-005",
    "V040-N-001",
    "V040-N-009",
    "V040-O-004",
    "V040-V-004",
    "V040-W-002",
    "V040-X-002",
    "V040-Y-001",
    "V040-Y-002",
}
A001_PAPER_INVARIANT = {
    "invariant_id": "V040-A-001:paper-geometry-preservation",
    "expression": {
        "operator": "preserve_source_paper_geometry",
        "width_field": "paper_width",
        "height_field": "paper_height",
        "orientation_field": "orientation",
        "section_field": "section_boundary",
        "canonical_section_field": "sectPr_boundary",
        "section_mapping": "one_to_one",
        "unit": "twentieth_of_a_point",
        "unit_definition": "one twentieth of one typographic point (twip)",
        "width_height_validity": {"finite": True, "greater_than": 0},
        "custom_paper": "preserve_exact_source_pair",
        "standard_paper_normalization": False,
        "swap_width_height": False,
        "supply_missing_default": False,
        "reject_only_for_nonstandard_size": False,
        "orientation_is_independent": True,
        "invalid_or_unknown_unit": "fail_closed",
    },
    "fields": ["paper_width", "paper_height", "orientation", "section_boundary"],
}
UNIT_DEFINITIONS = {
    "pt": ("typographic_length", "one point equals 1/72 inch"),
    "mm": ("physical_length", "one millimetre equals 1/1000 metre"),
    "character": (
        "character_count",
        "count of character-width indentation units defined by the source paragraph model",
    ),
    "count": ("count", "dimensionless discrete count"),
    "ratio": ("dimensionless_ratio", "dimensionless ratio preserved without unit conversion"),
    "source_native_length": (
        "length",
        "source-declared length representation preserved without conversion",
    ),
    "twentieth_of_a_point": (
        "length",
        "one twentieth of one typographic point (twip)",
    ),
}


def _validate_checkpoint_state(
    *,
    head: str,
    main: str,
    origin_main: str,
    status_lines: list[str],
    parent_oids: list[str],
    changed_paths: set[str],
) -> list[str]:
    errors: list[str] = []
    if main != BASE_COMMIT or origin_main != BASE_COMMIT:
        errors.append("main/origin-main drifted from the frozen entry base")
    if head == BASE_COMMIT:
        expected = {f"?? {path}" for path in ALLOWED_PATHS}
        if set(status_lines) != expected:
            errors.append("pre-checkpoint worktree is not the exact three untracked assets")
        if parent_oids or changed_paths:
            errors.append("pre-checkpoint state supplied candidate commit evidence")
    else:
        if status_lines:
            errors.append("candidate-checkpoint worktree is not clean")
        if parent_oids != [BASE_COMMIT]:
            errors.append("candidate checkpoint is not one commit directly on the base")
        if changed_paths != ALLOWED_PATHS:
            errors.append("candidate checkpoint does not change exactly the three assets")
    return errors


def _load_github_event(
    environ: Any | None = None,
) -> dict[str, Any] | None:
    environ = os.environ if environ is None else environ
    keys = (
        "GITHUB_ACTIONS",
        "GITHUB_EVENT_NAME",
        "GITHUB_EVENT_PATH",
        "GITHUB_SHA",
        "GITHUB_REPOSITORY",
        "GITHUB_REF",
    )
    values = {key: environ.get(key) for key in keys}
    if not any(value is not None for value in values.values()):
        return None
    if values["GITHUB_ACTIONS"] != "true":
        raise AssertionError("GitHub evidence adapter requires GITHUB_ACTIONS=true")
    event_name = values["GITHUB_EVENT_NAME"]
    if event_name not in {"pull_request", "push"}:
        raise AssertionError("GitHub evidence adapter event is unsupported")
    github_sha = values["GITHUB_SHA"]
    repository = values["GITHUB_REPOSITORY"]
    if not isinstance(github_sha, str) or not re.fullmatch(r"[0-9a-f]{40}", github_sha):
        raise AssertionError("GitHub checkout SHA is missing or invalid")
    if not isinstance(repository, str) or not re.fullmatch(
        r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository
    ):
        raise AssertionError("GitHub repository identity is missing or invalid")
    if repository != CANONICAL_REPOSITORY:
        raise AssertionError("GitHub repository identity differs from canonical authority")
    github_ref = values["GITHUB_REF"]
    if not isinstance(github_ref, str) or not github_ref.startswith("refs/"):
        raise AssertionError("GitHub ref identity is missing or invalid")
    raw_path = values["GITHUB_EVENT_PATH"]
    if not isinstance(raw_path, str) or not raw_path:
        raise AssertionError("GitHub event path is missing")
    event_path = Path(raw_path)
    if not event_path.is_absolute() or event_path.is_symlink() or not event_path.is_file():
        raise AssertionError("GitHub event file is not a regular absolute path")
    resolved_event = event_path.resolve(strict=True)
    try:
        resolved_event.relative_to(REPO.resolve())
    except ValueError:
        pass
    else:
        raise AssertionError("GitHub event file cannot come from the repository")
    if resolved_event.stat().st_size > GITHUB_EVENT_MAX_BYTES:
        raise AssertionError("GitHub event file exceeds the size limit")
    try:
        payload = json.loads(resolved_event.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AssertionError("GitHub event file is invalid") from exc
    if not isinstance(payload, dict):
        raise AssertionError("GitHub event payload must be an object")
    return {
        "payload": payload,
        "github_event_name": event_name,
        "github_sha": github_sha,
        "github_repository": repository,
        "github_ref": github_ref,
    }


def _validate_github_pull_request_checkout(
    event: Any,
    *,
    head: str,
    parent_oids: list[str],
    commit_message: bytes,
    status_lines: list[str],
    checkout_blob_oids: dict[str, str],
    working_blob_oids: dict[str, str],
) -> list[str]:
    errors: list[str] = []
    if not isinstance(event, dict) or set(event) != {
        "payload",
        "github_event_name",
        "github_sha",
        "github_repository",
        "github_ref",
    }:
        return ["GitHub pull-request evidence envelope is missing or not closed"]
    if event.get("github_event_name") != "pull_request":
        errors.append("GitHub pull-request evidence has the wrong event kind")
    if not re.fullmatch(r"refs/pull/[1-9][0-9]*/merge", str(event.get("github_ref"))):
        errors.append("GitHub pull-request ref is not a canonical merge ref")
    payload = event.get("payload")
    pull_request = payload.get("pull_request") if isinstance(payload, dict) else None
    repository = payload.get("repository") if isinstance(payload, dict) else None
    if not isinstance(pull_request, dict) or not isinstance(repository, dict):
        return ["GitHub pull-request payload is missing required objects"]
    base = pull_request.get("base")
    pr_head = pull_request.get("head")
    base_sha = base.get("sha") if isinstance(base, dict) else None
    head_sha = pr_head.get("sha") if isinstance(pr_head, dict) else None
    merge_sha = pull_request.get("merge_commit_sha")
    changed_files = pull_request.get("changed_files")
    commit_count = pull_request.get("commits")
    repository_name = repository.get("full_name")
    if base_sha != BASE_COMMIT:
        errors.append("GitHub pull-request base SHA differs from the frozen base")
    if not isinstance(head_sha, str) or not re.fullmatch(r"[0-9a-f]{40}", head_sha):
        errors.append("GitHub pull-request head SHA is missing or invalid")
    if merge_sha is not None and (
        not isinstance(merge_sha, str)
        or not re.fullmatch(r"[0-9a-f]{40}", merge_sha)
    ):
        errors.append("GitHub pull-request merge SHA is invalid")
    if (
        repository_name != CANONICAL_REPOSITORY
        or event.get("github_repository") != CANONICAL_REPOSITORY
    ):
        errors.append("GitHub repository identity conflicts with the event payload")
    if status_lines:
        errors.append("GitHub pull-request checkout worktree is not clean")

    if isinstance(head_sha, str) and head == head_sha:
        if event.get("github_sha") not in {head_sha, merge_sha}:
            errors.append("GitHub runner SHA conflicts with the direct head checkout")
        if parent_oids != [BASE_COMMIT]:
            errors.append("direct pull-request head is not one commit on the frozen base")
    elif (
        event.get("github_sha") == head
        and parent_oids == [BASE_COMMIT, head_sha]
        and commit_message
        == f"Merge {head_sha} into {BASE_COMMIT}\n".encode("ascii")
    ):
        pass
    else:
        if event.get("github_sha") == head and len(parent_oids) == 2:
            if parent_oids != [BASE_COMMIT, head_sha]:
                errors.append("synthetic pull-request merge parents conflict with event SHAs")
            else:
                errors.append(
                    "synthetic pull-request merge message conflicts with event SHAs"
                )
        else:
            errors.append("GitHub pull-request checkout shape is unknown")

    if type(changed_files) is not int or changed_files != len(ALLOWED_PATHS):
        errors.append("GitHub pull-request changed-file count is not exactly three")
    if type(commit_count) is not int or commit_count != 1:
        errors.append("GitHub pull-request candidate commit count is not exactly one")
    if set(checkout_blob_oids) != ALLOWED_PATHS:
        errors.append("GitHub checkout tree does not contain exactly the three allowed paths")
    if set(working_blob_oids) != ALLOWED_PATHS:
        errors.append("GitHub working-tree blob evidence is incomplete or contains extras")
    for path in ALLOWED_PATHS:
        checkout_oid = checkout_blob_oids.get(path)
        working_oid = working_blob_oids.get(path)
        if not isinstance(checkout_oid, str) or not re.fullmatch(
            r"[0-9a-f]{40}", checkout_oid
        ):
            errors.append(f"GitHub checkout tree blob OID is invalid: {path}")
        if checkout_oid != working_oid:
            errors.append(f"GitHub checkout tree blob differs from the clean file: {path}")
        fixed_oid = FIXED_CHECKOUT_BLOB_OIDS.get(path)
        if fixed_oid is not None and checkout_oid != fixed_oid:
            errors.append(f"GitHub checkout tree blob differs from the frozen fixture: {path}")
    return errors


def _is_allowed_push_ref(value: Any) -> bool:
    if value == "refs/heads/main":
        return True
    if not isinstance(value, str) or re.fullmatch(
        r"refs/heads/(?:skill|adapter|fix)/[A-Za-z0-9][A-Za-z0-9._/-]*",
        value,
    ) is None:
        return False
    tail = value.split("/", 3)[-1]
    segments = tail.split("/")
    return not (
        "//" in value
        or ".." in value
        or "@{" in value
        or value.endswith(("/", ".", ".lock"))
        or any(
            segment in {"", ".", ".."}
            or segment.startswith(".")
            or segment.endswith((".", ".lock"))
            for segment in segments
        )
    )


def _validate_integrated_tree(
    *,
    checkout_blob_oids: dict[str, str],
    working_blob_oids: dict[str, str],
    authority_blob_oids: dict[str, str],
    context: str,
) -> list[str]:
    errors: list[str] = []
    if set(checkout_blob_oids) != ALLOWED_PATHS:
        errors.append(f"{context} checkout tree does not contain exactly the governed paths")
    if set(working_blob_oids) != ALLOWED_PATHS:
        errors.append(f"{context} working-tree blob evidence is not closed")
    for path in ALLOWED_PATHS:
        checkout_oid = checkout_blob_oids.get(path)
        working_oid = working_blob_oids.get(path)
        if not isinstance(checkout_oid, str) or re.fullmatch(
            r"[0-9a-f]{40}", checkout_oid
        ) is None:
            errors.append(f"{context} checkout blob is invalid: {path}")
        if checkout_oid != working_oid:
            errors.append(f"{context} checkout blob differs from the clean file: {path}")
        fixed_oid = FIXED_CHECKOUT_BLOB_OIDS.get(path)
        if fixed_oid is not None and checkout_oid != fixed_oid:
            errors.append(f"{context} fixture blob differs from frozen evidence: {path}")
    if set(authority_blob_oids) != set(FROZEN_AUTHORITY_BLOB_OIDS):
        errors.append(f"{context} authority tree evidence is not closed")
    for path, expected_oid in FROZEN_AUTHORITY_BLOB_OIDS.items():
        if authority_blob_oids.get(path) != expected_oid:
            errors.append(f"{context} authority blob differs from frozen evidence: {path}")
    return errors


def _validate_github_push_checkout(
    event: Any,
    *,
    head: str,
    parent_oids: list[str],
    status_lines: list[str],
    checkout_blob_oids: dict[str, str],
    working_blob_oids: dict[str, str],
    authority_blob_oids: dict[str, str],
) -> list[str]:
    errors: list[str] = []
    envelope_keys = {
        "payload",
        "github_event_name",
        "github_sha",
        "github_repository",
        "github_ref",
    }
    if not isinstance(event, dict) or set(event) != envelope_keys:
        return ["GitHub push evidence envelope is missing or not closed"]
    if event.get("github_event_name") != "push":
        errors.append("GitHub push evidence has the wrong event kind")
    payload = event.get("payload")
    required = {
        "before",
        "after",
        "ref",
        "created",
        "deleted",
        "forced",
        "head_commit",
        "commits",
        "repository",
    }
    if not isinstance(payload, dict) or not required <= set(payload):
        return errors + ["GitHub push payload is missing required fields"]
    before = payload.get("before")
    after = payload.get("after")
    payload_ref = payload.get("ref")
    head_commit = payload.get("head_commit")
    commits = payload.get("commits")
    repository = payload.get("repository")
    repository_name = repository.get("full_name") if isinstance(repository, dict) else None
    if (
        repository_name != CANONICAL_REPOSITORY
        or event.get("github_repository") != CANONICAL_REPOSITORY
    ):
        errors.append("GitHub push repository identity conflicts with the event payload")
    if payload_ref != event.get("github_ref"):
        errors.append("GitHub push ref conflicts with the runner ref")
    if not _is_allowed_push_ref(payload_ref):
        errors.append("GitHub push ref is outside the allowed branch heads")
    for name, value in (("before", before), ("after", after)):
        if not isinstance(value, str) or re.fullmatch(r"[0-9a-f]{40}", value) is None:
            errors.append(f"GitHub push {name} SHA is missing or invalid")
        elif value == "0" * 40:
            errors.append(f"GitHub push {name} SHA cannot be all zero")
    if before == after:
        errors.append("GitHub push before and after SHAs must differ")
    if after != event.get("github_sha") or after != head:
        errors.append("GitHub push after SHA conflicts with runner or checkout HEAD")
    if not isinstance(head_commit, dict) or head_commit.get("id") != after:
        errors.append("GitHub push head_commit does not bind the after SHA")
    for flag in ("created", "deleted", "forced"):
        value = payload.get(flag)
        if type(value) is not bool or value:
            errors.append(f"GitHub push {flag} must be the boolean false")
    if not isinstance(commits, list):
        errors.append("GitHub push commits must be a list")
    else:
        commit_ids: list[str] = []
        for item in commits:
            commit_id = item.get("id") if isinstance(item, dict) else None
            if not isinstance(commit_id, str) or re.fullmatch(
                r"[0-9a-f]{40}", commit_id
            ) is None:
                errors.append("GitHub push commits contains a malformed commit")
                continue
            if commit_id in commit_ids:
                errors.append("GitHub push commits repeats a commit identity")
            commit_ids.append(commit_id)
        if after not in commit_ids:
            errors.append("GitHub push commits does not contain the after commit")
    if status_lines:
        errors.append("GitHub push checkout worktree is not clean")
    if parent_oids == [before]:
        pass
    elif len(parent_oids) == 2 and parent_oids[0] == before:
        second_parent = parent_oids[1]
        if (
            not isinstance(second_parent, str)
            or re.fullmatch(r"[0-9a-f]{40}", second_parent) is None
            or second_parent in {before, after}
        ):
            errors.append("GitHub push merge second parent is invalid")
    else:
        errors.append("GitHub push parent graph is not a linear or standard merge update")
    errors.extend(
        _validate_integrated_tree(
            checkout_blob_oids=checkout_blob_oids,
            working_blob_oids=working_blob_oids,
            authority_blob_oids=authority_blob_oids,
            context="GitHub push",
        )
    )
    return errors


def _validate_local_integrated_checkout(
    *,
    head: str,
    main: str,
    origin_main: str,
    status_lines: list[str],
    main_fixture_blob_oids: dict[str, str | None],
    main_checkout_blob_oids: dict[str, str | None],
    main_authority_blob_oids: dict[str, str | None],
    base_is_main_ancestor: bool,
    main_is_head_ancestor: bool,
    checkout_blob_oids: dict[str, str],
    working_blob_oids: dict[str, str],
    authority_blob_oids: dict[str, str],
) -> list[str]:
    errors: list[str] = []
    if main != origin_main:
        errors.append("local integrated main and origin/main differ")
    if main_fixture_blob_oids != FIXED_CHECKOUT_BLOB_OIDS:
        errors.append("local integrated main lacks the frozen R1 fixtures")
    if not base_is_main_ancestor:
        errors.append("local integrated main is not proven to descend from the frozen base")
    if set(main_checkout_blob_oids) != ALLOWED_PATHS:
        errors.append("local integrated main path evidence is not closed")
    for path in ALLOWED_PATHS:
        main_oid = main_checkout_blob_oids.get(path)
        if not isinstance(main_oid, str) or re.fullmatch(
            r"[0-9a-f]{40}", main_oid
        ) is None:
            errors.append(f"local integrated main path is missing or not regular: {path}")
        fixed_oid = FIXED_CHECKOUT_BLOB_OIDS.get(path)
        if fixed_oid is not None and main_oid != fixed_oid:
            errors.append(f"local integrated main fixture differs from frozen evidence: {path}")
    if set(main_authority_blob_oids) != set(FROZEN_AUTHORITY_BLOB_OIDS):
        errors.append("local integrated main authority evidence is not closed")
    for path, expected_oid in FROZEN_AUTHORITY_BLOB_OIDS.items():
        if main_authority_blob_oids.get(path) != expected_oid:
            errors.append(
                f"local integrated main authority differs from frozen evidence: {path}"
            )
    if status_lines:
        errors.append("local integrated worktree is not clean")
    if head != main and not main_is_head_ancestor:
        errors.append("local integrated HEAD is not main or a descendant of main")
    errors.extend(
        _validate_integrated_tree(
            checkout_blob_oids=checkout_blob_oids,
            working_blob_oids=working_blob_oids,
            authority_blob_oids=authority_blob_oids,
            context="local integrated",
        )
    )
    return errors


def _tree_blob_oid(revision: str, path: str) -> str:
    output = _git("ls-tree", revision, "--", path)
    match = re.fullmatch(rf"100644 blob ([0-9a-f]{{40}})\t{re.escape(path)}", output)
    if match is None:
        raise AssertionError(f"checkout tree path is missing or not a regular file: {path}")
    return match.group(1)


def _try_tree_blob_oid(revision: str, path: str) -> str | None:
    try:
        output = subprocess.check_output(
            ["git", "ls-tree", revision, "--", path],
            cwd=REPO,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except subprocess.CalledProcessError:
        return None
    match = re.fullmatch(rf"100644 blob ([0-9a-f]{{40}})\t{re.escape(path)}", output)
    return match.group(1) if match is not None else None


def _authority_tree_blob_oids(revision: str) -> dict[str, str | None]:
    return {
        path: _try_tree_blob_oid(revision, path)
        for path in sorted(FROZEN_AUTHORITY_BLOB_OIDS)
    }


def _try_working_blob_oid(path: str) -> str | None:
    try:
        return subprocess.check_output(
            ["git", "hash-object", "--", path],
            cwd=REPO,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except subprocess.CalledProcessError:
        return None


def _git_is_ancestor(ancestor: str, descendant: str) -> bool:
    result = subprocess.run(
        ["git", "merge-base", "--is-ancestor", ancestor, descendant],
        cwd=REPO,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return result.returncode == 0


def _parse_commit_parent_oids(raw_commit: bytes) -> list[str]:
    if not isinstance(raw_commit, bytes) or not raw_commit.endswith(b"\n"):
        raise AssertionError("raw commit object must be complete newline-terminated bytes")
    header_block, separator, _message = raw_commit.partition(b"\n\n")
    if not separator:
        raise AssertionError("raw commit object is missing the header boundary")
    lines = header_block.split(b"\n")
    if not lines or re.fullmatch(rb"tree [0-9a-f]{40}", lines[0]) is None:
        raise AssertionError("raw commit object has an invalid tree header")

    parents: list[str] = []
    index = 1
    while index < len(lines) and lines[index].startswith(b"parent "):
        match = re.fullmatch(rb"parent ([0-9a-f]{40})", lines[index])
        if match is None:
            raise AssertionError("raw commit object has an invalid parent header")
        parent = match.group(1).decode("ascii")
        if parent in parents:
            raise AssertionError("raw commit object repeats a parent")
        parents.append(parent)
        if len(parents) > 2:
            raise AssertionError("raw commit object has an unsupported parent shape")
        index += 1

    seen_headers: set[bytes] = set()
    current_header: bytes | None = None
    for line in lines[index:]:
        if line.startswith(b" "):
            if current_header not in COMMIT_MULTILINE_HEADERS:
                raise AssertionError("raw commit object has an invalid header continuation")
            continue
        match = re.fullmatch(rb"([a-z][a-z0-9-]*) (.+)", line)
        if match is None:
            raise AssertionError("raw commit object has a malformed header")
        current_header = match.group(1)
        if current_header in {b"tree", b"parent"}:
            raise AssertionError("raw commit object has a non-contiguous structural header")
        if current_header not in COMMIT_SINGLE_LINE_HEADERS | COMMIT_MULTILINE_HEADERS:
            raise AssertionError("raw commit object has an unsupported header")
        if current_header in seen_headers:
            raise AssertionError("raw commit object repeats a header")
        if current_header == b"committer" and b"author" not in seen_headers:
            raise AssertionError("raw commit object orders committer before author")
        seen_headers.add(current_header)
    if not {b"author", b"committer"} <= seen_headers:
        raise AssertionError("raw commit object lacks author or committer identity")
    return parents


def _commit_parent_oids(revision: str) -> list[str]:
    return _commit_object_evidence(revision)[0]


def _commit_object_evidence(revision: str) -> tuple[list[str], bytes]:
    try:
        raw_commit = subprocess.check_output(
            ["git", "cat-file", "commit", revision],
            cwd=REPO,
        )
    except subprocess.CalledProcessError as exc:
        raise AssertionError("raw commit object is unavailable") from exc
    parents = _parse_commit_parent_oids(raw_commit)
    _headers, separator, message = raw_commit.partition(b"\n\n")
    if not separator or not message:
        raise AssertionError("raw commit object is missing its message")
    return parents, message


def _runtime_checkpoint_evidence(
    environ: Any | None = None,
) -> tuple[str, list[str]]:
    event = _load_github_event(environ)
    if event is None:
        head = _git("rev-parse", "HEAD")
        main = _git("rev-parse", "main")
        origin_main = _git("rev-parse", "origin/main")
        status = _git("status", "--short", "--untracked-files=all").splitlines()
        if main == BASE_COMMIT and origin_main == BASE_COMMIT:
            parents = [] if head == BASE_COMMIT else _commit_parent_oids(head)
            changed = (
                set()
                if head == BASE_COMMIT
                else set(
                    _git("diff", "--name-only", f"{BASE_COMMIT}..{head}").splitlines()
                )
            )
            return (
                "local",
                _validate_checkpoint_state(
                    head=head,
                    main=main,
                    origin_main=origin_main,
                    status_lines=status,
                    parent_oids=parents,
                    changed_paths=changed,
                ),
            )
        main_fixtures = {
            path: _try_tree_blob_oid(main, path)
            for path in sorted(FIXED_CHECKOUT_BLOB_OIDS)
        }
        main_checkout_blobs = {
            path: _try_tree_blob_oid(main, path) for path in sorted(ALLOWED_PATHS)
        }
        checkout_blobs = {
            path: _try_tree_blob_oid("HEAD", path) for path in sorted(ALLOWED_PATHS)
        }
        working_blobs = {
            path: _try_working_blob_oid(path) for path in sorted(ALLOWED_PATHS)
        }
        return (
            "local_integrated",
            _validate_local_integrated_checkout(
                head=head,
                main=main,
                origin_main=origin_main,
                status_lines=status,
                main_fixture_blob_oids=main_fixtures,
                main_checkout_blob_oids=main_checkout_blobs,
                main_authority_blob_oids=_authority_tree_blob_oids(main),
                base_is_main_ancestor=_git_is_ancestor(BASE_COMMIT, main),
                main_is_head_ancestor=_git_is_ancestor(main, head),
                checkout_blob_oids=checkout_blobs,
                working_blob_oids=working_blobs,
                authority_blob_oids=_authority_tree_blob_oids("HEAD"),
            ),
        )
    head = _git("rev-parse", "HEAD")
    parents, commit_message = _commit_object_evidence("HEAD")
    status = _git("status", "--short", "--untracked-files=all").splitlines()
    checkout_blobs = {
        path: _tree_blob_oid("HEAD", path) for path in sorted(ALLOWED_PATHS)
    }
    working_blobs = {
        path: _git("hash-object", "--", path) for path in sorted(ALLOWED_PATHS)
    }
    if event.get("github_event_name") == "pull_request":
        return (
            "github_pull_request",
            _validate_github_pull_request_checkout(
                event,
                head=head,
                parent_oids=parents,
                commit_message=commit_message,
                status_lines=status,
                checkout_blob_oids=checkout_blobs,
                working_blob_oids=working_blobs,
            ),
        )
    return (
        "github_push_integrated",
        _validate_github_push_checkout(
            event,
            head=head,
            parent_oids=parents,
            status_lines=status,
            checkout_blob_oids=checkout_blobs,
            working_blob_oids=working_blobs,
            authority_blob_oids=_authority_tree_blob_oids("HEAD"),
        ),
    )


def _contains_generic_placeholder(value: Any) -> bool:
    serialized = json.dumps(value, ensure_ascii=False, sort_keys=True)
    return any(item in serialized for item in GENERIC_PLACEHOLDERS)


def _validate_reference_list(
    value: Any,
    *,
    allowed_targets: set[str],
    context: str,
    field_name: str,
    owner_entity: dict[str, Any],
) -> list[str]:
    if not isinstance(value, list) or not value:
        return [f"{context} must contain a typed reference or explicit non-applicability"]
    errors: list[str] = []
    seen: set[str] = set()
    for item in value:
        if not isinstance(item, dict):
            errors.append(f"{context} contains an untyped reference")
            continue
        if set(item) == {"entity_id", "relation"}:
            if item["entity_id"] not in allowed_targets:
                errors.append(f"{context} references the wrong or missing entity kind")
            if item["relation"] != REFERENCE_RELATIONS[field_name]:
                errors.append(f"{context} uses an unsupported relation")
            signature = f"entity:{item['entity_id']}"
        elif set(item) == {
            "applicability",
            "reason_code",
            "authority_ref_sha256",
            "contract_kind",
            "atomic_obligation_id",
            "explanation",
        }:
            expected_reason, expected_explanation = NA_REASON_EXPLANATIONS[field_name]
            semantic_sha = owner_entity["source_refs"][1]["literal_sha256"]
            if item["applicability"] != "not_applicable" or not re.fullmatch(
                r"[0-9a-f]{64}", str(item["authority_ref_sha256"])
            ) or item["reason_code"] != expected_reason or item[
                "explanation"
            ] != expected_explanation or item["authority_ref_sha256"] != semantic_sha or item[
                "contract_kind"
            ] != owner_entity["contract_kind"] or item[
                "atomic_obligation_id"
            ] != owner_entity["atomic_obligation_id"]:
                errors.append(f"{context} has invalid non-applicability evidence")
            signature = "not_applicable"
        else:
            errors.append(f"{context} reference shape is not closed")
            continue
        if signature in seen:
            errors.append(f"{context} contains a duplicate reference")
        seen.add(signature)
    return errors


def _git(*args: str) -> str:
    return subprocess.check_output(["git", *args], cwd=REPO, text=True).strip()


def _blob(oid: str) -> str:
    return subprocess.check_output(
        ["git", "cat-file", "blob", oid], cwd=REPO
    ).decode("utf-8")


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _canonical_bytes(value: Any) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode("utf-8")


def _extract_decisions() -> dict[str, dict[str, Any]]:
    text = _blob(AUTHORITY_INPUTS[0][1])
    rows: dict[str, dict[str, Any]] = {}
    pattern = re.compile(r"^\| (V040-[A-Z]-\d{3}) \|.*?(?=\r?$)", re.MULTILINE)
    for match in pattern.finditer(text):
        decision_id = match.group(1)
        if decision_id in rows:
            raise AssertionError(f"duplicate authority decision: {decision_id}")
        literal = match.group(0)
        cells = [cell.strip() for cell in literal.strip("|").split("|")]
        rows[decision_id] = {
            "decision_id": decision_id,
            "semantic_text": cells[1],
            "ownership_layer": cells[2],
            "override_boundary": cells[3],
            "mode_and_safety": cells[4],
            "required_tests": cells[5],
            "migration": cells[6],
            "source_ref": {
                "path": SEMANTIC_PATH,
                "git_blob_oid": AUTHORITY_INPUTS[0][1],
                "codepoint_start": match.start(),
                "codepoint_end": match.end(),
                "literal": literal,
                "literal_sha256": _sha256(literal.encode("utf-8")),
            },
        }
    if len(rows) != 171:
        raise AssertionError(f"expected 171 authority rows, got {len(rows)}")
    return rows


def _extract_ownership_precedence() -> dict[str, Any]:
    text = _blob(V0418_OID)
    match = re.search(r"```ownership-machine-json\n(.*?)```", text, re.DOTALL)
    if match is None:
        raise AssertionError("V0.4.1.8 ownership appendix missing")
    raw = match.group(1)
    appendix = json.loads(raw)
    raw_claim = re.search(r"Raw appendix SHA-256: `([0-9a-f]{64})`", text)
    canonical_claim = re.search(
        r"Canonical appendix SHA-256: `([0-9a-f]{64})`", text
    )
    if raw_claim is None or canonical_claim is None:
        raise AssertionError("V0.4.1.8 digest claims missing")
    if _sha256(raw.encode("utf-8")) != V0418_RAW_SHA256:
        raise AssertionError("V0.4.1.8 raw digest changed")
    if _sha256(_canonical_bytes(appendix)) != V0418_CANONICAL_SHA256:
        raise AssertionError("V0.4.1.8 canonical digest changed")
    if raw_claim.group(1) != V0418_RAW_SHA256:
        raise AssertionError("V0.4.1.8 raw digest claim changed")
    if canonical_claim.group(1) != V0418_CANONICAL_SHA256:
        raise AssertionError("V0.4.1.8 canonical digest claim changed")
    return appendix


def _validate_source_ref(ref: Any, blob_cache: dict[str, str]) -> list[str]:
    errors: list[str] = []
    required = {
        "path",
        "git_blob_oid",
        "codepoint_start",
        "codepoint_end",
        "literal",
        "literal_sha256",
    }
    if not isinstance(ref, dict) or set(ref) != required:
        return ["source reference shape is not closed"]
    oid = ref.get("git_blob_oid")
    if not isinstance(oid, str) or not re.fullmatch(r"[0-9a-f]{40}", oid):
        return ["source reference OID is invalid"]
    try:
        if oid not in blob_cache:
            blob_cache[oid] = _blob(oid)
        text = blob_cache[oid]
        start = ref["codepoint_start"]
        end = ref["codepoint_end"]
        if type(start) is not int or type(end) is not int or not (0 <= start < end <= len(text)):
            errors.append("source reference span is invalid")
        else:
            literal = text[start:end]
            if literal != ref["literal"]:
                errors.append("source reference literal differs from Git blob")
            if _sha256(literal.encode("utf-8")) != ref["literal_sha256"]:
                errors.append("source reference literal digest differs")
    except (OSError, subprocess.CalledProcessError, UnicodeError, TypeError, KeyError):
        errors.append("source reference cannot be reconstructed")
    return errors


def _validate_domain(
    definition: dict[str, Any],
    context: str,
    owner_entity: dict[str, Any] | None = None,
) -> list[str]:
    errors: list[str] = []
    scope = definition.get("domain_scope")
    if scope not in DOMAIN_SCOPES:
        return [f"{context} has invalid domain_scope"]
    domain = definition.get("domain_contract")
    if not isinstance(domain, dict):
        return [f"{context} domain_contract must be an object"]
    required_by_scope = {
        "project_closed": {"members", "member_source_refs", "closure_rule"},
        "source_preserved": {
            "source_system",
            "source_version",
            "source_type",
            "canonical_preservation_shape",
            "allowed_fields",
            "unit_and_finite_constraints",
            "round_trip_equivalence",
            "provenance",
            "validator_boundary",
        },
        "external_versioned": {
            "owner",
            "external_version",
            "locator",
            "digest",
            "project_mapping",
            "offline_validator_boundary",
        },
        "project_open_constrained": {
            "value_shape",
            "finite_rules",
            "range_rules",
            "cross_field_invariants",
        },
    }
    if set(domain) != required_by_scope[scope]:
        errors.append(f"{context} domain_contract fields do not close {scope}")
    elif scope == "project_closed":
        members = domain["members"]
        if not isinstance(members, list) or not members or len(members) != len(
            {json.dumps(item, ensure_ascii=False, sort_keys=True) for item in members}
        ):
            errors.append(f"{context} project_closed members are not complete/unique")
    elif scope == "source_preserved":
        if not isinstance(domain["allowed_fields"], list) or not domain["allowed_fields"]:
            errors.append(f"{context} source_preserved allowed_fields are empty")
        shape = domain["canonical_preservation_shape"]
        if not isinstance(shape, dict) or set(shape) != {
            "object_kind",
            "identity_fields",
            "payload_fields",
            "ordering_fields",
        }:
            errors.append(f"{context} source_preserved canonical shape is not object-specific")
        elif not shape["identity_fields"] or not shape["payload_fields"]:
            errors.append(f"{context} source_preserved identity/payload is empty")
        else:
            shape_fields = (
                shape["identity_fields"]
                + shape["payload_fields"]
                + shape["ordering_fields"]
            )
            if len(shape_fields) != len(set(shape_fields)) or domain["allowed_fields"] != shape_fields:
                errors.append(f"{context} source_preserved allowed fields differ from canonical shape")
        round_trip = domain["round_trip_equivalence"]
        if not isinstance(round_trip, dict) or set(round_trip) != {
            "identity_comparison",
            "payload_comparison",
            "ordering_comparison",
            "unknown_member_policy",
        }:
            errors.append(f"{context} source_preserved round-trip rule is generic")
        elif round_trip != {
            "identity_comparison": "exact",
            "payload_comparison": "canonical_xml_or_byte_exact",
            "ordering_comparison": "exact_sequence",
            "unknown_member_policy": "reject",
        }:
            errors.append(f"{context} source_preserved round-trip values are unsupported")
        units = domain["unit_and_finite_constraints"]
        if not isinstance(units, dict) or set(units) != {
            "numeric_fields",
            "field_units",
            "finite_numbers_only",
            "source_units_required",
        }:
            errors.append(f"{context} source_preserved numeric/unit closure is incomplete")
        else:
            numeric_fields = units["numeric_fields"]
            field_units = units["field_units"]
            if (
                not isinstance(numeric_fields, list)
                or len(numeric_fields) != len(set(numeric_fields))
                or not set(numeric_fields) <= set(domain["allowed_fields"])
                or units["finite_numbers_only"] is not True
                or units["source_units_required"] is not True
            ):
                errors.append(f"{context} source_preserved numeric fields are invalid")
            unit_fields = []
            for item in field_units if isinstance(field_units, list) else []:
                if not isinstance(item, dict) or set(item) != {
                    "field",
                    "dimension",
                    "canonical_unit",
                    "allowed_units",
                    "unit_definition",
                } or not item.get("dimension") or not item.get("allowed_units") or item.get(
                    "canonical_unit"
                ) not in item.get("allowed_units", []):
                    errors.append(f"{context} source_preserved field unit is invalid")
                    continue
                expected_definition = UNIT_DEFINITIONS.get(item["canonical_unit"])
                if expected_definition != (item["dimension"], item["unit_definition"]):
                    errors.append(f"{context} source_preserved unit definition is unsupported")
                unit_fields.append(item["field"])
            if unit_fields != numeric_fields:
                errors.append(f"{context} source_preserved unit fields differ from numeric fields")
        if owner_entity is not None:
            semantic_ref = owner_entity["source_refs"][1]
            expected_provenance = {
                "decision_id": owner_entity["source_decision_id"],
                "atomic_obligation_id": owner_entity["atomic_obligation_id"],
                "semantic_literal_sha256": semantic_ref["literal_sha256"],
            }
            if domain["source_system"] != "V0.4.0-frozen-semantic-authority":
                errors.append(f"{context} source_preserved source system changed")
            if domain["source_version"] != AUTHORITY_INPUTS[0][1]:
                errors.append(f"{context} source_preserved source version changed")
            if domain["source_type"] != shape.get("object_kind"):
                errors.append(f"{context} source_preserved source type differs from shape")
            if domain["provenance"] != expected_provenance:
                errors.append(f"{context} source_preserved provenance differs from entity source")
            if domain["validator_boundary"] != {
                "owner": "P3a-R-test-only",
                "writes_docx": False,
                "may_normalize_payload": False,
                "may_transfer_execution_ownership": False,
            }:
                errors.append(f"{context} source_preserved validator boundary changed")
    elif scope == "external_versioned":
        if not re.fullmatch(r"[0-9a-f]{64}", str(domain["digest"])):
            errors.append(f"{context} external digest is invalid")
    elif scope == "project_open_constrained":
        if not domain["finite_rules"] or not domain["range_rules"] or not domain["cross_field_invariants"]:
            errors.append(f"{context} open domain lacks finite/invariant closure")
    return errors


def _resolve_reference_path(entity: dict[str, Any], path: str) -> bool:
    field_prefix = "definition.value_schema.fields."
    if path.startswith(field_prefix):
        field_name = path[len(field_prefix) :]
        fields = entity.get("definition", {}).get("value_schema", {}).get("fields", [])
        return any(
            isinstance(field, dict) and field.get("name") == field_name for field in fields
        )
    current: Any = entity
    for part in path.split("."):
        if not isinstance(current, dict) or part not in current:
            return False
        current = current[part]
    return True


def _validate_semantic_heading_reference(
    domain: dict[str, Any],
    *,
    owner_entity: dict[str, Any],
    field_name: str,
    entity_by_id: dict[str, dict[str, Any]],
) -> list[str]:
    reference_key = (owner_entity["source_decision_id"], field_name)
    if reference_key not in SEMANTIC_HEADING_REFERENCE_KEYS:
        return []
    context = f"{owner_entity['entity_id']}.{field_name}"
    errors: list[str] = []
    expected = _semantic_level_reference(owner_entity["source_decision_id"])
    value_contract = domain.get("value_contract")
    if not isinstance(value_contract, dict) or set(value_contract) != {
        "version",
        "kind",
        "approved_input",
        "level_mappings",
        "authority_boundary",
    }:
        return [f"{context} semantic-level value contract is missing or not closed"]
    approved_input = value_contract.get("approved_input")
    if approved_input != expected["value_contract"]["approved_input"]:
        errors.append(f"{context} approved semantic-level input contract changed")
    authority_boundary = value_contract.get("authority_boundary")
    if authority_boundary != expected["value_contract"]["authority_boundary"]:
        errors.append(f"{context} semantic-level authority boundary changed")
    if value_contract.get("version") != 1 or value_contract.get("kind") != (
        "approved_semantic_heading_level_mapping"
    ):
        errors.append(f"{context} semantic-level value contract identity changed")

    mappings = value_contract.get("level_mappings")
    expected_mappings = _semantic_level_mappings()
    if mappings != expected_mappings:
        errors.append(f"{context} semantic-level mapping changed")
    if isinstance(mappings, list) and all(
        isinstance(item, dict)
        and set(item)
        == {"semantic_level", "entity_id", "contract_kind", "field_paths"}
        for item in mappings
    ):
        levels = [item["semantic_level"] for item in mappings]
        if levels != [1, 2, 3, 4] or len(levels) != len(set(levels)):
            errors.append(f"{context} semantic levels are not exactly 1 through 4")
        recomputed_targets = _semantic_level_targets(mappings)
        if domain.get("targets") != recomputed_targets:
            errors.append(f"{context} value targets differ from level mappings")
    else:
        errors.append(f"{context} semantic-level mappings are not closed")

    targets = domain.get("targets")
    if not isinstance(targets, list) or any(
        not isinstance(target, dict) or target.get("contract_kind") != "property"
        for target in targets
    ):
        errors.append(f"{context} value source must contain property entities only")

    consumer_gate = domain.get("consumer_gate")
    expected_gate = expected["consumer_gate"]
    if consumer_gate != expected_gate:
        errors.append(f"{context} execution consumer gate changed")
    if isinstance(consumer_gate, dict):
        gate_entity = entity_by_id.get(str(consumer_gate.get("entity_id")))
        gate_paths = consumer_gate.get("field_paths")
        if (
            gate_entity is None
            or gate_entity.get("contract_kind") != "capability"
            or not isinstance(gate_paths, list)
            or gate_paths != ["definition.input_boundary"]
            or not all(_resolve_reference_path(gate_entity, path) for path in gate_paths)
        ):
            errors.append(f"{context} execution consumer gate cannot be resolved")
    return errors


def _validate_n007_border_reference(
    domain: dict[str, Any],
    *,
    owner_entity: dict[str, Any],
    field_name: str,
    entity_by_id: dict[str, dict[str, Any]],
) -> list[str]:
    if (owner_entity["source_decision_id"], field_name) != (
        "V040-N-007",
        "border_style",
    ):
        return []
    context = f"{owner_entity['entity_id']}.{field_name}"
    errors: list[str] = []
    expected = _n007_border_reference()
    selection = domain.get("selection_contract")
    if not isinstance(selection, dict) or set(selection) != {
        "version",
        "kind",
        "approved_input",
        "value_branches",
        "authority_boundary",
    }:
        errors.append(f"{context} table-semantic selection contract is missing or not closed")
    else:
        if selection.get("version") != 1 or selection.get("kind") != (
            "approved_owning_table_semantic_selection"
        ):
            errors.append(f"{context} table-semantic selection identity changed")
        if selection.get("approved_input") != expected["selection_contract"][
            "approved_input"
        ]:
            errors.append(f"{context} table-semantic fail-closed input behavior changed")
        if selection.get("authority_boundary") != expected["selection_contract"][
            "authority_boundary"
        ]:
            errors.append(f"{context} table-semantic approval boundary changed")
        branches = selection.get("value_branches")
        if branches != N007_VALUE_BRANCHES:
            errors.append(f"{context} compatible border value branches changed")
        if isinstance(branches, list) and all(
            isinstance(branch, dict)
            and set(branch)
            == {
                "owning_table_semantic",
                "source_label",
                "entity_id",
                "contract_kind",
                "field_paths",
            }
            for branch in branches
        ):
            semantics = [branch["owning_table_semantic"] for branch in branches]
            if semantics != [
                "simple_data_table",
                "multilevel_header_table",
                "text_comparison_table",
            ] or len(semantics) != len(set(semantics)):
                errors.append(f"{context} compatible table semantic branches are not exact")
            allowed_paths_by_entity = {
                f"r1-domain/{decision_id}:registry": [
                    f"definition.value_schema.fields.{field}"
                    for field in allowed_fields
                ]
                for decision_id, allowed_fields in N007_COMPATIBLE_BORDER_FIELDS.items()
            }
            for branch in branches:
                if branch.get("field_paths") != allowed_paths_by_entity.get(
                    branch.get("entity_id")
                ):
                    errors.append(
                        f"{context} branch violates compatible-border-field allowlist"
                    )
            if domain.get("targets") != _n007_value_targets(branches):
                errors.append(f"{context} border value targets differ from compatible branches")
        else:
            errors.append(f"{context} compatible border branches are not closed")

    targets = domain.get("targets")
    if not isinstance(targets, list) or any(
        not isinstance(target, dict) or target.get("contract_kind") != "property"
        for target in targets
    ):
        errors.append(f"{context} immediate border values must come from properties only")

    deferred = domain.get("deferred_non_value_branch")
    expected_deferred = expected["deferred_non_value_branch"]
    if not isinstance(deferred, dict) or set(deferred) != set(expected_deferred):
        return errors + [f"{context} N005 deferred non-value branch is missing or not closed"]
    for key in (
        "when",
        "current_value_status",
        "blocked_to",
        "may_emit_value",
        "may_invent_default",
        "may_execute",
        "approval_boundary",
    ):
        if deferred.get(key) != expected_deferred[key]:
            errors.append(f"{context} N005 deferred branch {key} changed")
    deferred_target = deferred.get("target")
    expected_target = expected_deferred["target"]
    if deferred_target != expected_target:
        errors.append(f"{context} N005 deferred target or authority changed")
    if isinstance(deferred_target, dict):
        target_entity = entity_by_id.get(str(deferred_target.get("entity_id")))
        paths = deferred_target.get("field_paths")
        source_authority = deferred_target.get("source_authority")
        semantic_ref = (
            target_entity.get("source_refs", [None, None])[1]
            if isinstance(target_entity, dict)
            and len(target_entity.get("source_refs", [])) > 1
            else None
        )
        if (
            target_entity is None
            or target_entity.get("contract_kind") != "later_rule"
            or not isinstance(paths, list)
            or paths != expected_target["field_paths"]
            or not all(_resolve_reference_path(target_entity, path) for path in paths)
            or not isinstance(semantic_ref, dict)
            or not isinstance(source_authority, dict)
            or source_authority.get("source_decision_id")
            != target_entity.get("source_decision_id")
            or source_authority.get("path") != semantic_ref.get("path")
            or source_authority.get("git_blob_oid") != semantic_ref.get("git_blob_oid")
            or source_authority.get("literal_sha256")
            != semantic_ref.get("literal_sha256")
        ):
            errors.append(f"{context} N005 deferred authority cannot be resolved")
    return errors


def _validate_field_reference(
    domain: Any,
    *,
    owner_entity: dict[str, Any],
    field_name: str,
    entity_by_id: dict[str, dict[str, Any]],
) -> list[str]:
    context = f"{owner_entity['entity_id']}.{field_name}"
    expected = FIELD_REFERENCE_EXPECTATIONS.get(
        (owner_entity["source_decision_id"], field_name)
    )
    if expected is None:
        return [f"{context} has an undeclared reference"]
    errors: list[str] = []
    if domain != expected:
        errors.append(f"{context} differs from the frozen typed-reference declaration")
    if not isinstance(domain, dict) or set(domain) != set(expected):
        errors.append(f"{context} reference shape is not closed")
        if isinstance(domain, dict):
            errors.extend(
                _validate_semantic_heading_reference(
                    domain,
                    owner_entity=owner_entity,
                    field_name=field_name,
                    entity_by_id=entity_by_id,
                )
            )
            errors.extend(
                _validate_n007_border_reference(
                    domain,
                    owner_entity=owner_entity,
                    field_name=field_name,
                    entity_by_id=entity_by_id,
                )
            )
        return errors
    targets = domain.get("targets")
    if not isinstance(targets, list) or not targets:
        return errors + [f"{context} reference target bundle is empty"]
    seen_targets: set[str] = set()
    for target in targets:
        if not isinstance(target, dict) or set(target) != {
            "entity_id",
            "contract_kind",
            "field_paths",
        }:
            errors.append(f"{context} target declaration is not closed")
            continue
        target_id = target["entity_id"]
        if target_id in seen_targets:
            errors.append(f"{context} repeats a target entity")
        seen_targets.add(target_id)
        target_entity = entity_by_id.get(target_id)
        if target_entity is None:
            errors.append(f"{context} target entity is missing")
            continue
        if target_entity["contract_kind"] != target["contract_kind"]:
            errors.append(f"{context} target contract kind differs")
        paths = target["field_paths"]
        if not isinstance(paths, list) or not paths or len(paths) != len(set(paths)):
            errors.append(f"{context} target field paths are empty or duplicated")
            continue
        for path in paths:
            if not isinstance(path, str) or not _resolve_reference_path(target_entity, path):
                errors.append(f"{context} target field path cannot be resolved")
    errors.extend(
        _validate_semantic_heading_reference(
            domain,
            owner_entity=owner_entity,
            field_name=field_name,
            entity_by_id=entity_by_id,
        )
    )
    errors.extend(
        _validate_n007_border_reference(
            domain,
            owner_entity=owner_entity,
            field_name=field_name,
            entity_by_id=entity_by_id,
        )
    )
    return errors


def _validate_field_preservation(
    domain: Any,
    *,
    owner_entity: dict[str, Any],
    object_name: str,
    field_name: str,
    field_type: str,
    selector_tag: str,
) -> list[str]:
    context = f"{owner_entity['entity_id']}.{field_name}.{selector_tag}"
    if not isinstance(domain, dict) or domain.get("kind") != "source_preserved":
        return [f"{context} preservation branch is not typed"]
    numeric = field_type == "number"
    expected_domain_keys = {"kind", "preservation_contract"}
    if numeric:
        expected_domain_keys |= {
            "dimension",
            "canonical_unit",
            "allowed_units",
            "unit_definition",
            "finite",
        }
    errors: list[str] = []
    if set(domain) != expected_domain_keys:
        errors.append(f"{context} preservation domain fields are not closed")
    contract = domain.get("preservation_contract")
    required = {
        "source_system",
        "source_version",
        "source_type",
        "source_selector",
        "canonical_preservation_shape",
        "allowed_fields",
        "unit_and_finite_constraints",
        "round_trip_equivalence",
        "provenance",
        "validator_boundary",
    }
    if not isinstance(contract, dict) or set(contract) != required:
        return errors + [f"{context} preservation contract is incomplete"]
    selector = FIELD_PRESERVATION_SELECTORS.get(
        (owner_entity["source_decision_id"], field_name, selector_tag)
    )
    if selector is None or contract["source_selector"] != selector:
        errors.append(f"{context} preservation selector is unsupported")
    if contract["source_system"] != "V0.4.0-frozen-semantic-authority":
        errors.append(f"{context} preservation source system changed")
    if contract["source_version"] != AUTHORITY_INPUTS[0][1]:
        errors.append(f"{context} preservation source version changed")
    expected_type = f"{object_name}.{field_name}.source_value"
    if contract["source_type"] != expected_type:
        errors.append(f"{context} preservation source type changed")
    ordering = ["item_order"] if field_type in {"array", "object"} else []
    expected_shape = {
        "object_kind": expected_type,
        "identity_fields": ["owner_entity_id", "field_name"],
        "payload_fields": ["value"],
        "ordering_fields": ordering,
    }
    if contract["canonical_preservation_shape"] != expected_shape:
        errors.append(f"{context} preservation shape changed")
    expected_allowed = ["owner_entity_id", "field_name", "value", *ordering]
    if contract["allowed_fields"] != expected_allowed:
        errors.append(f"{context} preservation allowed fields changed")
    if numeric:
        expected_units = {
            "numeric_fields": ["value"],
            "field_units": [
                {
                    "field": "value",
                    "dimension": domain.get("dimension"),
                    "canonical_unit": domain.get("canonical_unit"),
                    "allowed_units": domain.get("allowed_units"),
                    "unit_definition": domain.get("unit_definition"),
                }
            ],
            "finite_numbers_only": True,
            "source_units_required": True,
        }
    else:
        expected_units = {
            "numeric_fields": [],
            "field_units": [],
            "finite_numbers_only": True,
            "source_units_required": False,
        }
    if contract["unit_and_finite_constraints"] != expected_units:
        errors.append(f"{context} preservation unit closure changed")
    payload_comparison = (
        "exact_scalar"
        if field_type in {"string", "number", "boolean"}
        else "canonical_xml_or_byte_exact"
    )
    expected_round_trip = {
        "identity_comparison": "exact",
        "payload_comparison": payload_comparison,
        "ordering_comparison": "exact_sequence" if ordering else "not_applicable",
        "unknown_member_policy": "reject",
    }
    if contract["round_trip_equivalence"] != expected_round_trip:
        errors.append(f"{context} preservation round-trip changed")
    if contract["provenance"] != {
        "decision_id": owner_entity["source_decision_id"],
        "atomic_obligation_id": owner_entity["atomic_obligation_id"],
        "field_name": field_name,
        "selector_tag": selector_tag,
        "semantic_literal_sha256": owner_entity["source_refs"][1]["literal_sha256"],
    }:
        errors.append(f"{context} preservation provenance changed")
    if contract["validator_boundary"] != {
        "owner": "P3a-R-test-only",
        "writes_docx": False,
        "may_normalize_payload": False,
        "may_transfer_execution_ownership": False,
    }:
        errors.append(f"{context} preservation boundary changed")
    if numeric:
        canonical = domain.get("canonical_unit")
        definition = UNIT_DEFINITIONS.get(canonical)
        if (
            domain.get("finite") is not True
            or definition is None
            or domain.get("dimension") != definition[0]
            or domain.get("unit_definition") != definition[1]
            or domain.get("allowed_units") != [canonical]
        ):
            errors.append(f"{context} numeric preservation unit is unsupported")
    return errors


def _walk_field_preservations(
    value: Any,
    *,
    owner_entity: dict[str, Any],
    object_name: str,
    field_name: str,
    field_type: str,
    selector_tag: str = "direct",
) -> tuple[list[str], set[tuple[str, str, str]]]:
    errors: list[str] = []
    observed: set[tuple[str, str, str]] = set()
    if value == "source_preserved":
        return [f"{owner_entity['entity_id']}.{field_name} contains bare source_preserved"], observed
    if isinstance(value, dict):
        if value.get("kind") == "source_preserved":
            observed.add((owner_entity["source_decision_id"], field_name, selector_tag))
            return (
                _validate_field_preservation(
                    value,
                    owner_entity=owner_entity,
                    object_name=object_name,
                    field_name=field_name,
                    field_type=field_type,
                    selector_tag=selector_tag,
                ),
                observed,
            )
        if value.get("kind") == "conditional":
            for case in value.get("cases", []):
                case_errors, case_observed = _walk_field_preservations(
                    case.get("value"),
                    owner_entity=owner_entity,
                    object_name=object_name,
                    field_name=field_name,
                    field_type=field_type,
                    selector_tag=f"case:{case.get('when')}",
                )
                errors.extend(case_errors)
                observed.update(case_observed)
    return errors, observed


def _validate_contract(
    contract: Any,
    decisions: dict[str, dict[str, Any]],
    ownership: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    if not isinstance(contract, dict):
        return ["contract must be an object"]
    required_top = {
        "schema_version",
        "status",
        "base_commit",
        "authority_inputs",
        "ownership_precedence",
        "serialization",
        "entities",
        "decision_dispositions",
        "unresolved_obligations",
        "summary",
        "non_authorizations",
    }
    if set(contract) != required_top:
        errors.append("top-level contract fields are not closed")
    if contract.get("schema_version") != "v0.4.1.7-r1-domain-contract-1":
        errors.append("schema_version is unsupported")
    if contract.get("status") != "test_only_candidate":
        errors.append("status is unsupported")
    if contract.get("base_commit") != BASE_COMMIT:
        errors.append("base commit changed")
    expected_authorities = [
        {"path": path, "git_blob_oid": oid, "permitted_contribution": role}
        for path, oid, role in AUTHORITY_INPUTS
    ]
    if contract.get("authority_inputs") != expected_authorities:
        errors.append("authority input identity/order changed")
    precedence = contract.get("ownership_precedence")
    expected_precedence = {
        "path": V0418_PATH,
        "git_blob_oid": V0418_OID,
        "role": "ownership_route_disposition_only",
        "semantic_value_authority": False,
        "decision_count": 171,
        "atomic_obligation_count": 580,
        "raw_appendix_sha256": V0418_RAW_SHA256,
        "canonical_appendix_sha256": V0418_CANONICAL_SHA256,
        "r1_acceptance_contract": V0417_PATH,
        "r1_acceptance_contract_git_blob_oid": V0417_OID,
    }
    if precedence != expected_precedence:
        errors.append("ownership precedence binding changed")

    entities = contract.get("entities")
    dispositions = contract.get("decision_dispositions")
    if not isinstance(entities, list) or not isinstance(dispositions, list):
        return errors + ["entities/dispositions must be arrays"]
    entity_ids = [item.get("entity_id") for item in entities if isinstance(item, dict)]
    if len(entity_ids) != len(entities) or len(entity_ids) != len(set(entity_ids)):
        errors.append("entity IDs are missing or duplicated")
    entity_by_id = {
        item["entity_id"]: item
        for item in entities
        if isinstance(item, dict) and isinstance(item.get("entity_id"), str)
    }
    atomic_items = ownership.get("atomic_obligations", [])
    atomic_by_id = {item["obligation_id"]: item for item in atomic_items}
    atomic_ids = [item.get("atomic_obligation_id") for item in entities if isinstance(item, dict)]
    if Counter(atomic_ids) != Counter({key: 1 for key in atomic_by_id}):
        errors.append("atomic obligations are not disposed exactly once")
    decision_ids = [item.get("decision_id") for item in dispositions if isinstance(item, dict)]
    if Counter(decision_ids) != Counter({key: 1 for key in decisions}):
        errors.append("decision rows are not disposed bidirectionally")
    referenced_entities: list[str] = []
    for disposition in dispositions:
        if not isinstance(disposition, dict) or set(disposition) != {
            "decision_id",
            "entity_ids",
        }:
            errors.append("decision disposition shape is invalid")
            continue
        ids = disposition["entity_ids"]
        if not isinstance(ids, list) or not ids or len(ids) != len(set(ids)):
            errors.append(f"{disposition['decision_id']} disposition is empty/duplicated")
        else:
            referenced_entities.extend(ids)
    if Counter(referenced_entities) != Counter({key: 1 for key in entity_ids}):
        errors.append("entities are orphaned or multiply referenced")

    capability_ids = {
        item["entity_id"]
        for item in entities
        if isinstance(item, dict) and item.get("contract_kind") == "capability"
    }
    property_ids = {
        item["entity_id"]
        for item in entities
        if isinstance(item, dict) and item.get("contract_kind") == "property"
    }
    external_contract_ids = {
        item["entity_id"]
        for item in entities
        if isinstance(item, dict)
        and item.get("contract_kind") in {"later_rule", "non_registry_contract"}
    }
    atomic_by_decision: dict[str, list[dict[str, Any]]] = {}
    for atomic in atomic_items:
        atomic_by_decision.setdefault(atomic["source_decision_id"], []).append(atomic)

    blob_cache: dict[str, str] = {}
    observed_field_references: set[tuple[str, str]] = set()
    observed_field_preservations: set[tuple[str, str, str]] = set()
    for entity in entities:
        if not isinstance(entity, dict):
            errors.append("entity is not an object")
            continue
        required_entity = {
            "entity_id",
            "contract_kind",
            "source_decision_id",
            "atomic_obligation_id",
            "ownership",
            "source_refs",
            "definition",
        }
        if set(entity) != required_entity:
            errors.append(f"{entity.get('entity_id')} entity fields are not closed")
            continue
        entity_id = entity["entity_id"]
        kind = entity["contract_kind"]
        if kind not in ALLOWED_KINDS:
            errors.append(f"{entity_id} uses forbidden contract kind")
            continue
        atomic = atomic_by_id.get(entity["atomic_obligation_id"])
        if atomic is None:
            continue
        if entity["source_decision_id"] != atomic["source_decision_id"]:
            errors.append(f"{entity_id} source decision differs from ownership precedence")
        expected_ownership = {
            "primary_owner": atomic["primary_responsible_owner"],
            "responsibility_kind": atomic["responsibility_kind"],
            "operation_key": atomic["operation_key"],
            "route_intersection": atomic["route_intersection"],
            "selected_stage": atomic["selected_stage"],
            "target_stage": atomic["target_stage"],
        }
        if entity["ownership"] != expected_ownership:
            errors.append(f"{entity_id} ownership/route differs from V0.4.1.8")
        refs = entity["source_refs"]
        if not isinstance(refs, list) or len(refs) != 2:
            errors.append(f"{entity_id} must bind atomic and semantic sources")
        else:
            for ref in refs:
                errors.extend(
                    f"{entity_id}: {message}"
                    for message in _validate_source_ref(ref, blob_cache)
                )
            if refs[0] != atomic["atomic_source"]:
                errors.append(f"{entity_id} atomic source differs from ownership precedence")
            if refs[1] != decisions[entity["source_decision_id"]]["source_ref"]:
                errors.append(f"{entity_id} semantic source differs from extractor")
        definition = entity["definition"]
        if not isinstance(definition, dict):
            errors.append(f"{entity_id} definition must be structured")
            continue
        if "semantic_requirements" in definition or "opaque_value" in definition:
            errors.append(f"{entity_id} uses an opaque semantic fallback")
        if _contains_generic_placeholder(definition):
            errors.append(f"{entity_id} retains a generic semantic placeholder")
        serialized_definition = json.dumps(definition, ensure_ascii=False, sort_keys=True)
        for sibling in atomic_by_decision.get(entity["source_decision_id"], []):
            if sibling["obligation_id"] == entity["atomic_obligation_id"]:
                continue
            sibling_boundary = sibling.get("source_mode_or_boundary")
            if (
                isinstance(sibling_boundary, str)
                and len(sibling_boundary) >= 12
                and sibling_boundary in serialized_definition
            ):
                errors.append(
                    f"{entity_id} copies sibling responsibility {sibling['obligation_id']}"
                )
        if kind == "property":
            required = {
                "representation_kind",
                "domain_scope",
                "domain_contract",
                "value_schema",
                "cross_field_invariants",
                "unit_contract",
                "comparison_precision",
                "ranges",
                "mode_layer",
                "qa_missing_behavior",
                "capability_refs",
                "comparison_conclusion",
            }
            if set(definition) != required:
                errors.append(f"{entity_id} property fields are incomplete")
            else:
                errors.extend(_validate_domain(definition, entity_id, entity))
                value_schema = definition["value_schema"]
                if not isinstance(value_schema, dict) or set(value_schema) != {
                    "object_name",
                    "fields",
                    "required_fields",
                    "additional_fields",
                    "variants",
                }:
                    errors.append(f"{entity_id} value_schema is opaque")
                    field_names: set[str] = set()
                else:
                    fields = value_schema["fields"]
                    if not isinstance(fields, list) or not fields:
                        errors.append(f"{entity_id} has no typed property fields")
                        field_names = set()
                    else:
                        field_names = {
                            item.get("name") for item in fields if isinstance(item, dict)
                        }
                        if len(field_names) != len(fields) or None in field_names:
                            errors.append(f"{entity_id} property field names are missing/duplicated")
                        for field in fields:
                            if not isinstance(field, dict) or set(field) != {
                                "name",
                                "type",
                                "required",
                                "domain",
                                "source_component",
                            }:
                                errors.append(f"{entity_id} property field shape is not closed")
                                continue
                            domain_spec = field["domain"]
                            if not isinstance(domain_spec, dict) or domain_spec.get("kind") not in {
                                "const",
                                "enum",
                                "exact_numeric",
                                "source_preserved",
                                "conditional",
                                "reference",
                                "open_bounded",
                            }:
                                errors.append(f"{entity_id}.{field.get('name')} has an untyped domain")
                                continue
                            if field["type"] == "number":
                                if domain_spec["kind"] not in {
                                    "exact_numeric",
                                    "open_bounded",
                                    "source_preserved",
                                } or not domain_spec.get("dimension") or not domain_spec.get(
                                    "allowed_units"
                                ) or domain_spec.get("canonical_unit") not in domain_spec.get(
                                    "allowed_units", []
                                ):
                                    errors.append(
                                        f"{entity_id}.{field['name']} lacks numeric dimension/unit closure"
                                    )
                                required_numeric_keys = {
                                    "exact_numeric": {
                                        "kind",
                                        "value",
                                        "unit",
                                        "dimension",
                                        "canonical_unit",
                                        "allowed_units",
                                        "unit_definition",
                                    },
                                    "open_bounded": {
                                        "kind",
                                        "minimum",
                                        "maximum",
                                        "unit",
                                        "dimension",
                                        "canonical_unit",
                                        "allowed_units",
                                        "unit_definition",
                                        "inclusive_minimum",
                                        "inclusive_maximum",
                                    },
                                    "source_preserved": {
                                        "kind",
                                        "preservation_contract",
                                        "dimension",
                                        "canonical_unit",
                                        "allowed_units",
                                        "unit_definition",
                                        "finite",
                                    },
                                }
                                if set(domain_spec) != required_numeric_keys.get(
                                    domain_spec["kind"], set()
                                ) or (
                                    domain_spec["kind"] == "source_preserved"
                                    and domain_spec.get("finite") is not True
                                ):
                                    errors.append(
                                        f"{entity_id}.{field['name']} numeric domain fields are not closed"
                                    )
                                canonical = domain_spec.get("canonical_unit")
                                unit_definition = UNIT_DEFINITIONS.get(canonical)
                                if (
                                    unit_definition is None
                                    or domain_spec.get("dimension") != unit_definition[0]
                                    or domain_spec.get("unit_definition") != unit_definition[1]
                                    or domain_spec.get("allowed_units") != [canonical]
                                    or (
                                        domain_spec["kind"] in {"exact_numeric", "open_bounded"}
                                        and domain_spec.get("unit") != canonical
                                    )
                                ):
                                    errors.append(
                                        f"{entity_id}.{field['name']} numeric unit definition is unsupported"
                                    )
                            elif domain_spec["kind"] in {"exact_numeric", "open_bounded"}:
                                errors.append(f"{entity_id}.{field['name']} numeric domain has non-number type")
                            if domain_spec["kind"] == "reference":
                                reference_key = (
                                    entity["source_decision_id"],
                                    field["name"],
                                )
                                observed_field_references.add(reference_key)
                                errors.extend(
                                    _validate_field_reference(
                                        domain_spec,
                                        owner_entity=entity,
                                        field_name=field["name"],
                                        entity_by_id=entity_by_id,
                                    )
                                )
                            preservation_errors, preservation_keys = _walk_field_preservations(
                                domain_spec,
                                owner_entity=entity,
                                object_name=value_schema["object_name"],
                                field_name=field["name"],
                                field_type=field["type"],
                            )
                            errors.extend(preservation_errors)
                            observed_field_preservations.update(preservation_keys)
                    required_fields = value_schema["required_fields"]
                    if not isinstance(required_fields, list) or not set(required_fields) <= field_names:
                        errors.append(f"{entity_id} required_fields do not reference typed fields")
                    if value_schema["additional_fields"] is not False:
                        errors.append(f"{entity_id} property schema is not closed")
                invariants = definition["cross_field_invariants"]
                if not isinstance(invariants, list) or not invariants:
                    errors.append(f"{entity_id} invariants are not structured")
                else:
                    for invariant in invariants:
                        if not isinstance(invariant, dict) or set(invariant) != {
                            "invariant_id",
                            "expression",
                            "fields",
                        } or not set(invariant.get("fields", [])) <= field_names:
                            errors.append(f"{entity_id} invariant is not field-bound")
                            continue
                        if invariant.get("invariant_id") == A001_PAPER_INVARIANT[
                            "invariant_id"
                        ]:
                            if entity["atomic_obligation_id"] != "V040-A-001:registry" or invariant != A001_PAPER_INVARIANT:
                                errors.append(f"{entity_id} paper geometry invariant changed")
                            continue
                        expression = invariant["expression"]
                        if not isinstance(expression, dict) or expression.get("operator") != "and":
                            errors.append(f"{entity_id} invariant expression is not typed")
                            continue
                        terms = expression.get("terms")
                        expected_terms = [
                            {
                                "field": field["name"],
                                "comparator": field["domain"]["kind"],
                                "expected": {
                                    key: value
                                    for key, value in field["domain"].items()
                                    if key != "kind"
                                },
                            }
                            for field in value_schema["fields"]
                        ] if isinstance(value_schema, dict) and "fields" in value_schema else []
                        if terms != expected_terms:
                            errors.append(f"{entity_id} invariant differs from typed fields")
                    if entity["atomic_obligation_id"] == "V040-A-001:registry" and A001_PAPER_INVARIANT not in invariants:
                        errors.append(f"{entity_id} paper geometry invariant is missing")
                ranges = definition["ranges"]
                if not isinstance(ranges, list):
                    errors.append(f"{entity_id} ranges are not structured")
                else:
                    for range_item in ranges:
                        if not isinstance(range_item, dict) or range_item.get("field") not in field_names:
                            errors.append(f"{entity_id} range is not field-bound")
                precision = definition["comparison_precision"]
                if not isinstance(precision, dict) or set(precision) != {
                    "field_precisions",
                    "exact_fields",
                }:
                    errors.append(f"{entity_id} comparison precision is not field-bound")
                else:
                    precision_fields = {
                        item.get("field")
                        for item in precision["field_precisions"]
                        if isinstance(item, dict)
                    } | set(precision["exact_fields"])
                    if not precision_fields <= field_names:
                        errors.append(f"{entity_id} precision names unknown fields")
                if isinstance(value_schema, dict) and "fields" in value_schema:
                    numeric_fields = {
                        field["name"]: field["domain"]
                        for field in value_schema["fields"]
                        if field["type"] == "number"
                    }
                    expected_units = [
                        {
                            "field": name,
                            "dimension": domain.get("dimension"),
                            "canonical_unit": domain.get("canonical_unit"),
                            "allowed_units": domain.get("allowed_units"),
                            "unit_definition": domain.get("unit_definition"),
                        }
                        for name, domain in numeric_fields.items()
                    ]
                    unit_contract = definition["unit_contract"]
                    if not isinstance(unit_contract, dict) or set(unit_contract) != {
                        "field_units",
                        "unitless_fields",
                    } or unit_contract.get("field_units") != expected_units or set(
                        unit_contract.get("unitless_fields", [])
                    ) != field_names - set(numeric_fields):
                        errors.append(f"{entity_id} unit contract differs from typed fields")
                    expected_ranges = []
                    for name, domain in numeric_fields.items():
                        if domain["kind"] == "exact_numeric":
                            expected_ranges.append(
                                {
                                    "field": name,
                                    "minimum": domain["value"],
                                    "maximum": domain["value"],
                                    "inclusive_minimum": True,
                                    "inclusive_maximum": True,
                                    "unit": domain["unit"],
                                }
                            )
                        elif domain["kind"] == "open_bounded":
                            expected_ranges.append(
                                {
                                    "field": name,
                                    "minimum": domain["minimum"],
                                    "maximum": domain["maximum"],
                                    "inclusive_minimum": domain["inclusive_minimum"],
                                    "inclusive_maximum": domain["inclusive_maximum"],
                                    "unit": domain["unit"],
                                }
                            )
                    if ranges != expected_ranges:
                        errors.append(f"{entity_id} ranges differ from typed fields")
                    expected_precision = {
                        "field_precisions": [
                            {
                                "field": name,
                                "absolute_tolerance": 0,
                                "unit": domain.get("canonical_unit"),
                            }
                            for name, domain in numeric_fields.items()
                        ],
                        "exact_fields": [
                            field["name"]
                            for field in value_schema["fields"]
                            if field["name"] not in numeric_fields
                        ],
                    }
                    if precision != expected_precision:
                        errors.append(f"{entity_id} precision differs from typed fields")
                    if definition["domain_scope"] == "source_preserved":
                        source_units = definition["domain_contract"].get(
                            "unit_and_finite_constraints"
                        )
                        preserved_unit_map = {
                            item["field"]: item
                            for item in source_units.get("field_units", [])
                        } if isinstance(source_units, dict) else {}
                        for name, domain in numeric_fields.items():
                            expected_unit = {
                                "field": name,
                                "dimension": domain.get("dimension"),
                                "canonical_unit": domain.get("canonical_unit"),
                                "allowed_units": domain.get("allowed_units"),
                                "unit_definition": domain.get("unit_definition"),
                            }
                            if preserved_unit_map.get(name) != expected_unit:
                                errors.append(
                                    f"{entity_id}.{name} differs from source-preserved unit closure"
                                )
                errors.extend(
                    _validate_reference_list(
                        definition["capability_refs"],
                        allowed_targets=capability_ids,
                        context=f"{entity_id}.capability_refs",
                        field_name="capability_refs",
                        owner_entity=entity,
                    )
                )
                if entity["atomic_obligation_id"] == "V040-A-001:registry" and isinstance(
                    definition.get("value_schema"), dict
                ) and isinstance(definition.get("domain_contract"), dict):
                    conclusion = definition["comparison_conclusion"]
                    expected_paper_conclusion = {
                        "authority_literal": decisions["V040-A-001"]["semantic_text"],
                        "selected_scope": "source_preserved",
                        "rejected_scopes": ["project_closed", "project_open_constrained"],
                        "reason": "section_paper_geometry preserves the frozen source width, height, orientation, and section boundary without a paper-size enum",
                        "unit": "twentieth_of_a_point",
                        "unit_definition": "one twentieth of one typographic point (twip)",
                        "valid_source_pair": {
                            "width": {"finite": True, "greater_than": 0},
                            "height": {"finite": True, "greater_than": 0},
                        },
                        "custom_paper_policy": "preserve_exact_source_pair",
                        "standard_paper_normalization": False,
                        "swap_width_height": False,
                        "supply_missing_default": False,
                        "reject_only_for_nonstandard_size": False,
                        "invalid_missing_nonpositive_nonfinite_or_unknown_unit": "fail_closed",
                        "orientation_policy": "preserve_independently",
                        "section_field_mapping": {
                            "value_schema_field": "section_boundary",
                            "canonical_shape_field": "sectPr_boundary",
                            "relation": "one_to_one",
                        },
                    }
                    if conclusion != expected_paper_conclusion:
                        errors.append(f"{entity_id} paper comparison conclusion changed")
                    paper_fields = {
                        field["name"]: field
                        for field in definition["value_schema"]["fields"]
                    }
                    for field_name in ("paper_width", "paper_height"):
                        paper_domain = paper_fields[field_name]["domain"]
                        if (
                            paper_domain.get("canonical_unit")
                            != "twentieth_of_a_point"
                            or paper_domain.get("allowed_units")
                            != ["twentieth_of_a_point"]
                            or paper_domain.get("dimension") != "length"
                            or paper_domain.get("unit_definition")
                            != "one twentieth of one typographic point (twip)"
                            or paper_domain.get("finite") is not True
                        ):
                            errors.append(f"{entity_id}.{field_name} paper unit changed")
                    paper_shape = definition["domain_contract"][
                        "canonical_preservation_shape"
                    ]
                    if paper_shape.get("identity_fields") != [
                        "section_id",
                        "sectPr_boundary",
                    ] or conclusion["section_field_mapping"] != {
                        "value_schema_field": "section_boundary",
                        "canonical_shape_field": "sectPr_boundary",
                        "relation": "one_to_one",
                    }:
                        errors.append(f"{entity_id} section boundary mapping changed")
        elif kind == "constraint":
            required = {
                "domain_scope",
                "domain_contract",
                "dependency_classes",
                "registry_property_ids",
                "external_obligations",
                "capability_dependencies",
                "decidable_invariants",
                "validator_owner",
            }
            if set(definition) != required:
                errors.append(f"{entity_id} constraint fields are incomplete")
            else:
                errors.extend(_validate_domain(definition, entity_id, entity))
                if not isinstance(definition["dependency_classes"], list) or not definition["dependency_classes"]:
                    errors.append(f"{entity_id} has no dependency classes")
                errors.extend(
                    _validate_reference_list(
                        definition["registry_property_ids"],
                        allowed_targets=property_ids,
                        context=f"{entity_id}.registry_property_ids",
                        field_name="registry_property_ids",
                        owner_entity=entity,
                    )
                )
                errors.extend(
                    _validate_reference_list(
                        definition["capability_dependencies"],
                        allowed_targets=capability_ids,
                        context=f"{entity_id}.capability_dependencies",
                        field_name="capability_dependencies",
                        owner_entity=entity,
                    )
                )
                errors.extend(
                    _validate_reference_list(
                        definition["external_obligations"],
                        allowed_targets=external_contract_ids,
                        context=f"{entity_id}.external_obligations",
                        field_name="external_obligations",
                        owner_entity=entity,
                    )
                )
                invariants = definition["decidable_invariants"]
                if not isinstance(invariants, list) or not invariants:
                    errors.append(f"{entity_id} constraint has no decidable invariant")
                else:
                    for invariant in invariants:
                        if not isinstance(invariant, dict) or set(invariant) != {
                            "invariant_id",
                            "subject",
                            "predicate",
                            "expected",
                            "source_scope",
                        } or invariant.get("source_scope") != entity["atomic_obligation_id"]:
                            errors.append(f"{entity_id} invariant is not atomic-local")
                            continue
                        expected = invariant["expected"]
                        if entity["ownership"]["responsibility_kind"] == "preservation":
                            if (
                                invariant["predicate"] != "round_trip_object_equivalent"
                                or not isinstance(expected, dict)
                                or expected.get("object_kind")
                                != definition["domain_contract"].get("source_type")
                            ):
                                errors.append(f"{entity_id} preservation invariant is not object-specific")
                        elif entity["atomic_obligation_id"].endswith(":registry"):
                            if (
                                invariant["predicate"]
                                != "required_and_forbidden_action_sets_match"
                                or not isinstance(expected, dict)
                                or not expected.get("required_actions")
                                or not expected.get("forbidden_actions")
                                or expected.get("unknown_action_policy") != "reject"
                            ):
                                errors.append(f"{entity_id} registry invariant is not action-closed")
        elif kind == "capability":
            required = {
                "executor_or_auditor_class",
                "availability",
                "owner",
                "input_boundary",
                "output_boundary",
            }
            if set(definition) != required:
                errors.append(f"{entity_id} capability fields are incomplete")
            elif definition["availability"] == "implemented":
                errors.append(f"{entity_id} invents implemented capability")
        else:
            required = {
                "owner",
                "target_stage",
                "contract_locator",
                "input_boundary",
                "output_boundary",
                "preserved_original_obligation",
            }
            if set(definition) != required:
                errors.append(f"{entity_id} later/non-registry fields are incomplete")
    if observed_field_references != set(FIELD_REFERENCE_EXPECTATIONS):
        errors.append("typed field-reference graph is incomplete or contains extras")
    if observed_field_preservations != set(FIELD_PRESERVATION_SELECTORS):
        errors.append("field-level preservation graph is incomplete or contains extras")
    if contract.get("unresolved_obligations") != []:
        errors.append("unresolved obligations must be zero")
    summary = contract.get("summary")
    if not isinstance(summary, dict):
        errors.append("summary must be an object")
    else:
        expected_kind_counts = dict(sorted(Counter(item["contract_kind"] for item in entities).items()))
        expected_domain_counts = dict(
            sorted(
                Counter(
                    item["definition"]["domain_scope"]
                    for item in entities
                    if item["contract_kind"] in {"property", "constraint"}
                    and isinstance(item.get("definition"), dict)
                    and "domain_scope" in item["definition"]
                ).items()
            )
        )
        expected_summary = {
            "decision_rows": 171,
            "atomic_obligations": 580,
            "entities": len(entities),
            "unresolved": 0,
            "kind_counts": expected_kind_counts,
            "domain_scope_counts": expected_domain_counts,
        }
        if summary != expected_summary:
            errors.append("summary differs from contract contents")
    non_authorizations = contract.get("non_authorizations")
    required_stops = {"R2", "C2", "P3b", "public_CLI", "Ready", "release", "push", "PR", "merge"}
    if not isinstance(non_authorizations, list) or not required_stops.issubset(non_authorizations):
        errors.append("non-authorization boundary is incomplete")
    return errors


def _render_report(contract: dict[str, Any], raw_bytes: bytes) -> bytes:
    canonical_digest = _sha256(_canonical_bytes(contract))
    raw_digest = _sha256(raw_bytes)
    summary = contract["summary"]
    lines = [
        "# V0.4.1.7 R1 Semantic-Domain Contract Report",
        "",
        "Status: test-only candidate; not an R1 acceptance and not an R2 authorization.",
        "",
        f"- Base commit: `{contract['base_commit']}`",
        f"- Raw contract SHA-256: `{raw_digest}`",
        f"- Canonical contract SHA-256: `{canonical_digest}`",
        f"- Decision rows: {summary['decision_rows']}",
        f"- Atomic obligations: {summary['atomic_obligations']}",
        f"- Entities: {summary['entities']}",
        f"- Unresolved: {summary['unresolved']}",
        "",
        "## Authority identities",
        "",
        "| Role | Path | Git blob OID |",
        "| --- | --- | --- |",
    ]
    for item in contract["authority_inputs"]:
        lines.append(
            f"| {item['permitted_contribution']} | `{item['path']}` | `{item['git_blob_oid']}` |"
        )
    precedence = contract["ownership_precedence"]
    lines.append(
        f"| ownership_route_disposition_only | `{precedence['path']}` | `{precedence['git_blob_oid']}` |"
    )
    lines.extend(["", "## Classification counts", ""])
    for key, value in sorted(summary["kind_counts"].items()):
        lines.append(f"- `{key}`: {value}")
    lines.extend(["", "## Domain-scope counts", ""])
    for key, value in sorted(summary["domain_scope_counts"].items()):
        lines.append(f"- `{key}`: {value}")
    semantic_level_fields = [
        (entity["source_decision_id"], field["name"])
        for entity in contract["entities"]
        if entity["contract_kind"] == "property"
        for field in entity["definition"].get("value_schema", {}).get("fields", [])
        if isinstance(field.get("domain"), dict)
        and field["domain"].get("value_contract", {}).get("kind")
        == "approved_semantic_heading_level_mapping"
    ]
    lines.extend(
        [
            "",
            "## Semantic heading-level closure",
            "",
            f"- Typed semantic-level fields: {len(semantic_level_fields)}",
            "- Allowed levels: `1`, `2`, `3`, `4`",
            "- Value-style authorities: `V040-E-002`, `V040-E-003`, `V040-E-004`",
            "- Selection source: frozen semantic artifact or user/publisher-approved artifact",
            "- `execution-p5b`: consumer/gate only; never a value source",
            "",
            "## N007 border-value closure",
            "",
            "- Immediate value targets: `V040-N-003`, `V040-N-004`, `V040-N-006` properties only",
            "- `V040-N-005`: exact deferred non-value branch, blocked to `V0.4.2`",
            "- Unknown table semantic: reject; ambiguous classification: blocked QA",
        ]
    )
    lines.extend(
        [
            "",
            "## Boundaries",
            "",
            "V0.4.0 is the sole semantic/value authority. V0.4.1.8 contributes only ownership, route, and atomic disposition precedence. The report does not accept R1 and does not authorize R2, C2, P3b, a public CLI, Ready, release, push, PR, or merge.",
            "",
        ]
    )
    return "\n".join(lines).encode("utf-8")


class V0417R1DomainContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.contract_bytes = CONTRACT_PATH.read_bytes()
        cls.report_bytes = REPORT_PATH.read_bytes()
        cls.contract = json.loads(cls.contract_bytes)
        cls.decisions = _extract_decisions()
        cls.ownership = _extract_ownership_precedence()

    def test_001_exact_main_staleness_and_three_path_boundary(self) -> None:
        mode, errors = _runtime_checkpoint_evidence()
        self.assertIn(
            mode,
            {
                "local",
                "local_integrated",
                "github_pull_request",
                "github_push_integrated",
            },
        )
        self.assertEqual([], errors)

    def test_001b_precommit_and_candidate_checkpoint_lifecycle_matrix(self) -> None:
        pre = dict(
            head=BASE_COMMIT,
            main=BASE_COMMIT,
            origin_main=BASE_COMMIT,
            status_lines=[f"?? {path}" for path in sorted(ALLOWED_PATHS)],
            parent_oids=[],
            changed_paths=set(),
        )
        candidate = dict(
            head="1" * 40,
            main=BASE_COMMIT,
            origin_main=BASE_COMMIT,
            status_lines=[],
            parent_oids=[BASE_COMMIT],
            changed_paths=set(ALLOWED_PATHS),
        )
        self.assertEqual([], _validate_checkpoint_state(**pre))
        self.assertEqual([], _validate_checkpoint_state(**candidate))
        mutations = []
        for base in (pre, candidate):
            drift = copy.deepcopy(base)
            drift["main"] = "2" * 40
            mutations.append(drift)
        extra_untracked = copy.deepcopy(pre)
        extra_untracked["status_lines"].append("?? extra")
        mutations.append(extra_untracked)
        dirty_candidate = copy.deepcopy(candidate)
        dirty_candidate["status_lines"] = [" M tracked"]
        mutations.append(dirty_candidate)
        chain = copy.deepcopy(candidate)
        chain["parent_oids"] = ["3" * 40]
        mutations.append(chain)
        extra_path = copy.deepcopy(candidate)
        extra_path["changed_paths"].add("extra")
        mutations.append(extra_path)
        for mutation in mutations:
            self.assertTrue(_validate_checkpoint_state(**mutation))

    def test_001c_github_pull_request_evidence_fail_closed_matrix(self) -> None:
        self.assertIsNone(_load_github_event({}))
        with self.assertRaises(AssertionError):
            _load_github_event({"GITHUB_ACTIONS": "true"})
        with self.assertRaises(AssertionError):
            _load_github_event(
                {
                    "GITHUB_ACTIONS": "true",
                    "GITHUB_EVENT_NAME": "schedule",
                    "GITHUB_EVENT_PATH": "/does/not/matter",
                    "GITHUB_SHA": "1" * 40,
                    "GITHUB_REPOSITORY": CANONICAL_REPOSITORY,
                    "GITHUB_REF": "refs/heads/main",
                }
            )
        with self.assertRaises(AssertionError):
            _load_github_event(
                {
                    "GITHUB_ACTIONS": "true",
                    "GITHUB_EVENT_NAME": "pull_request",
                    "GITHUB_EVENT_PATH": str(TEST_PATH),
                    "GITHUB_SHA": "1" * 40,
                    "GITHUB_REPOSITORY": CANONICAL_REPOSITORY,
                    "GITHUB_REF": "refs/pull/1/merge",
                }
            )
        with tempfile.TemporaryDirectory() as tmp:
            event_path = Path(tmp) / "event.json"
            event_path.write_text("not-json", encoding="utf-8")
            event_environ = {
                "GITHUB_ACTIONS": "true",
                "GITHUB_EVENT_NAME": "pull_request",
                "GITHUB_EVENT_PATH": str(event_path),
                "GITHUB_SHA": "1" * 40,
                "GITHUB_REPOSITORY": CANONICAL_REPOSITORY,
                "GITHUB_REF": "refs/pull/1/merge",
            }
            with self.assertRaises(AssertionError):
                _load_github_event(event_environ)
            parsed_payload = {
                "repository": {"full_name": CANONICAL_REPOSITORY},
                "pull_request": {},
            }
            event_path.write_text(json.dumps(parsed_payload), encoding="utf-8")
            self.assertEqual(
                {
                    "payload": parsed_payload,
                    "github_event_name": "pull_request",
                    "github_sha": "1" * 40,
                    "github_repository": CANONICAL_REPOSITORY,
                    "github_ref": "refs/pull/1/merge",
                },
                _load_github_event(event_environ),
            )
            noncanonical_environ = {
                **event_environ,
                "GITHUB_REPOSITORY": "owner/repository",
            }
            with self.assertRaises(AssertionError):
                _load_github_event(noncanonical_environ)

        head_sha = "1" * 40
        merge_sha = "2" * 40
        repository = CANONICAL_REPOSITORY
        payload = {
            "repository": {"full_name": repository},
            "pull_request": {
                "base": {"sha": BASE_COMMIT},
                "head": {"sha": head_sha},
                "merge_commit_sha": merge_sha,
                "changed_files": 3,
                "commits": 1,
            },
        }
        checkout_blobs = {
            **FIXED_CHECKOUT_BLOB_OIDS,
            "tests/test_v0417_p3ar_r1_domain_contract.py": "3" * 40,
        }

        def evidence(
            *,
            github_sha: str,
            payload_value: Any | None = None,
        ) -> dict[str, Any]:
            return {
                "payload": copy.deepcopy(payload if payload_value is None else payload_value),
                "github_event_name": "pull_request",
                "github_sha": github_sha,
                "github_repository": repository,
                "github_ref": "refs/pull/1/merge",
            }

        direct = dict(
            event=evidence(github_sha=head_sha),
            head=head_sha,
            parent_oids=[BASE_COMMIT],
            commit_message=b"candidate\n",
            status_lines=[],
            checkout_blob_oids=copy.deepcopy(checkout_blobs),
            working_blob_oids=copy.deepcopy(checkout_blobs),
        )
        synthetic = copy.deepcopy(direct)
        synthetic.update(
            event=evidence(github_sha=merge_sha),
            head=merge_sha,
            parent_oids=[BASE_COMMIT, head_sha],
            commit_message=f"Merge {head_sha} into {BASE_COMMIT}\n".encode("ascii"),
        )
        self.assertEqual([], _validate_github_pull_request_checkout(**direct))
        direct_default_event_sha = copy.deepcopy(direct)
        direct_default_event_sha["event"]["github_sha"] = merge_sha
        self.assertEqual(
            [],
            _validate_github_pull_request_checkout(**direct_default_event_sha),
        )
        self.assertEqual([], _validate_github_pull_request_checkout(**synthetic))

        current_head_sha = "661df8af3d2705f8d3494de1e1cd2276f56fe6d6"
        current_merge_sha = "38df0d02da3cd95883f2232ff9f764f32cd448f0"
        current_payload = copy.deepcopy(payload)
        current_payload["pull_request"]["head"]["sha"] = current_head_sha
        current_payload["pull_request"]["merge_commit_sha"] = current_head_sha
        current_synthetic = copy.deepcopy(synthetic)
        current_synthetic.update(
            event=evidence(
                github_sha=current_merge_sha,
                payload_value=current_payload,
            ),
            head=current_merge_sha,
            parent_oids=[BASE_COMMIT, current_head_sha],
            commit_message=(
                f"Merge {current_head_sha} into {BASE_COMMIT}\n".encode("ascii")
            ),
        )
        self.assertEqual(
            [],
            _validate_github_pull_request_checkout(**current_synthetic),
        )
        null_merge_sha = copy.deepcopy(current_synthetic)
        null_merge_sha["event"]["payload"]["pull_request"]["merge_commit_sha"] = None
        self.assertEqual(
            [],
            _validate_github_pull_request_checkout(**null_merge_sha),
        )

        mutations: list[dict[str, Any]] = []
        missing_event = copy.deepcopy(direct)
        missing_event["event"] = {"payload": {}}
        mutations.append(missing_event)
        wrong_base = copy.deepcopy(direct)
        wrong_base["event"]["payload"]["pull_request"]["base"]["sha"] = "4" * 40
        mutations.append(wrong_base)
        wrong_head = copy.deepcopy(direct)
        wrong_head["event"]["payload"]["pull_request"]["head"]["sha"] = "4" * 40
        mutations.append(wrong_head)
        wrong_runner_sha = copy.deepcopy(direct)
        wrong_runner_sha["event"]["github_sha"] = "4" * 40
        mutations.append(wrong_runner_sha)
        wrong_parents = copy.deepcopy(synthetic)
        wrong_parents["parent_oids"] = [head_sha, BASE_COMMIT]
        mutations.append(wrong_parents)
        wrong_merge_message = copy.deepcopy(synthetic)
        wrong_merge_message["commit_message"] = b"Merge unexpected into evidence\n"
        mutations.append(wrong_merge_message)
        stale_merge_sha = copy.deepcopy(synthetic)
        stale_merge_sha["event"]["payload"]["pull_request"][
            "merge_commit_sha"
        ] = "f" * 40
        self.assertEqual(
            [],
            _validate_github_pull_request_checkout(**stale_merge_sha),
        )
        for invalid_merge_sha in ("F" * 40, "f" * 39, "g" * 40, 1):
            invalid_merge = copy.deepcopy(synthetic)
            invalid_merge["event"]["payload"]["pull_request"][
                "merge_commit_sha"
            ] = invalid_merge_sha
            mutations.append(invalid_merge)
        dirty = copy.deepcopy(direct)
        dirty["status_lines"] = [" M tests/test_v0417_p3ar_r1_domain_contract.py"]
        mutations.append(dirty)
        missing_path = copy.deepcopy(direct)
        missing_path["checkout_blob_oids"].pop(
            "tests/fixtures/v0417/p3a_r/r1-domain-report.md"
        )
        mutations.append(missing_path)
        extra_path = copy.deepcopy(direct)
        extra_path["checkout_blob_oids"]["unexpected"] = "5" * 40
        extra_path["working_blob_oids"]["unexpected"] = "5" * 40
        mutations.append(extra_path)
        blob_mismatch = copy.deepcopy(direct)
        blob_mismatch["working_blob_oids"][
            "tests/test_v0417_p3ar_r1_domain_contract.py"
        ] = "6" * 40
        mutations.append(blob_mismatch)
        frozen_blob_mismatch = copy.deepcopy(direct)
        frozen_blob_mismatch["checkout_blob_oids"][
            "tests/fixtures/v0417/p3a_r/r1-domain-contract.json"
        ] = "7" * 40
        frozen_blob_mismatch["working_blob_oids"][
            "tests/fixtures/v0417/p3a_r/r1-domain-contract.json"
        ] = "7" * 40
        mutations.append(frozen_blob_mismatch)
        changed_count = copy.deepcopy(direct)
        changed_count["event"]["payload"]["pull_request"]["changed_files"] = 4
        mutations.append(changed_count)
        missing_commit_count = copy.deepcopy(direct)
        missing_commit_count["event"]["payload"]["pull_request"].pop("commits")
        mutations.append(missing_commit_count)
        for invalid_count in (0, 2, True):
            wrong_commit_count = copy.deepcopy(direct)
            wrong_commit_count["event"]["payload"]["pull_request"]["commits"] = (
                invalid_count
            )
            mutations.append(wrong_commit_count)
        repository_conflict = copy.deepcopy(direct)
        repository_conflict["event"]["payload"]["repository"]["full_name"] = (
            "other/repository"
        )
        mutations.append(repository_conflict)
        noncanonical_repository = copy.deepcopy(direct)
        noncanonical_repository["event"]["github_repository"] = "other/repository"
        noncanonical_repository["event"]["payload"]["repository"]["full_name"] = (
            "other/repository"
        )
        mutations.append(noncanonical_repository)
        unknown_shape = copy.deepcopy(direct)
        unknown_shape["head"] = "8" * 40
        unknown_shape["event"]["github_sha"] = "8" * 40
        unknown_shape["parent_oids"] = [BASE_COMMIT]
        mutations.append(unknown_shape)
        for mutation in mutations:
            self.assertTrue(_validate_github_pull_request_checkout(**mutation))

    def test_001d_raw_commit_parent_headers_are_closed_and_ordered(self) -> None:
        first = "1" * 40
        second = "2" * 40

        def raw_commit(parents: list[str]) -> bytes:
            headers = [b"tree " + b"a" * 40]
            headers.extend(f"parent {parent}".encode("ascii") for parent in parents)
            headers.extend(
                [
                    b"author Test <test@example.invalid> 1 +0000",
                    b"committer Test <test@example.invalid> 1 +0000",
                ]
            )
            return b"\n".join(headers) + b"\n\nmessage\n"

        self.assertEqual([first], _parse_commit_parent_oids(raw_commit([first])))
        self.assertEqual(
            [first, second],
            _parse_commit_parent_oids(raw_commit([first, second])),
        )
        malformed = [
            raw_commit([first]).replace(
                f"parent {first}".encode("ascii"), b"parent not-an-oid"
            ),
            raw_commit([first, first]),
            raw_commit([first, second, "3" * 40]),
            (
                b"tree "
                + b"a" * 40
                + b"\nauthor Test <test@example.invalid> 1 +0000\nparent "
                + first.encode("ascii")
                + b"\ncommitter Test <test@example.invalid> 1 +0000\n\nmessage\n"
            ),
            raw_commit([first]).replace(
                b"committer Test", b"untrusted-header value\ncommitter Test"
            ),
            raw_commit([first]).replace(
                b"committer Test",
                b"author Other <other@example.invalid> 2 +0000\ncommitter Test",
            ),
            raw_commit([first]).replace(b"\n\nmessage\n", b"\nmessage\n"),
            raw_commit([first])[:-1],
        ]
        for raw in malformed:
            with self.subTest(raw=raw[:100]):
                with self.assertRaises(AssertionError):
                    _parse_commit_parent_oids(raw)

    def test_001e_real_depth_one_direct_and_synthetic_runtime_evidence(self) -> None:
        git_environment = os.environ.copy()
        git_environment.pop("GIT_INDEX_FILE", None)

        def git(repo: Path, *args: str) -> str:
            return subprocess.check_output(
                ["git", "-C", str(repo), *args],
                text=True,
                stderr=subprocess.STDOUT,
                env=git_environment,
            ).strip()

        def git_call(repo: Path, *args: str) -> None:
            subprocess.check_call(
                ["git", "-C", str(repo), *args],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                env=git_environment,
            )

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            seed = root / "seed"
            seed.mkdir()
            git_call(seed, "init", "-q")
            git_call(seed, "config", "user.name", "Depth One Test")
            git_call(seed, "config", "user.email", "depth-one@example.invalid")
            (seed / "README").write_text("base\n", encoding="utf-8")
            git_call(seed, "add", "README")
            git_call(seed, "commit", "-q", "-m", "base")
            base_oid = git(seed, "rev-parse", "HEAD")

            git_call(seed, "checkout", "-q", "-b", "candidate")
            source_by_path = {
                "tests/fixtures/v0417/p3a_r/r1-domain-contract.json": CONTRACT_PATH,
                "tests/fixtures/v0417/p3a_r/r1-domain-report.md": REPORT_PATH,
                "tests/test_v0417_p3ar_r1_domain_contract.py": TEST_PATH,
            }
            for relative, source in source_by_path.items():
                destination = seed / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(source.read_bytes())
            git_call(seed, "add", *sorted(ALLOWED_PATHS))
            git_call(seed, "commit", "-q", "-m", "candidate")
            candidate_oid = git(seed, "rev-parse", "HEAD")

            git_call(seed, "checkout", "-q", "--detach", base_oid)
            github_merge_message = f"Merge {candidate_oid} into {base_oid}"
            git_call(
                seed,
                "merge",
                "-q",
                "--no-ff",
                candidate_oid,
                "-m",
                github_merge_message,
            )
            merge_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "branch", "synthetic", merge_oid)
            candidate_tree = git(seed, "rev-parse", f"{candidate_oid}^{{tree}}")
            wrong_order_oid = git(
                seed,
                "commit-tree",
                candidate_tree,
                "-p",
                candidate_oid,
                "-p",
                base_oid,
                "-m",
                "wrong order",
            )
            git_call(seed, "branch", "wrong-order", wrong_order_oid)
            unknown_oid = git(
                seed,
                "commit-tree",
                candidate_tree,
                "-p",
                candidate_oid,
                "-m",
                "unknown shape",
            )
            git_call(seed, "branch", "unknown", unknown_oid)

            repository = "owner/repository"

            def clone(branch: str, name: str) -> Path:
                destination = root / name
                subprocess.check_call(
                    [
                        "git",
                        "clone",
                        "-q",
                        "--depth=1",
                        "--branch",
                        branch,
                        seed.resolve().as_uri(),
                        str(destination),
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    env=git_environment,
                )
                self.assertEqual("true", git(destination, "rev-parse", "--is-shallow-repository"))
                self.assertEqual("", git(destination, "show", "-s", "--format=%P", "HEAD"))
                self.assertIn(
                    git(destination, "rev-parse", "HEAD"),
                    (destination / ".git/shallow").read_text(encoding="ascii").splitlines(),
                )
                return destination

            def event_environ(
                checkout_oid: str,
                event_merge_oid: str | None,
                name: str,
            ) -> dict[str, str]:
                payload = {
                    "repository": {"full_name": repository},
                    "pull_request": {
                        "base": {"sha": base_oid},
                        "head": {"sha": candidate_oid},
                        "merge_commit_sha": event_merge_oid,
                        "changed_files": 3,
                        "commits": 1,
                    },
                }
                event_path = root / f"{name}-event.json"
                event_path.write_text(json.dumps(payload), encoding="utf-8")
                return {
                    "GITHUB_ACTIONS": "true",
                    "GITHUB_EVENT_NAME": "pull_request",
                    "GITHUB_EVENT_PATH": str(event_path),
                    "GITHUB_SHA": checkout_oid,
                    "GITHUB_REPOSITORY": repository,
                    "GITHUB_REF": "refs/pull/1/merge",
                }

            cases = [
                ("candidate", "direct", candidate_oid, merge_oid, [base_oid], None),
                (
                    "synthetic",
                    "synthetic",
                    merge_oid,
                    merge_oid,
                    [base_oid, candidate_oid],
                    None,
                ),
                (
                    "synthetic",
                    "synthetic-event-head-sha",
                    merge_oid,
                    candidate_oid,
                    [base_oid, candidate_oid],
                    None,
                ),
                (
                    "synthetic",
                    "synthetic-null-merge-sha",
                    merge_oid,
                    None,
                    [base_oid, candidate_oid],
                    None,
                ),
                (
                    "synthetic",
                    "synthetic-stale-merge-sha",
                    merge_oid,
                    "f" * 40,
                    [base_oid, candidate_oid],
                    None,
                ),
                (
                    "wrong-order",
                    "wrong-order",
                    wrong_order_oid,
                    wrong_order_oid,
                    [candidate_oid, base_oid],
                    "synthetic pull-request merge parents conflict",
                ),
                (
                    "unknown",
                    "unknown",
                    unknown_oid,
                    merge_oid,
                    [candidate_oid],
                    "checkout shape is unknown",
                ),
            ]
            for (
                branch,
                name,
                checkout_oid,
                event_merge_oid,
                parents,
                error_fragment,
            ) in cases:
                with self.subTest(name=name):
                    checkout = clone(branch, name)
                    environ = event_environ(checkout_oid, event_merge_oid, name)
                    with mock.patch.dict(
                        globals(),
                        {
                            "REPO": checkout,
                            "BASE_COMMIT": base_oid,
                            "CANONICAL_REPOSITORY": repository,
                        },
                    ), mock.patch.dict(os.environ, git_environment, clear=True):
                        observed_parents, observed_message = _commit_object_evidence(
                            "HEAD"
                        )
                        self.assertEqual(parents, observed_parents)
                        if branch == "synthetic":
                            self.assertEqual(
                                f"{github_merge_message}\n".encode("ascii"),
                                observed_message,
                            )
                        mode, errors = _runtime_checkpoint_evidence(environ)
                    self.assertEqual("github_pull_request", mode)
                    if error_fragment is None:
                        self.assertEqual([], errors)
                    else:
                        self.assertTrue(
                            any(error_fragment in error for error in errors),
                            errors,
                        )

    def test_001f_github_push_integrated_fail_closed_matrix(self) -> None:
        before = "1" * 40
        after = "2" * 40
        second_parent = "3" * 40
        repository = CANONICAL_REPOSITORY
        checkout_blobs = {
            **FIXED_CHECKOUT_BLOB_OIDS,
            "tests/test_v0417_p3ar_r1_domain_contract.py": "4" * 40,
        }
        payload = {
            "before": before,
            "after": after,
            "ref": "refs/heads/main",
            "created": False,
            "deleted": False,
            "forced": False,
            "head_commit": {"id": after},
            "commits": [{"id": after}],
            "repository": {"full_name": repository},
        }
        event = {
            "payload": payload,
            "github_event_name": "push",
            "github_sha": after,
            "github_repository": repository,
            "github_ref": "refs/heads/main",
        }
        linear = {
            "event": event,
            "head": after,
            "parent_oids": [before],
            "status_lines": [],
            "checkout_blob_oids": checkout_blobs,
            "working_blob_oids": copy.deepcopy(checkout_blobs),
            "authority_blob_oids": copy.deepcopy(FROZEN_AUTHORITY_BLOB_OIDS),
        }
        merge = copy.deepcopy(linear)
        merge["parent_oids"] = [before, second_parent]
        self.assertEqual([], _validate_github_push_checkout(**linear))
        self.assertEqual([], _validate_github_push_checkout(**merge))
        for allowed_ref in (
            "refs/heads/main",
            "refs/heads/skill/topic",
            "refs/heads/adapter/topic/subtopic",
            "refs/heads/fix/topic-1",
        ):
            mutation = copy.deepcopy(linear)
            mutation["event"]["payload"]["ref"] = allowed_ref
            mutation["event"]["github_ref"] = allowed_ref
            self.assertEqual([], _validate_github_push_checkout(**mutation))

        mutations: list[dict[str, Any]] = []

        def mutate(path: tuple[str, ...], value: Any) -> None:
            item = copy.deepcopy(linear)
            target: Any = item
            for key in path[:-1]:
                target = target[key]
            target[path[-1]] = value
            mutations.append(item)

        mutate(("event", "payload", "before"), "5" * 40)
        reverse = copy.deepcopy(merge)
        reverse["parent_oids"] = [second_parent, before]
        mutations.append(reverse)
        for flag in ("created", "deleted", "forced"):
            mutate(("event", "payload", flag), True)
            mutate(("event", "payload", flag), 0)
        for invalid_ref in (
            "refs/tags/v1",
            "refs/heads/feature/topic",
            "refs/heads/skill/",
            "refs/heads/fix/topic..bad",
            "refs/heads/skill/.hidden",
            "refs/heads/adapter/topic.",
            "refs/heads/fix/topic/.nested",
            "refs/heads/fix/topic/nested.",
        ):
            item = copy.deepcopy(linear)
            item["event"]["payload"]["ref"] = invalid_ref
            item["event"]["github_ref"] = invalid_ref
            mutations.append(item)
        mutate(("event", "github_repository"), "other/repository")
        noncanonical_repository = copy.deepcopy(linear)
        noncanonical_repository["event"]["github_repository"] = "other/repository"
        noncanonical_repository["event"]["payload"]["repository"]["full_name"] = (
            "other/repository"
        )
        mutations.append(noncanonical_repository)
        mutate(("event", "github_ref"), "refs/heads/fix/conflict")
        mutate(("event", "github_sha"), "6" * 40)
        mutate(("event", "payload", "head_commit"), {"id": "6" * 40})
        mutate(("event", "payload", "commits"), [{"id": "6" * 40}])
        mutate(("event", "payload", "commits"), "not-a-list")
        mutate(("event", "payload", "commits"), [{}])
        dirty = copy.deepcopy(linear)
        dirty["status_lines"] = [" M README"]
        mutations.append(dirty)
        fixture_drift = copy.deepcopy(linear)
        fixture_drift["checkout_blob_oids"][
            "tests/fixtures/v0417/p3a_r/r1-domain-contract.json"
        ] = "7" * 40
        fixture_drift["working_blob_oids"][
            "tests/fixtures/v0417/p3a_r/r1-domain-contract.json"
        ] = "7" * 40
        mutations.append(fixture_drift)
        authority_drift = copy.deepcopy(linear)
        authority_drift["authority_blob_oids"][SEMANTIC_PATH] = "8" * 40
        mutations.append(authority_drift)
        for parents in ([], [before, second_parent, "9" * 40], [after], [before, before]):
            item = copy.deepcopy(linear)
            item["parent_oids"] = parents
            mutations.append(item)
        missing_field = copy.deepcopy(linear)
        missing_field["event"]["payload"].pop("head_commit")
        mutations.append(missing_field)
        for mutation in mutations:
            self.assertTrue(_validate_github_push_checkout(**mutation))

        local = {
            "head": after,
            "main": before,
            "origin_main": before,
            "status_lines": [],
            "main_fixture_blob_oids": copy.deepcopy(FIXED_CHECKOUT_BLOB_OIDS),
            "main_checkout_blob_oids": copy.deepcopy(checkout_blobs),
            "main_authority_blob_oids": copy.deepcopy(FROZEN_AUTHORITY_BLOB_OIDS),
            "base_is_main_ancestor": True,
            "main_is_head_ancestor": True,
            "checkout_blob_oids": copy.deepcopy(checkout_blobs),
            "working_blob_oids": copy.deepcopy(checkout_blobs),
            "authority_blob_oids": copy.deepcopy(FROZEN_AUTHORITY_BLOB_OIDS),
        }
        self.assertEqual([], _validate_local_integrated_checkout(**local))
        for key, value in (
            ("origin_main", "9" * 40),
            ("status_lines", [" M README"]),
            ("base_is_main_ancestor", False),
            ("main_is_head_ancestor", False),
            ("main_fixture_blob_oids", {}),
            ("main_checkout_blob_oids", {}),
            ("main_authority_blob_oids", {}),
            ("authority_blob_oids", {}),
        ):
            mutation = copy.deepcopy(local)
            mutation[key] = value
            self.assertTrue(_validate_local_integrated_checkout(**mutation))

    def test_001g_real_push_depth_one_and_local_integrated_lifecycle(self) -> None:
        source_repo = REPO
        git_environment = os.environ.copy()
        git_environment.pop("GIT_INDEX_FILE", None)

        def git(repo: Path, *args: str) -> str:
            return subprocess.check_output(
                ["git", "-C", str(repo), *args],
                text=True,
                stderr=subprocess.STDOUT,
                env=git_environment,
            ).strip()

        def git_bytes(repo: Path, *args: str) -> bytes:
            return subprocess.check_output(
                ["git", "-C", str(repo), *args],
                stderr=subprocess.STDOUT,
                env=git_environment,
            )

        def git_call(repo: Path, *args: str) -> None:
            subprocess.check_call(
                ["git", "-C", str(repo), *args],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                env=git_environment,
            )

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            seed = root / "integrated-seed"
            seed.mkdir()
            git_call(seed, "init", "-q")
            git_call(seed, "config", "user.name", "Integrated Test")
            git_call(seed, "config", "user.email", "integrated@example.invalid")
            (seed / "README").write_text("unintegrated\n", encoding="utf-8")
            git_call(seed, "add", "README")
            git_call(seed, "commit", "-q", "-m", "unintegrated")
            unintegrated_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "branch", "unintegrated", unintegrated_oid)

            source_by_path = {
                "tests/fixtures/v0417/p3a_r/r1-domain-contract.json": CONTRACT_PATH.read_bytes(),
                "tests/fixtures/v0417/p3a_r/r1-domain-report.md": REPORT_PATH.read_bytes(),
                "tests/test_v0417_p3ar_r1_domain_contract.py": TEST_PATH.read_bytes(),
            }
            for path, oid in FROZEN_AUTHORITY_BLOB_OIDS.items():
                source_by_path[path] = git_bytes(source_repo, "cat-file", "blob", oid)
            for relative, content in source_by_path.items():
                destination = seed / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(content)
            git_call(seed, "add", *sorted(source_by_path))
            git_call(seed, "commit", "-q", "-m", "integrated")
            integrated_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "branch", "-M", "main")

            git_call(seed, "checkout", "-q", "-b", "main-two-paths", unintegrated_oid)
            for relative, content in source_by_path.items():
                if relative == "tests/test_v0417_p3ar_r1_domain_contract.py":
                    continue
                destination = seed / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(content)
            git_call(
                seed,
                "add",
                *sorted(set(source_by_path) - {"tests/test_v0417_p3ar_r1_domain_contract.py"}),
            )
            git_call(seed, "commit", "-q", "-m", "main has only two governed fixtures")
            main_two_paths_oid = git(seed, "rev-parse", "HEAD")
            test_destination = seed / "tests/test_v0417_p3ar_r1_domain_contract.py"
            test_destination.parent.mkdir(parents=True, exist_ok=True)
            test_destination.write_bytes(TEST_PATH.read_bytes())
            git_call(seed, "add", "tests/test_v0417_p3ar_r1_domain_contract.py")
            git_call(seed, "commit", "-q", "-m", "descendant restores test path")
            main_two_paths_descendant_oid = git(seed, "rev-parse", "HEAD")

            git_call(seed, "checkout", "-q", "-b", "linear", integrated_oid)
            git_call(seed, "commit", "-q", "--allow-empty", "-m", "linear")
            linear_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "checkout", "-q", "-b", "side", integrated_oid)
            git_call(seed, "commit", "-q", "--allow-empty", "-m", "side")
            side_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "checkout", "-q", "--detach", integrated_oid)
            git_call(seed, "merge", "-q", "--no-ff", side_oid, "-m", "merge")
            merge_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "branch", "merge", merge_oid)

            git_call(seed, "checkout", "-q", "-b", "fixture-drift", integrated_oid)
            with (seed / "tests/fixtures/v0417/p3a_r/r1-domain-contract.json").open(
                "ab"
            ) as stream:
                stream.write(b"\n")
            git_call(seed, "add", "tests/fixtures/v0417/p3a_r/r1-domain-contract.json")
            git_call(seed, "commit", "-q", "-m", "fixture drift")
            fixture_drift_oid = git(seed, "rev-parse", "HEAD")

            authority_path = SEMANTIC_PATH
            git_call(seed, "checkout", "-q", "-b", "authority-drift", integrated_oid)
            with (seed / authority_path).open("ab") as stream:
                stream.write(b"\n")
            git_call(seed, "add", authority_path)
            git_call(seed, "commit", "-q", "-m", "authority drift")
            authority_drift_oid = git(seed, "rev-parse", "HEAD")
            git_call(seed, "checkout", "-q", "-b", "main-authority-drift", integrated_oid)
            with (seed / authority_path).open("ab") as stream:
                stream.write(b"\n")
            git_call(seed, "add", authority_path)
            git_call(seed, "commit", "-q", "-m", "main authority drift")
            main_authority_drift_oid = git(seed, "rev-parse", "HEAD")
            (seed / authority_path).write_bytes(source_by_path[authority_path])
            git_call(seed, "add", authority_path)
            git_call(seed, "commit", "-q", "-m", "descendant restores authority")
            authority_restored_descendant_oid = git(seed, "rev-parse", "HEAD")
            integrated_tree = git(seed, "rev-parse", f"{integrated_oid}^{{tree}}")
            divergent_oid = git(
                seed,
                "commit-tree",
                integrated_tree,
                "-m",
                "divergent",
            )
            git_call(seed, "branch", "divergent", divergent_oid)
            git_call(seed, "checkout", "-q", "main")

            repository = "owner/repository"

            def clone_depth_one(branch: str, name: str) -> Path:
                destination = root / name
                subprocess.check_call(
                    [
                        "git",
                        "clone",
                        "-q",
                        "--depth=1",
                        "--branch",
                        branch,
                        seed.resolve().as_uri(),
                        str(destination),
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    env=git_environment,
                )
                self.assertEqual(
                    "true", git(destination, "rev-parse", "--is-shallow-repository")
                )
                self.assertEqual("", git(destination, "show", "-s", "--format=%P", "HEAD"))
                return destination

            def push_environ(
                *,
                before: str,
                after: str,
                name: str,
                ref: str = "refs/heads/main",
            ) -> dict[str, str]:
                payload = {
                    "before": before,
                    "after": after,
                    "ref": ref,
                    "created": False,
                    "deleted": False,
                    "forced": False,
                    "head_commit": {"id": after},
                    "commits": [{"id": after}],
                    "repository": {"full_name": repository},
                }
                event_path = root / f"{name}-push.json"
                event_path.write_text(json.dumps(payload), encoding="utf-8")
                return {
                    "GITHUB_ACTIONS": "true",
                    "GITHUB_EVENT_NAME": "push",
                    "GITHUB_EVENT_PATH": str(event_path),
                    "GITHUB_SHA": after,
                    "GITHUB_REPOSITORY": repository,
                    "GITHUB_REF": ref,
                }

            push_cases = [
                ("linear", linear_oid, [integrated_oid], None),
                ("merge", merge_oid, [integrated_oid, side_oid], None),
                (
                    "fixture-drift",
                    fixture_drift_oid,
                    [integrated_oid],
                    "fixture blob differs",
                ),
                (
                    "authority-drift",
                    authority_drift_oid,
                    [integrated_oid],
                    "authority blob differs",
                ),
            ]
            for branch, after, parents, error_fragment in push_cases:
                with self.subTest(push=branch):
                    checkout = clone_depth_one(branch, f"push-{branch}")
                    environ = push_environ(
                        before=integrated_oid,
                        after=after,
                        name=branch,
                    )
                    with mock.patch.dict(
                        globals(),
                        {
                            "REPO": checkout,
                            "CANONICAL_REPOSITORY": repository,
                        },
                    ), mock.patch.dict(os.environ, git_environment, clear=True):
                        self.assertEqual(parents, _commit_parent_oids("HEAD"))
                        mode, errors = _runtime_checkpoint_evidence(environ)
                    self.assertEqual("github_push_integrated", mode)
                    if error_fragment is None:
                        self.assertEqual([], errors)
                    else:
                        self.assertTrue(
                            any(error_fragment in error for error in errors), errors
                        )

            dirty_checkout = root / "push-linear"
            (dirty_checkout / "README").write_text("dirty\n", encoding="utf-8")
            dirty_environ = push_environ(
                before=integrated_oid,
                after=linear_oid,
                name="dirty",
            )
            with mock.patch.dict(
                globals(),
                {
                    "REPO": dirty_checkout,
                    "CANONICAL_REPOSITORY": repository,
                },
            ), mock.patch.dict(os.environ, git_environment, clear=True):
                mode, dirty_errors = _runtime_checkpoint_evidence(dirty_environ)
            self.assertEqual("github_push_integrated", mode)
            self.assertTrue(any("not clean" in error for error in dirty_errors))

            local = root / "local-integrated"
            subprocess.check_call(
                [
                    "git",
                    "clone",
                    "-q",
                    "--branch",
                    "main",
                    seed.resolve().as_uri(),
                    str(local),
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                env=git_environment,
            )
            git_call(local, "config", "user.name", "Integrated Test")
            git_call(local, "config", "user.email", "integrated@example.invalid")

            def local_evidence(
                expected_error: str | None = None,
                *,
                base_anchor: str = unintegrated_oid,
                expected_mode: str = "local_integrated",
            ) -> None:
                with mock.patch.dict(
                    globals(),
                    {"REPO": local, "BASE_COMMIT": base_anchor},
                ), mock.patch.dict(os.environ, git_environment, clear=True):
                    mode, errors = _runtime_checkpoint_evidence({})
                self.assertEqual(expected_mode, mode)
                if expected_error is None:
                    self.assertEqual([], errors)
                else:
                    self.assertTrue(any(expected_error in error for error in errors), errors)

            local_evidence()
            git_call(local, "checkout", "-q", "-b", "work")
            git_call(local, "commit", "-q", "--allow-empty", "-m", "descendant")
            local_evidence()
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                unintegrated_oid,
            )
            local_evidence("main and origin/main differ")
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                integrated_oid,
            )
            git_call(local, "checkout", "-q", "--detach", divergent_oid)
            local_evidence("not main or a descendant")
            git_call(local, "branch", "-f", "main", unintegrated_oid)
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                unintegrated_oid,
            )
            git_call(local, "checkout", "-q", "main")
            local_evidence(
                "pre-checkpoint worktree is not the exact three untracked assets",
                expected_mode="local",
            )

            git_call(local, "checkout", "-q", "--detach", divergent_oid)
            git_call(local, "branch", "-f", "main", divergent_oid)
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                divergent_oid,
            )
            git_call(local, "checkout", "-q", "main")
            local_evidence("not proven to descend from the frozen base")

            git_call(
                local,
                "checkout",
                "-q",
                "--detach",
                main_two_paths_descendant_oid,
            )
            git_call(local, "branch", "-f", "main", main_two_paths_oid)
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                main_two_paths_oid,
            )
            local_evidence("main path is missing or not regular")

            git_call(
                local,
                "checkout",
                "-q",
                "--detach",
                authority_restored_descendant_oid,
            )
            git_call(local, "branch", "-f", "main", main_authority_drift_oid)
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                main_authority_drift_oid,
            )
            local_evidence("main authority differs from frozen evidence")

            git_call(local, "checkout", "-q", "--detach", integrated_oid)
            git_call(local, "branch", "-f", "main", integrated_oid)
            git_call(
                local,
                "update-ref",
                "refs/remotes/origin/main",
                integrated_oid,
            )
            git_call(local, "checkout", "-q", "main")
            local_evidence(
                "not proven to descend from the frozen base",
                base_anchor="f" * 40,
            )

    def test_002_authority_identity_and_171_private_extraction(self) -> None:
        mode, errors = _runtime_checkpoint_evidence()
        self.assertEqual([], errors)
        revision = (
            BASE_COMMIT
            if mode == "local"
            else "HEAD"
        )
        for path, oid, _role in AUTHORITY_INPUTS:
            self.assertEqual(oid, _tree_blob_oid(revision, path))
        self.assertEqual(V0417_OID, _tree_blob_oid(revision, V0417_PATH))
        self.assertEqual(V0418_OID, _tree_blob_oid(revision, V0418_PATH))
        self.assertEqual(171, len(self.decisions))
        self.assertEqual(171, len(set(self.decisions)))

    def test_003_v0418_ownership_precedence_is_nonsemantic_and_closed(self) -> None:
        self.assertEqual(171, len(self.ownership["decision_dispositions"]))
        self.assertEqual(580, len(self.ownership["atomic_obligations"]))
        self.assertFalse(self.contract["ownership_precedence"]["semantic_value_authority"])
        self.assertEqual([], _validate_contract(self.contract, self.decisions, self.ownership))

    def test_004_bidirectional_decision_and_atomic_disposition(self) -> None:
        entities = self.contract["entities"]
        self.assertEqual(580, len(entities))
        self.assertEqual(
            set(self.decisions),
            {item["decision_id"] for item in self.contract["decision_dispositions"]},
        )
        self.assertEqual(
            {item["obligation_id"] for item in self.ownership["atomic_obligations"]},
            {item["atomic_obligation_id"] for item in entities},
        )

    def test_005_structured_kind_domain_and_unresolved_closure(self) -> None:
        self.assertEqual([], self.contract["unresolved_obligations"])
        self.assertEqual(
            ALLOWED_KINDS,
            {item["contract_kind"] for item in self.contract["entities"]},
        )
        domains = {
            item["definition"]["domain_scope"]
            for item in self.contract["entities"]
            if item["contract_kind"] in {"property", "constraint"}
        }
        self.assertTrue(domains <= DOMAIN_SCOPES)
        self.assertEqual(
            {"project_closed", "source_preserved", "project_open_constrained"},
            domains,
        )

    def test_006_page_paper_size_is_source_supported_comparison(self) -> None:
        entity = next(
            item
            for item in self.contract["entities"]
            if item["atomic_obligation_id"] == "V040-A-001:registry"
        )
        self.assertEqual("property", entity["contract_kind"])
        self.assertEqual("source_preserved", entity["definition"]["domain_scope"])
        conclusion = entity["definition"]["comparison_conclusion"]
        self.assertEqual("source_preserved", conclusion["selected_scope"])
        self.assertEqual(
            ["project_closed", "project_open_constrained"],
            conclusion["rejected_scopes"],
        )
        self.assertIn("纸张大小", conclusion["authority_literal"])
        self.assertIn("默认保持原稿", conclusion["authority_literal"])
        shape = entity["definition"]["domain_contract"]["canonical_preservation_shape"]
        self.assertEqual("section_paper_geometry", shape["object_kind"])
        self.assertEqual(
            ["paper_width", "paper_height", "orientation"],
            shape["payload_fields"],
        )

    def test_006b_no_generic_templates_and_cross_references_are_explicit(self) -> None:
        self.assertFalse(_contains_generic_placeholder(self.contract))
        property_names = [
            item["definition"]["value_schema"]["object_name"]
            for item in self.contract["entities"]
            if item["contract_kind"] == "property"
        ]
        self.assertEqual(len(property_names), len(set(property_names)))
        registry_action_signatures = [
            json.dumps(
                item["definition"]["decidable_invariants"][0]["expected"],
                ensure_ascii=False,
                sort_keys=True,
            )
            for item in self.contract["entities"]
            if item["contract_kind"] == "constraint"
            and item["atomic_obligation_id"].endswith(":registry")
        ]
        self.assertEqual(
            len(registry_action_signatures), len(set(registry_action_signatures))
        )
        actual = Counter()
        explicit_na = Counter()
        for entity in self.contract["entities"]:
            definition = entity["definition"]
            for field in (
                "capability_refs",
                "registry_property_ids",
                "capability_dependencies",
                "external_obligations",
            ):
                if field not in definition:
                    continue
                self.assertTrue(definition[field], f"{entity['entity_id']}.{field}")
                for item in definition[field]:
                    if "entity_id" in item:
                        actual[field] += 1
                    else:
                        explicit_na[field] += 1
        self.assertGreater(actual["capability_refs"], 0)
        self.assertGreater(actual["registry_property_ids"], 0)
        self.assertGreater(actual["capability_dependencies"], 0)
        self.assertGreater(actual["external_obligations"], 0)
        self.assertGreater(explicit_na["registry_property_ids"], 0)
        self.assertGreater(explicit_na["capability_dependencies"], 0)
        self.assertGreater(explicit_na["external_obligations"], 0)

    def test_006c_source_preserved_shapes_are_object_specific(self) -> None:
        signatures = []
        for entity in self.contract["entities"]:
            definition = entity["definition"]
            if definition.get("domain_scope") != "source_preserved":
                continue
            domain = definition["domain_contract"]
            signatures.append(
                json.dumps(
                    {
                        "source_type": domain["source_type"],
                        "shape": domain["canonical_preservation_shape"],
                        "allowed_fields": domain["allowed_fields"],
                        "round_trip": domain["round_trip_equivalence"],
                    },
                    ensure_ascii=False,
                    sort_keys=True,
                )
            )
        self.assertEqual(len(signatures), len(set(signatures)))

    def test_006d_atomic_sibling_leakage_special_matrix(self) -> None:
        special = {
            "V040-D-006",
            "V040-D-008",
            "V040-T-005",
            "V040-Z-002",
            "V040-Z-004",
            "V040-Z-011",
            "V040-Z-012",
        }
        atomic_by_decision = {}
        for atomic in self.ownership["atomic_obligations"]:
            atomic_by_decision.setdefault(atomic["source_decision_id"], []).append(atomic)
        for decision_id in sorted(special):
            scoped_entities = [
                item for item in self.contract["entities"] if item["source_decision_id"] == decision_id
            ]
            for scoped in scoped_entities:
                if scoped["contract_kind"] == "constraint":
                    for invariant in scoped["definition"]["decidable_invariants"]:
                        self.assertEqual(
                            scoped["atomic_obligation_id"], invariant["source_scope"]
                        )
                elif scoped["contract_kind"] == "capability":
                    self.assertEqual(
                        scoped["atomic_obligation_id"],
                        scoped["definition"]["input_boundary"]["atomic_obligation_id"],
                    )
            candidates = [
                item
                for item in self.contract["entities"]
                if item["source_decision_id"] == decision_id
                and item["contract_kind"] == "capability"
            ]
            self.assertTrue(candidates, decision_id)
            target = candidates[0]
            sibling = next(
                item
                for item in atomic_by_decision[decision_id]
                if item["obligation_id"] != target["atomic_obligation_id"]
                and isinstance(item.get("source_mode_or_boundary"), str)
                and len(item["source_mode_or_boundary"]) >= 12
            )
            mutation = copy.deepcopy(self.contract)
            mutated_target = next(
                item for item in mutation["entities"] if item["entity_id"] == target["entity_id"]
            )
            mutated_target["definition"]["input_boundary"]["leaked_sibling_action"] = sibling[
                "source_mode_or_boundary"
            ]
            errors = _validate_contract(mutation, self.decisions, self.ownership)
            self.assertTrue(
                any("copies sibling responsibility" in error for error in errors),
                decision_id,
            )
        d006_registry = next(
            item
            for item in self.contract["entities"]
            if item["atomic_obligation_id"] == "V040-D-006:registry"
        )
        expected = d006_registry["definition"]["decidable_invariants"][0]["expected"]
        self.assertEqual(["preserve_complex_header_footer"], expected["required_actions"])
        self.assertEqual(["rebuild_complex_header_footer"], expected["forbidden_actions"])

    def test_007_mutation_duplicate_extra_missing_and_kind_rejection(self) -> None:
        mutations = []
        missing = copy.deepcopy(self.contract)
        missing["entities"].pop()
        mutations.append(missing)
        duplicate = copy.deepcopy(self.contract)
        duplicate["entities"].append(copy.deepcopy(duplicate["entities"][0]))
        mutations.append(duplicate)
        extra = copy.deepcopy(self.contract)
        extra["entities"][0]["atomic_obligation_id"] = "V040-X-999:invented"
        mutations.append(extra)
        reserved = copy.deepcopy(self.contract)
        reserved["entities"][0]["contract_kind"] = "future"
        mutations.append(reserved)
        unresolved = copy.deepcopy(self.contract)
        unresolved["unresolved_obligations"] = [{"id": "not-closed"}]
        mutations.append(unresolved)
        for mutation in mutations:
            with self.subTest(mutation=len(mutations)):
                self.assertTrue(_validate_contract(mutation, self.decisions, self.ownership))

    def test_008_mutation_source_domain_opaque_and_capability_rejection(self) -> None:
        source = copy.deepcopy(self.contract)
        source["entities"][0]["source_refs"][0]["literal"] += "x"
        self.assertTrue(_validate_contract(source, self.decisions, self.ownership))
        prop_index = next(
            i for i, item in enumerate(self.contract["entities"]) if item["contract_kind"] == "property"
        )
        opaque = copy.deepcopy(self.contract)
        opaque["entities"][prop_index]["definition"]["value_schema"] = "free text"
        self.assertTrue(_validate_contract(opaque, self.decisions, self.ownership))
        bad_domain = copy.deepcopy(self.contract)
        bad_domain["entities"][prop_index]["definition"]["domain_scope"] = "reserved"
        self.assertTrue(_validate_contract(bad_domain, self.decisions, self.ownership))
        capability_index = next(
            i for i, item in enumerate(self.contract["entities"]) if item["contract_kind"] == "capability"
        )
        implemented = copy.deepcopy(self.contract)
        implemented["entities"][capability_index]["definition"]["availability"] = "implemented"
        self.assertTrue(_validate_contract(implemented, self.decisions, self.ownership))

    def test_008b_each_contract_kind_rejects_definition_mutation(self) -> None:
        for kind in sorted(ALLOWED_KINDS):
            mutation = copy.deepcopy(self.contract)
            entity = next(item for item in mutation["entities"] if item["contract_kind"] == kind)
            removed = next(iter(entity["definition"]))
            entity["definition"].pop(removed)
            with self.subTest(kind=kind, removed=removed):
                self.assertTrue(
                    _validate_contract(mutation, self.decisions, self.ownership)
                )

    def test_008c_semantic_field_dependency_and_reference_mutations(self) -> None:
        property_index = next(
            i
            for i, item in enumerate(self.contract["entities"])
            if item["contract_kind"] == "property" and item["definition"]["ranges"]
        )
        constraint_index = next(
            i
            for i, item in enumerate(self.contract["entities"])
            if item["contract_kind"] == "constraint"
        )
        mutations = []
        generic = copy.deepcopy(self.contract)
        generic["entities"][property_index]["definition"]["representation_kind"] = (
            "atomic_typed_authority_contract"
        )
        mutations.append(generic)
        wrong_scope = copy.deepcopy(self.contract)
        wrong_scope["entities"][property_index]["definition"]["domain_scope"] = (
            "source_preserved"
        )
        mutations.append(wrong_scope)
        empty_dependency = copy.deepcopy(self.contract)
        empty_dependency["entities"][constraint_index]["definition"][
            "registry_property_ids"
        ] = []
        mutations.append(empty_dependency)
        wrong_unit = copy.deepcopy(self.contract)
        wrong_unit["entities"][property_index]["definition"]["unit_contract"][
            "field_units"
        ][0]["unit"] = "inch"
        mutations.append(wrong_unit)
        wrong_range = copy.deepcopy(self.contract)
        wrong_range["entities"][property_index]["definition"]["ranges"][0][
            "maximum"
        ] += 1
        mutations.append(wrong_range)
        wrong_invariant = copy.deepcopy(self.contract)
        wrong_invariant["entities"][property_index]["definition"][
            "cross_field_invariants"
        ][0]["expression"]["terms"][0]["expected"] = {"value": "wrong"}
        mutations.append(wrong_invariant)
        wrong_reference = copy.deepcopy(self.contract)
        wrong_reference["entities"][property_index]["definition"]["capability_refs"][0][
            "entity_id"
        ] = next(iter(
            item["entity_id"]
            for item in self.contract["entities"]
            if item["contract_kind"] == "property"
        ))
        mutations.append(wrong_reference)
        for index, mutation in enumerate(mutations):
            with self.subTest(index=index):
                self.assertTrue(
                    _validate_contract(mutation, self.decisions, self.ownership)
                )

    def test_008d_unit_reference_and_preservation_binding_mutations(self) -> None:
        def rejected(mutation: dict[str, Any]) -> None:
            self.assertTrue(
                _validate_contract(mutation, self.decisions, self.ownership)
            )

        numeric_owner_id = "r1-domain/V040-B-002:registry"
        for key in ("dimension", "allowed_units"):
            mutation = copy.deepcopy(self.contract)
            owner = next(item for item in mutation["entities"] if item["entity_id"] == numeric_owner_id)
            numeric = next(
                item
                for item in owner["definition"]["value_schema"]["fields"]
                if item["type"] == "number"
            )
            numeric["domain"].pop(key)
            rejected(mutation)

        paper = copy.deepcopy(self.contract)
        owner = next(
            item
            for item in paper["entities"]
            if item["atomic_obligation_id"] == "V040-A-001:registry"
        )
        owner["definition"]["unit_contract"]["field_units"] = []
        owner["definition"]["unit_contract"]["unitless_fields"].append("paper_width")
        rejected(paper)

        for field_name in (
            "registry_property_ids",
            "capability_dependencies",
            "external_obligations",
        ):
            base_owner = next(
                item
                for item in self.contract["entities"]
                if any(
                    ref.get("applicability") == "not_applicable"
                    for ref in item["definition"].get(field_name, [])
                )
            )
            for key, value in (
                ("reason_code", "wrong_reason"),
                ("authority_ref_sha256", "0" * 64),
                ("atomic_obligation_id", "wrong:atomic"),
                ("contract_kind", "property"),
                ("explanation", "wrong explanation"),
            ):
                mutation = copy.deepcopy(self.contract)
                owner = next(
                    item for item in mutation["entities"] if item["entity_id"] == base_owner["entity_id"]
                )
                ref_item = next(
                    ref
                    for ref in owner["definition"][field_name]
                    if ref.get("applicability") == "not_applicable"
                )
                ref_item[key] = value
                rejected(mutation)

        for field_name in REFERENCE_RELATIONS:
            base_owner = next(
                item
                for item in self.contract["entities"]
                if any("entity_id" in ref for ref in item["definition"].get(field_name, []))
            )
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item for item in mutation["entities"] if item["entity_id"] == base_owner["entity_id"]
            )
            owner["definition"][field_name][0]["relation"] = "wrong_relation"
            rejected(mutation)
            duplicate = copy.deepcopy(self.contract)
            owner = next(
                item for item in duplicate["entities"] if item["entity_id"] == base_owner["entity_id"]
            )
            owner["definition"][field_name].append(copy.deepcopy(owner["definition"][field_name][0]))
            rejected(duplicate)

        preserved_owner = next(
            item
            for item in self.contract["entities"]
            if item["atomic_obligation_id"] == "V040-A-001:registry"
        )
        preserved_mutations = (
            ("source_system", "invented"),
            ("source_version", "invented"),
            ("source_type", "invented"),
            ("provenance", {"invented": True}),
            ("unit_and_finite_constraints", {"invented": True}),
            ("round_trip_equivalence", {"invented": True}),
            ("validator_boundary", {"writes_docx": True}),
        )
        for key, value in preserved_mutations:
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item for item in mutation["entities"] if item["entity_id"] == preserved_owner["entity_id"]
            )
            owner["definition"]["domain_contract"][key] = value
            rejected(mutation)
        mutation = copy.deepcopy(self.contract)
        owner = next(
            item for item in mutation["entities"] if item["entity_id"] == preserved_owner["entity_id"]
        )
        owner["definition"]["domain_contract"]["allowed_fields"].append("invented")
        rejected(mutation)

    def test_008e_all_21_typed_field_references_and_graph_mutations(self) -> None:
        observed = {}
        for entity in self.contract["entities"]:
            if entity["contract_kind"] != "property":
                continue
            for field in entity["definition"]["value_schema"]["fields"]:
                if field["domain"].get("kind") == "reference":
                    observed[(entity["source_decision_id"], field["name"])] = field[
                        "domain"
                    ]
        self.assertEqual(21, len(observed))
        self.assertEqual(FIELD_REFERENCE_EXPECTATIONS, observed)
        for reference_key in sorted(observed):
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item
                for item in mutation["entities"]
                if item["source_decision_id"] == reference_key[0]
                and item["contract_kind"] == "property"
            )
            field = next(
                item
                for item in owner["definition"]["value_schema"]["fields"]
                if item["name"] == reference_key[1]
            )
            field["domain"]["targets"][0]["entity_id"] = "r1-domain/missing"
            invariant = owner["definition"]["cross_field_invariants"][0]
            term = next(
                item
                for item in invariant["expression"]["terms"]
                if item["field"] == field["name"]
            )
            term["expected"] = {
                key: value for key, value in field["domain"].items() if key != "kind"
            }
            self.assertTrue(
                _validate_contract(mutation, self.decisions, self.ownership),
                reference_key,
            )
        representative = next(iter(observed))
        for mutation_kind in ("contract_kind", "field_path", "relation", "missing", "duplicate"):
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item
                for item in mutation["entities"]
                if item["source_decision_id"] == representative[0]
                and item["contract_kind"] == "property"
            )
            field = next(
                item
                for item in owner["definition"]["value_schema"]["fields"]
                if item["name"] == representative[1]
            )
            if mutation_kind == "contract_kind":
                field["domain"]["targets"][0]["contract_kind"] = "capability"
            elif mutation_kind == "field_path":
                field["domain"]["targets"][0]["field_paths"][0] = "definition.missing"
            elif mutation_kind == "relation":
                field["domain"]["relation"] = "arbitrary"
            elif mutation_kind == "missing":
                field["domain"]["targets"].pop()
            else:
                field["domain"]["targets"][0]["field_paths"].append(
                    field["domain"]["targets"][0]["field_paths"][0]
                )
            term = next(
                item
                for item in owner["definition"]["cross_field_invariants"][0][
                    "expression"
                ]["terms"]
                if item["field"] == field["name"]
            )
            term["expected"] = {
                key: value for key, value in field["domain"].items() if key != "kind"
            }
            self.assertTrue(
                _validate_contract(mutation, self.decisions, self.ownership),
                mutation_kind,
            )

    def test_008f_all_nested_preservation_contracts_and_mutations(self) -> None:
        observed = set()
        for entity in self.contract["entities"]:
            if entity["contract_kind"] != "property":
                continue
            object_name = entity["definition"]["value_schema"]["object_name"]
            for field in entity["definition"]["value_schema"]["fields"]:
                errors, keys = _walk_field_preservations(
                    field["domain"],
                    owner_entity=entity,
                    object_name=object_name,
                    field_name=field["name"],
                    field_type=field["type"],
                )
                self.assertEqual([], errors)
                observed.update(keys)
        self.assertEqual(set(FIELD_PRESERVATION_SELECTORS), observed)
        self.assertEqual(
            OPEN_NESTED_PRESERVATION_DECISIONS,
            {
                decision_id
                for decision_id, _field, _tag in observed
                if next(
                    item
                    for item in self.contract["entities"]
                    if item["source_decision_id"] == decision_id
                    and item["contract_kind"] == "property"
                )["definition"]["domain_scope"]
                == "project_open_constrained"
            },
        )
        owner_key = ("V040-B-005", "complex_relative_layout", "direct")
        for contract_key, value in (
            ("source_system", "invented"),
            ("source_version", "invented"),
            ("source_type", "invented"),
            ("source_selector", "invented"),
            ("canonical_preservation_shape", {"invented": True}),
            ("unit_and_finite_constraints", {"invented": True}),
            ("round_trip_equivalence", {"invented": True}),
            ("provenance", {"invented": True}),
            ("validator_boundary", {"writes_docx": True}),
        ):
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item
                for item in mutation["entities"]
                if item["source_decision_id"] == owner_key[0]
                and item["contract_kind"] == "property"
            )
            field = next(
                item
                for item in owner["definition"]["value_schema"]["fields"]
                if item["name"] == owner_key[1]
            )
            field["domain"]["preservation_contract"][contract_key] = value
            term = next(
                item
                for item in owner["definition"]["cross_field_invariants"][0][
                    "expression"
                ]["terms"]
                if item["field"] == field["name"]
            )
            term["expected"] = {
                key: value for key, value in field["domain"].items() if key != "kind"
            }
            self.assertTrue(
                _validate_contract(mutation, self.decisions, self.ownership),
                contract_key,
            )

    def test_008g_a001_unit_custom_paper_and_mapping_mutations(self) -> None:
        def owner_from(mutation: dict[str, Any]) -> dict[str, Any]:
            return next(
                item
                for item in mutation["entities"]
                if item["atomic_obligation_id"] == "V040-A-001:registry"
            )

        fake_unit = copy.deepcopy(self.contract)
        owner = owner_from(fake_unit)
        field = next(
            item
            for item in owner["definition"]["value_schema"]["fields"]
            if item["name"] == "paper_width"
        )
        field["domain"]["canonical_unit"] = "source_native_length"
        field["domain"]["allowed_units"] = ["source_native_length"]
        field["domain"]["unit_definition"] = UNIT_DEFINITIONS["source_native_length"][1]
        contract_units = field["domain"]["preservation_contract"][
            "unit_and_finite_constraints"
        ]["field_units"][0]
        contract_units.update(
            canonical_unit="source_native_length",
            allowed_units=["source_native_length"],
            unit_definition=UNIT_DEFINITIONS["source_native_length"][1],
        )
        unit = next(
            item
            for item in owner["definition"]["unit_contract"]["field_units"]
            if item["field"] == "paper_width"
        )
        unit.update(
            canonical_unit="source_native_length",
            allowed_units=["source_native_length"],
            unit_definition=UNIT_DEFINITIONS["source_native_length"][1],
        )
        next(
            item
            for item in owner["definition"]["comparison_precision"]["field_precisions"]
            if item["field"] == "paper_width"
        )["unit"] = "source_native_length"
        top_unit = next(
            item
            for item in owner["definition"]["domain_contract"][
                "unit_and_finite_constraints"
            ]["field_units"]
            if item["field"] == "paper_width"
        )
        top_unit.update(
            canonical_unit="source_native_length",
            allowed_units=["source_native_length"],
            unit_definition=UNIT_DEFINITIONS["source_native_length"][1],
        )
        term = next(
            item
            for item in owner["definition"]["cross_field_invariants"][0]["expression"][
                "terms"
            ]
            if item["field"] == "paper_width"
        )
        term["expected"] = {
            key: value for key, value in field["domain"].items() if key != "kind"
        }
        self.assertTrue(_validate_contract(fake_unit, self.decisions, self.ownership))

        for key, value in (
            ("custom_paper_policy", "reject_nonstandard"),
            ("standard_paper_normalization", True),
            ("swap_width_height", True),
            ("supply_missing_default", True),
            ("invalid_missing_nonpositive_nonfinite_or_unknown_unit", "accept"),
        ):
            mutation = copy.deepcopy(self.contract)
            owner_from(mutation)["definition"]["comparison_conclusion"][key] = value
            self.assertTrue(_validate_contract(mutation, self.decisions, self.ownership), key)
        mutation = copy.deepcopy(self.contract)
        owner_from(mutation)["definition"]["comparison_conclusion"][
            "section_field_mapping"
        ]["canonical_shape_field"] = "section_boundary"
        self.assertTrue(_validate_contract(mutation, self.decisions, self.ownership))
        mutation = copy.deepcopy(self.contract)
        owner_from(mutation)["definition"]["cross_field_invariants"].remove(
            A001_PAPER_INVARIANT
        )
        self.assertTrue(_validate_contract(mutation, self.decisions, self.ownership))

    def test_008h_five_semantic_heading_level_value_contracts_and_mutations(
        self,
    ) -> None:
        observed: dict[tuple[str, str], dict[str, Any]] = {}
        for entity in self.contract["entities"]:
            if entity["contract_kind"] != "property":
                continue
            for field in entity["definition"]["value_schema"]["fields"]:
                if (
                    entity["source_decision_id"],
                    field["name"],
                ) in SEMANTIC_HEADING_REFERENCE_KEYS:
                    observed[(entity["source_decision_id"], field["name"])] = field[
                        "domain"
                    ]
        self.assertEqual(5, len(observed))
        self.assertEqual(SEMANTIC_HEADING_REFERENCE_KEYS, set(observed))
        for reference_key, domain in observed.items():
            self.assertEqual(
                _semantic_level_reference(reference_key[0]),
                domain,
                reference_key,
            )
            self.assertEqual(
                [1, 2, 3, 4],
                [
                    item["semantic_level"]
                    for item in domain["value_contract"]["level_mappings"]
                ],
            )
            self.assertTrue(
                all(target["contract_kind"] == "property" for target in domain["targets"])
            )
            self.assertEqual("capability", domain["consumer_gate"]["contract_kind"])

        def mutate_field(
            reference_key: tuple[str, str], mutation_kind: str
        ) -> tuple[dict[str, Any], list[str]]:
            mutation = copy.deepcopy(self.contract)
            owner = next(
                item
                for item in mutation["entities"]
                if item["source_decision_id"] == reference_key[0]
                and item["contract_kind"] == "property"
            )
            field = next(
                item
                for item in owner["definition"]["value_schema"]["fields"]
                if item["name"] == reference_key[1]
            )
            domain = field["domain"]
            if mutation_kind == "missing_value_contract":
                domain.pop("value_contract")
            elif mutation_kind == "changed_level_mapping":
                mapping = domain["value_contract"]["level_mappings"][0]
                mapping.update(
                    _property_target(
                        "V040-E-003",
                        [
                            "font_size",
                            "bold",
                            "alignment",
                            "minimum_line_spacing",
                            "space_before",
                            "space_after",
                        ],
                    )
                )
                domain["targets"] = _semantic_level_targets(
                    domain["value_contract"]["level_mappings"]
                )
            elif mutation_kind == "capability_only_value_source":
                for mapping in domain["value_contract"]["level_mappings"]:
                    mapping.update(
                        {
                            "entity_id": domain["consumer_gate"]["entity_id"],
                            "contract_kind": "capability",
                            "field_paths": ["definition.input_boundary"],
                        }
                    )
                domain["targets"] = _semantic_level_targets(
                    domain["value_contract"]["level_mappings"]
                )
            elif mutation_kind == "ambiguous_or_missing_accepted":
                approved = domain["value_contract"]["approved_input"]
                approved["ambiguous_behavior"] = "accept"
                approved["missing_behavior"] = "invent_default"
            else:
                boundary = domain["value_contract"]["authority_boundary"]
                boundary["module_may_fix_role_level"] = True
                boundary["module_may_invent_or_expand_level"] = True
            term = next(
                item
                for item in owner["definition"]["cross_field_invariants"][0][
                    "expression"
                ]["terms"]
                if item["field"] == field["name"]
            )
            term["expected"] = {
                key: value for key, value in domain.items() if key != "kind"
            }
            return mutation, _validate_contract(
                mutation, self.decisions, self.ownership
            )

        expected_error = {
            "missing_value_contract": "semantic-level value contract",
            "changed_level_mapping": "semantic-level mapping changed",
            "capability_only_value_source": "value source must contain property entities only",
            "ambiguous_or_missing_accepted": "approved semantic-level input contract changed",
            "module_fixes_or_expands_level": "semantic-level authority boundary changed",
        }
        for reference_key in sorted(SEMANTIC_HEADING_REFERENCE_KEYS):
            for mutation_kind, message in expected_error.items():
                with self.subTest(
                    reference_key=reference_key, mutation_kind=mutation_kind
                ):
                    _mutation, errors = mutate_field(reference_key, mutation_kind)
                    self.assertTrue(errors)
                    self.assertTrue(
                        any(message in error for error in errors),
                        errors,
                    )

    def test_008i_n007_border_values_and_deferred_branch_mutations(self) -> None:
        owner = next(
            item
            for item in self.contract["entities"]
            if item["source_decision_id"] == "V040-N-007"
            and item["contract_kind"] == "property"
        )
        field = next(
            item
            for item in owner["definition"]["value_schema"]["fields"]
            if item["name"] == "border_style"
        )
        domain = field["domain"]
        self.assertEqual(_n007_border_reference(), domain)
        self.assertEqual(
            [
                "r1-domain/V040-N-003:registry",
                "r1-domain/V040-N-004:registry",
                "r1-domain/V040-N-006:registry",
            ],
            [target["entity_id"] for target in domain["targets"]],
        )
        self.assertTrue(
            all(target["contract_kind"] == "property" for target in domain["targets"])
        )
        self.assertEqual(
            {
                f"r1-domain/{decision_id}:registry": [
                    f"definition.value_schema.fields.{field_name}"
                    for field_name in field_names
                ]
                for decision_id, field_names in N007_COMPATIBLE_BORDER_FIELDS.items()
            },
            {
                target["entity_id"]: target["field_paths"]
                for target in domain["targets"]
            },
        )
        deferred = domain["deferred_non_value_branch"]
        self.assertEqual("engineering_parameter_table", deferred["when"]["owning_table_semantic"])
        self.assertEqual("later_rule", deferred["target"]["contract_kind"])
        self.assertEqual("deferred", deferred["current_value_status"])
        self.assertEqual("V0.4.2", deferred["blocked_to"])
        self.assertFalse(deferred["may_emit_value"])
        self.assertFalse(deferred["may_invent_default"])
        self.assertFalse(deferred["may_execute"])
        approved = domain["selection_contract"]["approved_input"]
        self.assertEqual("blocked_qa", approved["ambiguous_behavior"])
        self.assertEqual("reject", approved["missing_behavior"])
        self.assertEqual("reject", approved["unknown_behavior"])

        def mutated_errors(mutation_kind: str) -> list[str]:
            mutation = copy.deepcopy(self.contract)
            mutation_owner = next(
                item
                for item in mutation["entities"]
                if item["source_decision_id"] == "V040-N-007"
                and item["contract_kind"] == "property"
            )
            mutation_field = next(
                item
                for item in mutation_owner["definition"]["value_schema"]["fields"]
                if item["name"] == "border_style"
            )
            mutation_domain = mutation_field["domain"]
            if mutation_kind == "later_rule_as_value":
                branch = {
                    "owning_table_semantic": "engineering_parameter_table",
                    "source_label": "工程参数表",
                    "entity_id": "r1-domain/V040-N-005:future-primary",
                    "contract_kind": "later_rule",
                    "field_paths": ["definition.preserved_original_obligation"],
                }
                mutation_domain["selection_contract"]["value_branches"].append(branch)
                mutation_domain["targets"] = _n007_value_targets(
                    mutation_domain["selection_contract"]["value_branches"]
                )
            elif mutation_kind == "missing_deferred_branch":
                mutation_domain.pop("deferred_non_value_branch")
            elif mutation_kind == "deferred_emits_value":
                mutation_domain["deferred_non_value_branch"]["may_emit_value"] = True
            elif mutation_kind == "unknown_fallback":
                approved_input = mutation_domain["selection_contract"]["approved_input"]
                approved_input["unknown_behavior"] = "emit_default"
                approved_input["missing_behavior"] = "invent_default"
            elif mutation_kind in {
                "header_alignment_as_border",
                "long_text_alignment_as_border",
            }:
                decision_id, field_name = (
                    ("V040-N-004", "header_alignment")
                    if mutation_kind == "header_alignment_as_border"
                    else ("V040-N-006", "long_text_alignment")
                )
                branch = next(
                    item
                    for item in mutation_domain["selection_contract"]["value_branches"]
                    if item["entity_id"] == f"r1-domain/{decision_id}:registry"
                )
                branch["field_paths"].append(
                    f"definition.value_schema.fields.{field_name}"
                )
                mutation_domain["targets"] = _n007_value_targets(
                    mutation_domain["selection_contract"]["value_branches"]
                )
            else:
                target = mutation_domain["deferred_non_value_branch"]["target"]
                target["field_paths"] = ["definition.preserved_original_obligation"]
                target["source_authority"]["literal_sha256"] = "0" * 64
            term = next(
                item
                for item in mutation_owner["definition"]["cross_field_invariants"][0][
                    "expression"
                ]["terms"]
                if item["field"] == "border_style"
            )
            term["expected"] = {
                key: value for key, value in mutation_domain.items() if key != "kind"
            }
            return _validate_contract(mutation, self.decisions, self.ownership)

        expected_errors = {
            "later_rule_as_value": "immediate border values must come from properties only",
            "missing_deferred_branch": "N005 deferred non-value branch is missing",
            "deferred_emits_value": "N005 deferred branch may_emit_value changed",
            "unknown_fallback": "table-semantic fail-closed input behavior changed",
            "header_alignment_as_border": "compatible-border-field allowlist",
            "long_text_alignment_as_border": "compatible-border-field allowlist",
            "arbitrary_deferred_hash_or_path": "N005 deferred target or authority changed",
        }
        for mutation_kind, message in expected_errors.items():
            with self.subTest(mutation_kind=mutation_kind):
                errors = mutated_errors(mutation_kind)
                self.assertTrue(errors)
                self.assertTrue(any(message in error for error in errors), errors)

    def test_009_raw_canonical_and_report_bytes_are_deterministic(self) -> None:
        self.assertNotIn(b"\r", self.contract_bytes)
        self.assertNotIn(b"\r", self.report_bytes)
        self.assertTrue(self.contract_bytes.endswith(b"\n"))
        self.assertTrue(self.report_bytes.endswith(b"\n"))
        expected_contract = (
            json.dumps(
                self.contract,
                ensure_ascii=False,
                sort_keys=True,
                indent=2,
            )
            + "\n"
        ).encode("utf-8")
        self.assertEqual(expected_contract, self.contract_bytes)
        expected = _render_report(self.contract, self.contract_bytes)
        self.assertEqual(expected, self.report_bytes)
        self.assertIn(_sha256(self.contract_bytes).encode("ascii"), self.report_bytes)
        self.assertIn(
            _sha256(_canonical_bytes(self.contract)).encode("ascii"),
            self.report_bytes,
        )
        with tempfile.TemporaryDirectory() as tmp:
            generated = Path(tmp) / "report.md"
            generated.write_bytes(expected)
            self.assertEqual(self.report_bytes, generated.read_bytes())
        self.assertEqual(
            sorted(self.contract["entities"], key=lambda item: item["entity_id"]),
            self.contract["entities"],
        )
        self.assertEqual(
            sorted(
                self.contract["decision_dispositions"],
                key=lambda item: item["decision_id"],
            ),
            self.contract["decision_dispositions"],
        )

    def test_010_three_assets_are_absent_from_production_load_paths(self) -> None:
        exact_paths = {
            str(CONTRACT_PATH.relative_to(REPO)),
            str(REPORT_PATH.relative_to(REPO)),
            str(TEST_PATH.relative_to(REPO)),
        }
        for root in (REPO / "format-monograph", REPO / "adapters"):
            for path in root.rglob("*"):
                if not path.is_file() or path.suffix not in {".py", ".json", ".md", ".yaml", ".yml"}:
                    continue
                try:
                    text = path.read_text(encoding="utf-8")
                except UnicodeError:
                    continue
                for exact in exact_paths:
                    self.assertNotIn(exact, text, f"production path references R1 asset: {path}")

    def test_011_privacy_and_no_private_payload(self) -> None:
        combined = self.contract_bytes + self.report_bytes + TEST_PATH.read_bytes()
        for forbidden in (
            b"/" + b"Users/",
            b"\\" + b"Users\\",
            b"BEGIN " + b"PRIVATE KEY",
            b"." + b"docx",
            b"." + b"pdf",
            b"move" + b"1/",
            b"move" + b"2/",
        ):
            self.assertNotIn(forbidden, combined)

    def test_012_all_four_domain_scope_validators_are_closed(self) -> None:
        domains = {
            "project_closed": {
                "members": [{"value": "approved"}],
                "member_source_refs": [{"decision_id": "V040-Z-006"}],
                "closure_rule": {"additional_members": False},
            },
            "source_preserved": {
                "source_system": "V0.4.0-frozen-semantic-authority",
                "source_version": AUTHORITY_INPUTS[0][1],
                "source_type": "document-property",
                "canonical_preservation_shape": {
                    "object_kind": "document-property",
                    "identity_fields": ["object_id"],
                    "payload_fields": ["value"],
                    "ordering_fields": ["document_order"],
                },
                "allowed_fields": ["object_id", "value", "document_order"],
                "unit_and_finite_constraints": {
                    "numeric_fields": [],
                    "field_units": [],
                    "finite_numbers_only": True,
                    "source_units_required": True,
                },
                "round_trip_equivalence": {
                    "identity_comparison": "exact",
                    "payload_comparison": "canonical_xml_or_byte_exact",
                    "ordering_comparison": "exact_sequence",
                    "unknown_member_policy": "reject",
                },
                "provenance": {"required": True},
                "validator_boundary": {"owner": "P3a-R", "writes_docx": False},
            },
            "external_versioned": {
                "owner": "external-authority-owner",
                "external_version": "frozen-version",
                "locator": "authority-controlled-locator",
                "digest": "0" * 64,
                "project_mapping": {"mode": "explicit-only"},
                "offline_validator_boundary": {"network_access": False},
            },
            "project_open_constrained": {
                "value_shape": {"type": "object"},
                "finite_rules": [{"finite_numbers_only": True}],
                "range_rules": [{"rule": "authority-bounded"}],
                "cross_field_invariants": [{"rule": "authority-clause-preserved"}],
            },
        }
        for scope, domain in domains.items():
            with self.subTest(scope=scope):
                self.assertEqual(
                    [],
                    _validate_domain(
                        {"domain_scope": scope, "domain_contract": domain}, scope
                    ),
                )
                mutation = copy.deepcopy(domain)
                mutation.pop(next(iter(mutation)))
                self.assertTrue(
                    _validate_domain(
                        {"domain_scope": scope, "domain_contract": mutation}, scope
                    )
                )


if __name__ == "__main__":
    unittest.main()
