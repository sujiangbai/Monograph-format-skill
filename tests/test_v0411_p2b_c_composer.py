#!/usr/bin/env python3
"""Contract tests for the disabled-only V0.4.1 P2b-C composer."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import random
import sys
import unittest
from copy import deepcopy
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
FIXTURES = ROOT / "tests" / "fixtures" / "v0411" / "p2b_c"
sys.path.insert(0, str(SCRIPTS))

from profile_v2_artifacts import (  # noqa: E402
    _schema_documents,
    _test_schema_overrides,
    _validate_artifact_dag_for_test,
    _validate_artifact_for_test,
    load_artifact_schema,
)
from profile_v2_authority import authority_contract_fingerprint  # noqa: E402
from profile_v2_canonical import _stamp_semantic_fingerprint_for_test  # noqa: E402
from profile_v2_composer import (  # noqa: E402
    ComposerDisabledError,
    _ContractAdapter,
    _apply_report,
    _report_documents,
)
from profile_v2_registry import load_registry  # noqa: E402
from profile_v2_scope import normalize_scope, scope_disjoint  # noqa: E402


COUNTS = {"PRE": 18, "CAND": 30, "CONF": 36, "REPORT": 24, "APP": 30, "FINAL": 30}
ASSERTION_IDS = tuple(
    f"T411-C-{family}-{index:03d}"
    for family, count in COUNTS.items()
    for index in range(1, count + 1)
)


def digest(label: str) -> str:
    return "sha256:" + hashlib.sha256(label.encode("utf-8")).hexdigest()


def load_test_registry() -> dict:
    return load_registry(
        FIXTURES / "property-registry.v2.1.test.json",
        allow_test=True,
        version="2.1",
    )


def raw_scope(*document_ids: str, conditions: list[dict] | None = None) -> dict:
    return {
        "selectors": [
            {
                "selector_kind": "document",
                "selector_ids": list(document_ids or ("document:book",)),
            }
        ],
        "exclusions": [],
        "mutually_exclusive_conditions": list(conditions or []),
    }


def stamp(document: dict, registry: dict) -> dict:
    schema = load_artifact_schema(
        document["artifact_kind"], version=document["schema_version"]
    )
    documents = _schema_documents(_test_schema_overrides(registry))
    return _stamp_semantic_fingerprint_for_test(
        document,
        schema=schema,
        documents=documents,
        registry=registry,
    )


def composer_test_adapter(registry: dict) -> _ContractAdapter:
    """Build the private test adapter outside the production composer module."""

    def validate(document: dict) -> None:
        _validate_artifact_for_test(
            document,
            registry=registry,
            features={"profile_v2_schema": True},
        )

    def feature_enabled(manifest: dict | None) -> bool:
        if manifest is None:
            return False
        try:
            validate(manifest)
        except ValueError:
            return False
        features = manifest.get("features", {})
        return bool(
            manifest.get("artifact_kind") == "feature-activation-manifest"
            and manifest.get("schema_version") == "2.1"
            and features.get("profile_v2_schema") is True
            and features.get("profile_v2_composer") is True
        )

    return _ContractAdapter(
        registry=registry,
        validate=validate,
        stamp=lambda document: stamp(document, registry),
        feature_enabled=feature_enabled,
    )


def feature(registry: dict, *, enabled: bool = True) -> dict:
    return stamp(
        {
            "artifact_kind": "feature-activation-manifest",
            "schema_version": "2.1",
            "registry_contract_version": "2.1",
            "authority_contract_version": "1.0",
            "artifact_id": "feature-activation-manifest:p2b-c",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                {
                    "input_id": "input:feature-source",
                    "role": "source_document",
                    "fingerprint": digest("feature-source"),
                }
            ],
            "semantic_fingerprint": digest("unstamped-feature"),
            "features": {
                "profile_v2_schema": enabled,
                "profile_v2_composer": enabled,
                "final_ready_eligible": False,
            },
        },
        registry,
    )


def asset(
    registry: dict,
    label: str,
    *,
    layer: str = "monograph_base",
    value: str = "10.50",
    unit: str = "unit.pt",
    mode: str = "report",
    confidence: str = "high",
    scope: dict | None = None,
    asset_scope: dict | None = None,
    activation: str = "approved",
    rule_status: str = "approved",
    summary: str | None = None,
) -> dict:
    rule_scope = deepcopy(scope or raw_scope("document:book"))
    boundary = deepcopy(asset_scope or (rule_scope if layer == "module" else raw_scope("document:book")))
    document = {
            "artifact_kind": "layered-rule-asset",
            "schema_version": "2.1",
            "artifact_id": f"layered-rule-asset:{label}",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                {
                    "input_id": f"input:{label}",
                    "role": "rule_asset",
                    "fingerprint": digest(f"source-{label}"),
                }
            ],
            "semantic_fingerprint": digest(f"unstamped-{label}"),
            "layer_kind": layer,
            "can_override_safety_invariants": False,
            "activation": activation,
            "asset_scope": boundary,
            "allowed_semantic_object_kinds": ["paragraph"],
            "rules": [
                {
                    "rule_id": f"RULE-{label.upper().replace('_', '-')}",
                    "semantic_object_kind": "paragraph",
                    "scope": rule_scope,
                    "confidence": confidence,
                    "status": rule_status,
                    "properties": [
                        {
                            "property_id": "test.paragraph-font-size",
                            "value": {"type": "decimal", "value": value},
                            "unit_id": unit,
                            "mode": mode,
                        }
                    ],
                }
            ],
        }
    if summary is not None:
        document["public_source_summary"] = summary
    return stamp(document, registry)


def compose(registry: dict, assets: list[dict], *, enabled: bool = True) -> tuple[dict, dict]:
    manifest = feature(registry, enabled=enabled)
    report = _report_documents(
        assets,
        manifest,
        adapter=composer_test_adapter(registry),
        input_fingerprint=digest("source-document"),
        structure_fingerprint=digest("structure-map"),
        artifact_id="conflict-report:p2b-c",
        created_by_tool={"tool_id": "format-monograph", "version": "0.4.1"},
        generated_at="2026-08-18T00:00:00Z",
    )
    return report, manifest


def approval(
    registry: dict,
    report: dict,
    decision: str,
    *,
    conflict_index: int = 0,
    candidate_index: int = 0,
    candidate_id: str | None = None,
    approval_id: str = "approval:p2b-c",
    previous: str | None = None,
) -> dict:
    conflict = report["approval_required_conflicts"][conflict_index]
    target = {
        "conflict_id": conflict["conflict_id"],
        "proposed_resolution_id": conflict["proposed_resolution_id"],
        "normalized_scope": deepcopy(conflict["key"]["normalized_scope"]),
    }
    if decision in {"select_candidate", "exclude_candidate"}:
        target["candidate_id"] = candidate_id or conflict["candidates"][candidate_index]["candidate_id"]
    decision_type = {
        "adopt_proposed": "conflict_resolution",
        "select_candidate": "conflict_resolution",
        "keep_original": "keep_original",
        "exclude_candidate": "qa_exclusion",
    }[decision]
    return stamp(
        {
            "artifact_kind": "qa-approval-artifact",
            "schema_version": "2.1",
            "registry_contract_version": "2.1",
            "authority_contract_version": "1.0",
            "artifact_id": f"qa-approval-artifact:{approval_id.split(':', 1)[1]}",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                {"input_id": "input:approval-source", "role": "source_document", "fingerprint": report["bindings"]["input_fingerprint"]},
                {"input_id": "input:approval-structure", "role": "structure", "fingerprint": report["bindings"]["structure_fingerprint"]},
                {"input_id": "input:approval-report", "role": "conflict_report", "fingerprint": report["semantic_fingerprint"]},
            ],
            "semantic_fingerprint": digest(f"unstamped-{approval_id}"),
            "approval_id": approval_id,
            "approver": {"actor_id": "actor:user", "actor_role": "user"},
            "decision_type": decision_type,
            "decision": decision,
            "reason": "Synthetic P2b-C decision.",
            "created_at": "2026-08-18T00:00:00Z",
            "bindings": {
                "input_fingerprint": report["bindings"]["input_fingerprint"],
                "structure_fingerprint": report["bindings"]["structure_fingerprint"],
                "composition_report_fingerprint": report["semantic_fingerprint"],
            },
            "target": target,
            "previous_approval_id": previous,
        },
        registry,
    )


def apply(registry: dict, report: dict, approvals: list[dict]) -> dict:
    return _apply_report(
        report,
        approvals,
        adapter=composer_test_adapter(registry),
        task_id="task:p2b-c",
        task_fingerprint=digest("task-p2b-c"),
        artifact_id="final-execution-profile:p2b-c",
        created_by_tool={"tool_id": "format-monograph", "version": "0.4.1"},
    )


class P2bCComposerTests(unittest.TestCase):
    maxDiff = None

    def setUp(self) -> None:
        self.registry = load_test_registry()

    def assert_valid(self, document: dict) -> None:
        _validate_artifact_for_test(
            document,
            registry=self.registry,
            features={"profile_v2_schema": True},
        )

    def test_assertion_id_contract_is_exact(self) -> None:
        self.assertEqual(168, len(ASSERTION_IDS))
        self.assertEqual(168, len(set(ASSERTION_IDS)))
        for family, count in COUNTS.items():
            self.assertEqual(
                count,
                len([item for item in ASSERTION_IDS if item.startswith(f"T411-C-{family}-")]),
            )

    def test_preflight_contract_pre_001_018(self) -> None:
        base = asset(self.registry, "pre_base")
        for index in range(1, 19):
            assertion_id = f"T411-C-PRE-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                if index == 1:
                    report, _ = compose(self.registry, [base])
                    self.assertEqual("resolvable", report["proposal_status"])
                    self.assert_valid(report)
                elif index in {2, 3, 4}:
                    with self.assertRaises(ComposerDisabledError):
                        compose(self.registry, [base], enabled=False)
                elif index == 5:
                    invalid = deepcopy(base)
                    invalid["semantic_fingerprint"] = digest("stale")
                    report, _ = compose(self.registry, [invalid])
                    self.assertEqual("fatal", report["proposal_status"])
                    self.assertTrue(report["fatal_diagnostics"])
                elif index == 6:
                    module = asset(
                        self.registry,
                        "pre_module",
                        layer="module",
                        scope=raw_scope("document:outside"),
                        asset_scope=raw_scope("document:inside"),
                    )
                    report, _ = compose(self.registry, [module])
                    self.assertEqual("module_asset_out_of_bounds", report["fatal_diagnostics"][0]["category"])
                elif index == 7:
                    report, _ = compose(self.registry, [base, deepcopy(base)])
                    self.assertEqual("fatal", report["proposal_status"])
                elif index in {8, 9}:
                    inactive = asset(
                        self.registry,
                        f"pre_inactive_{index}",
                        activation="disabled" if index == 8 else "approved",
                        rule_status="approved" if index == 8 else "draft",
                    )
                    report, _ = compose(self.registry, [inactive])
                    self.assertEqual("resolvable", report["proposal_status"])
                    self.assertFalse(report["proposed_resolutions"])
                    self.assertTrue(report["diagnostics"])
                elif index == 10:
                    with self.assertRaises(ValueError):
                        compose(self.registry, [])
                elif index == 11:
                    snapshot = deepcopy(base)
                    compose(self.registry, [base])
                    self.assertEqual(snapshot, base)
                elif index == 12:
                    report_a, _ = compose(self.registry, [base])
                    report_b, _ = compose(self.registry, [deepcopy(base)])
                    self.assertEqual(report_a, report_b)
                elif index == 13:
                    report, _ = compose(self.registry, [base])
                    self.assertEqual(authority_contract_fingerprint("1.0"), report["bindings"]["authority_contract_fingerprint"])
                elif index == 14:
                    report, manifest = compose(self.registry, [base])
                    self.assertEqual(manifest["semantic_fingerprint"], report["bindings"]["feature_activation_fingerprint"])
                elif index == 15:
                    report, _ = compose(self.registry, [base])
                    self.assertEqual("2.1", report["registry_contract_version"])
                elif index == 16:
                    report, _ = compose(self.registry, [base])
                    self.assertEqual("1.0", report["authority_contract_version"])
                elif index == 17:
                    report, _ = compose(self.registry, [base])
                    self.assertFalse(any(item["category"] == "safety_hard_failure" for item in report["fatal_diagnostics"]))
                else:
                    runtime = ROOT / "format-monograph" / "scripts" / "run_monograph.py"
                    self.assertNotIn("profile_v2_composer", runtime.read_text(encoding="utf-8"))

    def test_candidate_and_partition_contract_cand_001_030(self) -> None:
        low = asset(self.registry, "cand_low", layer="standard_supplement", value="9.00")
        high = asset(self.registry, "cand_high", layer="task_override", value="11.00")
        broad = asset(self.registry, "cand_broad", layer="monograph_base", value="10.00", scope=raw_scope("document:a", "document:b"))
        narrow = asset(self.registry, "cand_narrow", layer="publisher_template", value="12.00", scope=raw_scope("document:a"))
        low_high_report, _ = compose(self.registry, [low, high])
        low_high_reversed, _ = compose(self.registry, [high, low])
        subset_report, _ = compose(self.registry, [broad, narrow])
        pt = asset(self.registry, "cand_pt", value="10.00", unit="unit.pt")
        mm = asset(self.registry, "cand_mm", layer="module", value="3.527777777777777777", unit="unit.mm")
        unit_report, _ = compose(self.registry, [pt, mm])
        left = asset(self.registry, "cand_cross_l", scope=raw_scope("document:a", "document:b"))
        right = asset(self.registry, "cand_cross_r", layer="module", scope=raw_scope("document:b", "document:c"))
        crossing_report, _ = compose(self.registry, [left, right])
        composed = asset(self.registry, "cand_nfc", summary="caf\u00e9")
        decomposed = asset(self.registry, "cand_nfc", summary="cafe\u0301")
        nfc_report, _ = compose(self.registry, [composed])
        nfd_report, _ = compose(self.registry, [decomposed])
        for index in range(1, 31):
            assertion_id = f"T411-C-CAND-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                if index <= 8:
                    report = low_high_report if index % 2 else low_high_reversed
                    proposal = report["proposed_resolutions"][0]
                    self.assertEqual("task_override", proposal["final_layer_kind"])
                    self.assertEqual("11.00", proposal["proposed_binding"]["value"]["value"])
                    self.assertEqual(2, len(proposal["candidate_chain"]))
                    self.assertEqual(2, len(proposal["override_chain"]))
                elif index <= 16:
                    report = subset_report
                    self.assertEqual(2, len(report["candidate_groups"]))
                    scopes = [item["key"]["normalized_scope"] for item in report["proposed_resolutions"]]
                    self.assertTrue(scope_disjoint(scopes[0], scopes[1]))
                    values = sorted(item["proposed_binding"]["value"]["value"] for item in report["proposed_resolutions"])
                    self.assertEqual(["10.00", "12.00"], values)
                elif index <= 20:
                    report = unit_report
                    self.assertEqual("10.00", report["proposed_resolutions"][0]["proposed_binding"]["value"]["value"])
                elif index <= 24:
                    report = crossing_report
                    self.assertEqual("unresolvable", report["proposal_status"])
                    self.assertEqual("crossing_overlap", report["unresolvable_blockers"][0]["category"])
                elif index <= 27:
                    report = nfc_report
                    self.assertEqual(1, len(report["candidate_groups"]))
                    self.assertEqual(report["semantic_fingerprint"], nfd_report["semantic_fingerprint"])
                else:
                    report = subset_report
                    keys = [
                        (item["key"]["semantic_object_kind"], item["key"]["property_id"], item["key"]["normalized_scope"]["scope_id"])
                        for item in report["proposed_resolutions"]
                    ]
                    self.assertEqual(len(keys), len(set(keys)))

    def test_confidence_and_conflict_contract_conf_001_036(self) -> None:
        conflict_assets = [
            asset(self.registry, "conf_a", value="10.00"),
            asset(self.registry, "conf_b", value="11.00"),
        ]
        medium = asset(self.registry, "conf_medium", confidence="medium")
        low_lower = asset(self.registry, "conf_low_lower", layer="standard_supplement", confidence="low", value="8.00")
        high_upper = asset(self.registry, "conf_high_upper", layer="task_override", confidence="high", value="12.00")
        conflict_report, _ = compose(self.registry, conflict_assets)
        conflict_reversed, _ = compose(self.registry, list(reversed(conflict_assets)))
        medium_report, _ = compose(self.registry, [medium])
        precedence_report, _ = compose(self.registry, [low_lower, high_upper])
        low_residual = asset(self.registry, "conf_residual_low", confidence="low", scope=raw_scope("document:a", "document:b"))
        high_cut = asset(self.registry, "conf_residual_high", layer="task_override", confidence="high", scope=raw_scope("document:a"))
        residual_report, _ = compose(self.registry, [low_residual, high_cut])
        for index in range(1, 37):
            assertion_id = f"T411-C-CONF-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                if index <= 12:
                    report = conflict_reversed if index % 2 else conflict_report
                    self.assertEqual("awaiting_approval", report["proposal_status"])
                    conflict = report["approval_required_conflicts"][0]
                    self.assertEqual(2, len(conflict["candidates"]))
                    self.assertEqual(
                        {"adopt_proposed", "select_candidate", "keep_original", "exclude_candidate"},
                        set(conflict["allowed_decisions"]),
                    )
                elif index <= 20:
                    report = medium_report
                    self.assertEqual("awaiting_approval", report["proposal_status"])
                    self.assertEqual("medium", report["proposed_resolutions"][0]["confidence"])
                elif index <= 28:
                    report = precedence_report
                    self.assertEqual("resolvable", report["proposal_status"])
                    self.assertEqual("high", report["proposed_resolutions"][0]["confidence"])
                    self.assertEqual(2, len(report["proposed_resolutions"][0]["candidate_chain"]))
                else:
                    report = residual_report
                    self.assertEqual(1, len(report["approval_required_conflicts"]))
                    conflict_scope = report["approval_required_conflicts"][0]["key"]["normalized_scope"]
                    self.assertTrue(any("document:b" in selector["selector_ids"] for selector in conflict_scope["selectors"]))

    def test_report_contract_report_001_024(self) -> None:
        valid = asset(self.registry, "report_valid")
        conflict_assets = [asset(self.registry, "report_a", value="10.00"), asset(self.registry, "report_b", value="11.00")]
        crossing = [
            asset(self.registry, "report_cross_a", scope=raw_scope("document:a", "document:b")),
            asset(self.registry, "report_cross_b", layer="module", scope=raw_scope("document:b", "document:c")),
        ]
        valid_report, _ = compose(self.registry, [valid])
        approval_report, _ = compose(self.registry, conflict_assets)
        unresolvable_report, _ = compose(self.registry, crossing)
        invalid = deepcopy(valid)
        invalid["semantic_fingerprint"] = digest("invalid-report-asset")
        isolated_valid = asset(
            self.registry, "report_isolated", scope=raw_scope("document:z")
        )
        fatal_report, _ = compose(
            self.registry, [invalid, isolated_valid, *crossing]
        )
        unresolvable_application = apply(self.registry, unresolvable_report, [])
        fatal_application = apply(self.registry, fatal_report, [])
        for index in range(1, 25):
            assertion_id = f"T411-C-REPORT-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                if index <= 6:
                    report = valid_report
                    self.assertEqual("resolvable", report["proposal_status"])
                elif index <= 12:
                    report = approval_report
                    self.assertEqual("awaiting_approval", report["proposal_status"])
                elif index <= 18:
                    report = unresolvable_report
                    self.assertEqual("unresolvable", report["proposal_status"])
                    self.assertEqual("failed", report["scope_partitions"][0]["evidence_status"])
                    self.assertEqual("unresolvable", unresolvable_application["status"])
                else:
                    report = fatal_report
                    self.assertEqual("fatal", report["proposal_status"])
                    self.assertEqual("fatal", fatal_application["status"])
                    self.assertTrue(report["candidate_groups"])
                self.assertEqual(report["proposal_status"], (
                    "fatal" if report["fatal_diagnostics"] else
                    "unresolvable" if report["unresolvable_blockers"] else
                    "awaiting_approval" if report["approval_required_conflicts"] else
                    "resolvable"
                ))
                self.assert_valid(report)

    def test_approval_application_contract_app_001_030(self) -> None:
        app_assets = [
            asset(self.registry, "app_a", value="10.00"),
            asset(self.registry, "app_b", value="11.00"),
        ]
        report, manifest = compose(self.registry, app_assets)
        report_snapshot = deepcopy(report)
        adopt = approval(self.registry, report, "adopt_proposed", approval_id="approval:adopt")
        adopt_snapshot = deepcopy(adopt)
        adopt_result = apply(self.registry, report, [adopt])
        adopt_dag = _validate_artifact_dag_for_test(
            [adopt_result["final_profile"], report, manifest, *app_assets, adopt],
            registry=self.registry,
        )
        selected = approval(self.registry, report, "select_candidate", candidate_index=1, approval_id="approval:select")
        selected_snapshot = deepcopy(selected)
        selected_result = apply(self.registry, report, [selected])
        preserve_report, _ = compose(
            self.registry,
            [
                asset(self.registry, "app_preserve", value="9.00", mode="preserve"),
                asset(self.registry, "app_report", value="11.00"),
            ],
        )
        keep = approval(self.registry, preserve_report, "keep_original", approval_id="approval:keep")
        keep_result = apply(self.registry, preserve_report, [keep])
        proposal = report["proposed_resolutions"][0]
        conflict = report["approval_required_conflicts"][0]
        proposed_candidate = next(
            item
            for item in conflict["candidates"]
            if item["source"] == proposal["final_source"]
            and item["property_binding"] == proposal["proposed_binding"]
        )
        losing_candidate = next(
            item
            for item in conflict["candidates"]
            if item["candidate_id"] != proposed_candidate["candidate_id"]
        )
        exclude_winner = approval(
            self.registry,
            report,
            "exclude_candidate",
            candidate_id=proposed_candidate["candidate_id"],
            approval_id="approval:exclude-winner",
        )
        exclude_winner_result = apply(self.registry, report, [exclude_winner])
        exclude_loser = approval(
            self.registry,
            report,
            "exclude_candidate",
            candidate_id=losing_candidate["candidate_id"],
            approval_id="approval:exclude-loser",
        )
        exclude_loser_result = apply(self.registry, report, [exclude_loser])

        redundant_report, _ = compose(
            self.registry,
            [
                asset(self.registry, "app_redundant_high", value="12.00"),
                asset(
                    self.registry,
                    "app_redundant_medium",
                    value="12.00",
                    confidence="medium",
                ),
            ],
        )
        redundant_conflict = redundant_report["approval_required_conflicts"][0]
        redundant_target = next(
            item
            for item in redundant_conflict["candidates"]
            if item["confidence"] == "medium"
        )
        redundant_remaining = next(
            item
            for item in redundant_conflict["candidates"]
            if item["candidate_id"] != redundant_target["candidate_id"]
        )
        redundant_exclusion = approval(
            self.registry,
            redundant_report,
            "exclude_candidate",
            candidate_id=redundant_target["candidate_id"],
            approval_id="approval:exclude-redundant",
        )
        redundant_result = apply(
            self.registry, redundant_report, [redundant_exclusion]
        )

        unresolved_report, _ = compose(
            self.registry,
            [
                asset(self.registry, "app_unresolved_a", value="10.00"),
                asset(self.registry, "app_unresolved_b", value="11.00"),
                asset(self.registry, "app_unresolved_c", value="12.00"),
            ],
        )
        unresolved_exclusion = approval(
            self.registry,
            unresolved_report,
            "exclude_candidate",
            candidate_index=0,
            approval_id="approval:exclude-still-conflicted",
        )
        unresolved_result = apply(
            self.registry, unresolved_report, [unresolved_exclusion]
        )

        weak_report, _ = compose(
            self.registry,
            [
                asset(
                    self.registry,
                    "app_weak_low",
                    value="13.00",
                    confidence="low",
                ),
                asset(
                    self.registry,
                    "app_weak_medium",
                    value="13.00",
                    confidence="medium",
                ),
            ],
        )
        weak_conflict = weak_report["approval_required_conflicts"][0]
        weak_target = next(
            item for item in weak_conflict["candidates"] if item["confidence"] == "low"
        )
        weak_exclusion = approval(
            self.registry,
            weak_report,
            "exclude_candidate",
            candidate_id=weak_target["candidate_id"],
            approval_id="approval:exclude-to-medium",
        )
        weak_result = apply(self.registry, weak_report, [weak_exclusion])

        unknown_target = approval(
            self.registry,
            report,
            "exclude_candidate",
            candidate_id="candidate:" + ("f" * 64),
            approval_id="approval:exclude-unknown",
        )
        unknown_target_result = apply(self.registry, report, [unknown_target])

        foreign_report, _ = compose(
            self.registry,
            [
                asset(
                    self.registry,
                    "app_foreign_a1",
                    value="10.00",
                    scope=raw_scope("document:a"),
                ),
                asset(
                    self.registry,
                    "app_foreign_a2",
                    value="11.00",
                    scope=raw_scope("document:a"),
                ),
                asset(
                    self.registry,
                    "app_foreign_b1",
                    value="12.00",
                    scope=raw_scope("document:b"),
                ),
                asset(
                    self.registry,
                    "app_foreign_b2",
                    value="13.00",
                    scope=raw_scope("document:b"),
                ),
            ],
        )
        foreign_candidate_id = foreign_report["approval_required_conflicts"][1][
            "candidates"
        ][0]["candidate_id"]
        foreign_target = approval(
            self.registry,
            foreign_report,
            "exclude_candidate",
            conflict_index=0,
            candidate_id=foreign_candidate_id,
            approval_id="approval:exclude-foreign",
        )
        foreign_target_result = apply(
            self.registry, foreign_report, [foreign_target]
        )

        medium_report, _ = compose(
            self.registry, [asset(self.registry, "app_medium", confidence="medium")]
        )
        sole_exclusion = approval(
            self.registry,
            medium_report,
            "exclude_candidate",
            approval_id="approval:exclude-sole",
        )
        sole_exclusion_result = apply(
            self.registry, medium_report, [sole_exclusion]
        )
        unexpressible_keep = approval(
            self.registry, report, "keep_original", approval_id="approval:no-preserve"
        )
        unexpressible_keep_result = apply(
            self.registry, report, [unexpressible_keep]
        )
        first = approval(self.registry, report, "adopt_proposed", approval_id="approval:first")
        second = approval(self.registry, report, "select_candidate", approval_id="approval:second", previous=first["approval_id"])
        chain_result = apply(self.registry, report, [first, second])
        branch_root = approval(self.registry, report, "adopt_proposed", approval_id="approval:branch-root")
        branch_left = approval(self.registry, report, "adopt_proposed", approval_id="approval:branch-left", previous=branch_root["approval_id"])
        branch_right = approval(self.registry, report, "select_candidate", approval_id="approval:branch-right", previous=branch_root["approval_id"])
        branch_result = apply(self.registry, report, [branch_root, branch_left, branch_right])
        cycle_left = approval(
            self.registry,
            report,
            "adopt_proposed",
            approval_id="approval:cycle-left",
            previous="approval:cycle-right",
        )
        cycle_right = approval(
            self.registry,
            report,
            "select_candidate",
            approval_id="approval:cycle-right",
            previous="approval:cycle-left",
        )
        cycle_result = apply(self.registry, report, [cycle_left, cycle_right])
        unknown_previous = approval(self.registry, report, "adopt_proposed", approval_id="approval:unknown-prev", previous="approval:missing")
        unknown_result = apply(self.registry, report, [unknown_previous])
        stale = approval(self.registry, report, "adopt_proposed", approval_id="approval:stale")
        stale["bindings"]["composition_report_fingerprint"] = digest("stale-report")
        stale_result = apply(self.registry, report, [stale])
        for index in range(1, 31):
            assertion_id = f"T411-C-APP-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                if index <= 5:
                    result = adopt_result
                    if index == 1:
                        self.assertEqual("profile_generated", result["status"])
                    elif index == 2:
                        self.assertIsNotNone(result["final_profile"])
                    elif index == 3:
                        self.assertEqual(
                            [adopt["semantic_fingerprint"]],
                            result["final_profile"]["bindings"]["approval_fingerprints"],
                        )
                    elif index == 4:
                        self.assertEqual(
                            adopt["approval_id"],
                            result["final_profile"]["closure_evidence"][0]["qa_decision_id"],
                        )
                    else:
                        self.assertFalse(adopt_dag.runtime_eligible)
                elif index <= 10:
                    result = selected_result
                    selected_id = selected["target"]["candidate_id"]
                    resolved = result["final_profile"]["resolved_properties"][0]
                    selected_candidate = next(
                        item
                        for item in conflict["candidates"]
                        if item["candidate_id"] == selected_id
                    )
                    if index == 6:
                        self.assertEqual("profile_generated", result["status"])
                    elif index == 7:
                        self.assertEqual(
                            selected_candidate["property_binding"],
                            resolved["resolved_binding"],
                        )
                    elif index == 8:
                        self.assertEqual(selected_candidate["source"], resolved["final_source"])
                    elif index == 9:
                        self.assertTrue(
                            any(
                                item["candidate_id"] == selected_id
                                for item in resolved["candidate_chain"]
                            )
                        )
                    else:
                        self.assertEqual(
                            selected["approval_id"],
                            result["final_profile"]["closure_evidence"][0][
                                "qa_decision_id"
                            ],
                        )
                elif index <= 15:
                    if index == 15:
                        self.assertEqual(
                            "awaiting_approval", unexpressible_keep_result["status"]
                        )
                        self.assertIsNone(unexpressible_keep_result["final_profile"])
                    else:
                        resolved = keep_result["final_profile"]["resolved_properties"][0]
                        if index == 11:
                            self.assertEqual("profile_generated", keep_result["status"])
                        elif index == 12:
                            self.assertEqual("preserve", resolved["execution_mode"])
                        elif index == 13:
                            self.assertEqual("preserve", resolved["resolved_binding"]["mode"])
                        else:
                            self.assertTrue(
                                any(
                                    item["source"] == resolved["final_source"]
                                    and item["property_binding"]["mode"] == "preserve"
                                    for item in resolved["candidate_chain"]
                                )
                            )
                elif index == 16:
                    resolved = exclude_winner_result["final_profile"]["resolved_properties"][0]
                    self.assertEqual("profile_generated", exclude_winner_result["status"])
                    self.assertEqual(losing_candidate["property_binding"], resolved["resolved_binding"])
                    self.assertEqual(losing_candidate["layer_kind"], resolved["final_layer_kind"])
                    self.assertEqual(losing_candidate["source"], resolved["final_source"])
                    self.assertEqual(
                        [losing_candidate["candidate_id"]],
                        [item["candidate_id"] for item in resolved["candidate_chain"]],
                    )
                    self.assertEqual(
                        proposed_candidate["candidate_id"],
                        resolved["excluded_candidates"][0]["candidate"]["candidate_id"],
                    )
                elif index == 17:
                    resolved = exclude_loser_result["final_profile"]["resolved_properties"][0]
                    self.assertEqual("profile_generated", exclude_loser_result["status"])
                    self.assertEqual(proposal["proposed_binding"], resolved["resolved_binding"])
                    self.assertEqual(proposal["final_source"], resolved["final_source"])
                    self.assertEqual(
                        losing_candidate["candidate_id"],
                        resolved["excluded_candidates"][0]["candidate"]["candidate_id"],
                    )
                elif index == 18:
                    resolved = redundant_result["final_profile"]["resolved_properties"][0]
                    self.assertEqual("profile_generated", redundant_result["status"])
                    self.assertEqual("12.00", resolved["resolved_binding"]["value"]["value"])
                    self.assertEqual(redundant_remaining["layer_kind"], resolved["final_layer_kind"])
                    self.assertEqual(redundant_remaining["source"], resolved["final_source"])
                    self.assertEqual(
                        [redundant_remaining["candidate_id"]],
                        [item["candidate_id"] for item in resolved["candidate_chain"]],
                    )
                    self.assertEqual(
                        redundant_target["candidate_id"],
                        resolved["excluded_candidates"][0]["candidate"]["candidate_id"],
                    )
                    self.assertEqual(
                        "qa_exclusion",
                        resolved["excluded_candidates"][0]["exclusion_reason"],
                    )
                elif index == 19:
                    self.assertEqual("awaiting_approval", unresolved_result["status"])
                    self.assertIsNone(unresolved_result["final_profile"])
                elif index == 20:
                    self.assertEqual("awaiting_approval", weak_result["status"])
                    self.assertIsNone(weak_result["final_profile"])
                elif index == 21:
                    self.assertEqual("awaiting_approval", unknown_target_result["status"])
                    self.assertIsNone(unknown_target_result["final_profile"])
                    self.assertTrue(unknown_target_result["application_diagnostics"])
                elif index == 22:
                    self.assertEqual("awaiting_approval", foreign_target_result["status"])
                    self.assertIsNone(foreign_target_result["final_profile"])
                    self.assertTrue(foreign_target_result["application_diagnostics"])
                elif index == 23:
                    self.assertEqual("awaiting_approval", sole_exclusion_result["status"])
                    self.assertIsNone(sole_exclusion_result["final_profile"])
                elif index <= 26:
                    result = chain_result
                    if index == 24:
                        self.assertEqual("profile_generated", result["status"])
                    elif index == 25:
                        self.assertEqual(
                            [second["semantic_fingerprint"]],
                            result["final_profile"]["bindings"]["approval_fingerprints"],
                        )
                    else:
                        self.assertEqual(
                            second["approval_id"],
                            result["final_profile"]["closure_evidence"][0][
                                "qa_decision_id"
                            ],
                        )
                elif index <= 28:
                    result = branch_result if index == 27 else cycle_result
                    self.assertEqual("awaiting_approval", result["status"])
                elif index == 29:
                    result = unknown_result
                    self.assertEqual("awaiting_approval", result["status"])
                else:
                    result = stale_result
                    self.assertEqual("awaiting_approval", result["status"])
        self.assertEqual(report_snapshot, report)
        self.assertEqual(adopt_snapshot, adopt)
        self.assertEqual(selected_snapshot, selected)

    def test_final_contract_final_001_030(self) -> None:
        report, manifest = compose(self.registry, [asset(self.registry, "final_base")])
        report_snapshot = deepcopy(report)
        result = apply(self.registry, report, [])
        final = result["final_profile"]
        dag = _validate_artifact_dag_for_test(
            [final, report, manifest, asset(self.registry, "final_base")],
            registry=self.registry,
        )
        second = apply(self.registry, deepcopy(report), [])
        for index in range(1, 31):
            assertion_id = f"T411-C-FINAL-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                self.assertEqual("profile_generated", result["status"])
                self.assertEqual(report_snapshot, report)
                self.assertEqual(report_snapshot, result["report"])
                self.assertEqual("disabled", final["activation"])
                self.assertFalse(final["legacy_input"])
                self.assertFalse(final["final_ready_eligible"])
                self.assertFalse(final["delivery_allowed"])
                self.assertEqual([], final["bindings"]["approval_fingerprints"])
                self.assertEqual([], final["closure_evidence"])
                self.assertEqual(report["semantic_fingerprint"], final["bindings"]["composition_report_fingerprint"])
                self.assertEqual(len(report["proposed_resolutions"]), len(final["resolved_properties"]))
                self.assert_valid(final)
                if index <= 10:
                    self.assertFalse(dag.runtime_eligible)
                elif index <= 20:
                    self.assertEqual(final["semantic_fingerprint"], second["final_profile"]["semantic_fingerprint"])
                elif index <= 25:
                    self.assertFalse(any(item["safety_check"]["status"] != "pass" for item in final["resolved_properties"]))
                else:
                    self.assertNotIn("artifact_kind", result)
                    self.assertNotIn("schema_version", result)
                    self.assertNotIn("semantic_fingerprint", result)
                    self.assertNotIn("input_fingerprints", result)

    def test_deterministic_permutations_and_independent_expected_values(self) -> None:
        assets = [
            asset(self.registry, "perm_base", layer="monograph_base", value="10.00"),
            asset(self.registry, "perm_type", layer="book_type", value="11.00"),
            asset(self.registry, "perm_task", layer="task_override", value="12.00"),
        ]
        expected_layer = "task_override"
        expected_value = "12.00"
        expected_fingerprint = "sha256:a53e95af6e9922119e07be000528adb50e1a5f3176ada6b1468e7ab334fc2454"
        fingerprints = set()
        for seed in range(6):
            shuffled = deepcopy(assets)
            random.Random(seed).shuffle(shuffled)
            report, _ = compose(self.registry, shuffled)
            fingerprints.add(report["semantic_fingerprint"])
            self.assertEqual(expected_layer, report["proposed_resolutions"][0]["final_layer_kind"])
            self.assertEqual(expected_value, report["proposed_resolutions"][0]["proposed_binding"]["value"]["value"])
        self.assertEqual(1, len(fingerprints))
        self.assertEqual({expected_fingerprint}, fingerprints)

    def test_no_runtime_import_or_public_cli(self) -> None:
        runtime = (ROOT / "format-monograph" / "scripts" / "run_monograph.py").read_text(encoding="utf-8")
        skill = (ROOT / "format-monograph" / "SKILL.md").read_text(encoding="utf-8")
        self.assertNotIn("profile_v2_composer", runtime)
        self.assertNotIn("profile_v2_composer", skill)
        module = ROOT / "format-monograph" / "scripts" / "profile_v2_composer.py"
        source = module.read_text(encoding="utf-8")
        self.assertNotIn("argparse", source)
        for private_test_adapter in (
            "_schema_documents",
            "_test_schema_overrides",
            "_validate_artifact_for_test",
            "_stamp_semantic_fingerprint_for_test",
            "_compose_profile_for_test",
            "_apply_resolutions_for_test",
        ):
            self.assertNotIn(private_test_adapter, source)
        self.assertIsNotNone(importlib.util.spec_from_file_location("p2b_c_candidate", module))


if __name__ == "__main__":
    unittest.main()
