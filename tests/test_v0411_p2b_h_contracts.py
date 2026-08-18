from __future__ import annotations

import hashlib
import inspect
import json
import sys
import unittest
from collections import Counter
from copy import deepcopy
from pathlib import Path
from unittest import mock
from urllib.parse import urldefrag

from jsonschema import Draft202012Validator


REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "format-monograph" / "scripts"
SCHEMAS = REPO / "format-monograph" / "references" / "schemas" / "v2"
FIXTURES = REPO / "tests" / "fixtures"
H_FIXTURES = FIXTURES / "v0411" / "p2b_h"
sys.path.insert(0, str(SCRIPTS))

import profile_v2_artifacts as artifacts  # noqa: E402
from profile_v2_artifacts import (  # noqa: E402
    ArtifactContractError,
    ArtifactDagError,
    ArtifactRouteError,
    _contract_routes_from_matrix,
    _is_obvious_placeholder_fingerprint,
    _route_artifact_contract_from_routes,
    _schema_documents,
    _test_schema_overrides,
    _topological_artifact_order,
    _validate_artifact_dag_for_test,
    _validate_artifact_for_test,
    load_artifact_contract_matrix,
    load_artifact_schema,
    profile_v2_composer_contract_enabled,
    route_artifact_contract,
    schema_documents,
    schema_inventory_contract,
    validate_artifact,
    verify_contract_matrix_alignment,
)
from profile_v2_authority import (  # noqa: E402
    AuthorityContractError,
    authority_contract_fingerprint,
    authority_layer_ids,
    authority_rank,
    build_authority_layer_schema,
    load_authority_contract,
    validate_authority_contract_document,
    verify_authority_projection,
    verify_legacy_layer_compatibility,
)
from profile_v2_canonical import (  # noqa: E402
    CanonicalizationError,
    _node_digest,
    _resolve_schema,
    _semantic_projection,
    _stamp_semantic_fingerprint_for_test,
    audit_schema_composition,
    compute_semantic_fingerprint,
    effective_fingerprint_exclusion,
    fingerprint_field_inventory,
    fingerprint_field_node,
    load_canonical_composition_policy,
    stamp_semantic_fingerprint,
)
from profile_v2_registry import load_registry  # noqa: E402
from profile_v2_scope import normalize_scope  # noqa: E402


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


EXPECTATIONS = load_json(H_FIXTURES / "contract-expectations.json")
MINIMAL_ARTIFACTS = load_json(FIXTURES / "v041" / "minimal-artifacts.json")


def digest(label: str) -> str:
    return "sha256:" + hashlib.sha256(label.encode("utf-8")).hexdigest()


def test_registry() -> dict:
    return load_registry(
        FIXTURES / "v041" / "property-registry.v2.1.test.json",
        allow_test=True,
        version="2.1",
    )


def raw_document_scope(document_id: str = "document:synthetic") -> dict:
    return {
        "selectors": [
            {"selector_kind": "document", "selector_ids": [document_id]}
        ],
        "exclusions": [],
        "mutually_exclusive_conditions": [],
    }


def document_scope(document_id: str = "document:synthetic") -> dict:
    return normalize_scope(raw_document_scope(document_id))


def input_fingerprint(input_id: str, role: str, fingerprint: str) -> dict:
    return {"input_id": input_id, "role": role, "fingerprint": fingerprint}


def stamp_test(document: dict, registry: dict | None = None) -> dict:
    registry = registry or test_registry()
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


def layered_asset() -> dict:
    scope = raw_document_scope()
    return stamp_test(
        {
            "artifact_kind": "layered-rule-asset",
            "schema_version": "2.1",
            "artifact_id": "layered-rule-asset:synthetic-h",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                input_fingerprint("input:rule-source", "rule_asset", digest("rule-source"))
            ],
            "semantic_fingerprint": digest("unstamped-rule"),
            "layer_kind": "monograph_base",
            "can_override_safety_invariants": False,
            "activation": "disabled",
            "asset_scope": scope,
            "allowed_semantic_object_kinds": ["paragraph"],
            "rules": [
                {
                    "rule_id": "RULE-H-001",
                    "semantic_object_kind": "paragraph",
                    "scope": scope,
                    "confidence": "high",
                    "status": "draft",
                    "properties": [
                        {
                            "property_id": "test.paragraph-font-size",
                            "value": {"type": "decimal", "value": "10.50"},
                            "unit_id": "unit.pt",
                            "mode": "report",
                        }
                    ],
                }
            ],
        }
    )


def feature_manifest(*, composer: bool = False, schema_enabled: bool = True) -> dict:
    return stamp_test(
        {
            "artifact_kind": "feature-activation-manifest",
            "schema_version": "2.1",
            "registry_contract_version": "2.1",
            "authority_contract_version": "1.0",
            "artifact_id": "feature-activation-manifest:synthetic-h",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                input_fingerprint(
                    "input:feature-source",
                    "source_document",
                    digest("feature-source"),
                )
            ],
            "semantic_fingerprint": digest("unstamped-feature"),
            "features": {
                "profile_v2_schema": schema_enabled,
                "profile_v2_composer": composer,
                "final_ready_eligible": False,
            },
        }
    )


def property_key(scope: dict | None = None) -> dict:
    return {
        "semantic_object_kind": "paragraph",
        "property_id": "test.paragraph-font-size",
        "normalized_scope": scope or document_scope(),
    }


def property_binding() -> dict:
    return {
        "property_id": "test.paragraph-font-size",
        "value": {"type": "decimal", "value": "10.50"},
        "unit_id": "unit.pt",
        "mode": "report",
    }


def candidate() -> dict:
    return {
        "candidate_id": "candidate:h-001",
        "property_binding": property_binding(),
        "source": {
            "source_artifact_id": "layered-rule-asset:synthetic-h",
            "source_rule_id": "RULE-H-001",
        },
        "layer_kind": "monograph_base",
        "confidence": "high",
        "scope_status": "applicable",
    }


def report_bindings(feature: dict, rule: dict) -> dict:
    return {
        "input_fingerprint": digest("source-document"),
        "feature_activation_fingerprint": feature["semantic_fingerprint"],
        "property_registry_fingerprint": digest("test-registry"),
        "authority_contract_fingerprint": authority_contract_fingerprint("1.0"),
        "rule_asset_fingerprints": [rule["semantic_fingerprint"]],
        "structure_fingerprint": digest("structure-map"),
    }


def report_inputs(bindings: dict) -> list[dict]:
    return [
        input_fingerprint("input:source", "source_document", bindings["input_fingerprint"]),
        input_fingerprint("input:feature", "feature_activation", bindings["feature_activation_fingerprint"]),
        input_fingerprint("input:registry", "property_registry", bindings["property_registry_fingerprint"]),
        input_fingerprint("input:authority", "authority_contract", bindings["authority_contract_fingerprint"]),
        input_fingerprint("input:rule", "rule_asset", bindings["rule_asset_fingerprints"][0]),
        input_fingerprint("input:structure", "structure", bindings["structure_fingerprint"]),
    ]


def composition_report(
    feature: dict,
    rule: dict,
    *,
    status: str = "resolvable",
) -> dict:
    bindings = report_bindings(feature, rule)
    key = property_key()
    item = candidate()
    has_proposal = status in {"resolvable", "awaiting_approval"}
    has_approval = status == "awaiting_approval"
    fatal = status == "fatal"
    unresolvable = status == "unresolvable"
    document = {
        "artifact_kind": "conflict-report",
        "schema_version": "2.2",
        "registry_contract_version": "2.1",
        "authority_contract_version": "1.0",
        "artifact_id": f"conflict-report:synthetic-{status}",
        "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
        "input_fingerprints": report_inputs(bindings),
        "semantic_fingerprint": digest(f"unstamped-report-{status}"),
        "generated_at": "2026-08-17T00:00:00Z",
        "bindings": bindings,
        "candidate_groups": (
            [
                {
                    "candidate_group_id": "candidate-group:h-001",
                    "key": key,
                    "candidates": [item],
                    "excluded_candidates": [],
                }
            ]
            if has_proposal
            else []
        ),
        "scope_partitions": [],
        "proposed_resolutions": (
            [
                {
                    "proposed_resolution_id": "proposed-resolution:h-001",
                    "key": key,
                    "proposed_binding": property_binding(),
                    "final_layer_kind": "monograph_base",
                    "final_source": item["source"],
                    "candidate_chain": [item],
                    "override_chain": [item["candidate_id"]],
                    "confidence": "high",
                    "execution_mode": "report",
                }
            ]
            if has_proposal
            else []
        ),
        "fatal_diagnostics": (
            [
                {
                    "diagnostic_id": "fatal:h-001",
                    "category": "invalid_binding",
                    "reason_code": "FATAL-BINDING",
                }
            ]
            if fatal
            else []
        ),
        "unresolvable_blockers": (
            [
                {
                    "blocker_id": "blocker:h-001",
                    "category": "unknown_overlap",
                    "key": key,
                    "reason_code": "UNKNOWN-OVERLAP",
                }
            ]
            if unresolvable
            else []
        ),
        "approval_required_conflicts": (
            [
                {
                    "conflict_id": "conflict:h-001",
                    "proposed_resolution_id": "proposed-resolution:h-001",
                    "key": key,
                    "candidates": [item],
                    "allowed_decisions": ["adopt_proposed", "select_candidate", "keep_original"],
                }
            ]
            if has_approval
            else []
        ),
        "diagnostics": [],
        "proposal_status": status,
    }
    return stamp_test(document)


def approval(report: dict, *, decision: str = "select_candidate") -> dict:
    conflict = report["approval_required_conflicts"][0]
    bindings = {
        "input_fingerprint": report["bindings"]["input_fingerprint"],
        "structure_fingerprint": report["bindings"]["structure_fingerprint"],
        "composition_report_fingerprint": report["semantic_fingerprint"],
    }
    target = {
        "conflict_id": conflict["conflict_id"],
        "proposed_resolution_id": conflict["proposed_resolution_id"],
        "normalized_scope": conflict["key"]["normalized_scope"],
    }
    if decision in {"select_candidate", "exclude_candidate"}:
        target["candidate_id"] = conflict["candidates"][0]["candidate_id"]
    return stamp_test(
        {
            "artifact_kind": "qa-approval-artifact",
            "schema_version": "2.1",
            "registry_contract_version": "2.1",
            "authority_contract_version": "1.0",
            "artifact_id": "qa-approval-artifact:synthetic-h",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": [
                input_fingerprint("input:source", "source_document", bindings["input_fingerprint"]),
                input_fingerprint("input:structure", "structure", bindings["structure_fingerprint"]),
                input_fingerprint("input:report", "conflict_report", bindings["composition_report_fingerprint"]),
            ],
            "semantic_fingerprint": digest("unstamped-approval"),
            "approval_id": "approval:h-001",
            "approver": {"actor_id": "actor:user", "actor_role": "user"},
            "decision_type": (
                "qa_exclusion"
                if decision == "exclude_candidate"
                else "keep_original"
                if decision == "keep_original"
                else "conflict_resolution"
            ),
            "decision": decision,
            "reason": "Synthetic approval for contract testing.",
            "created_at": "2026-08-17T00:00:00Z",
            "bindings": bindings,
            "target": target,
            "previous_approval_id": None,
        }
    )


def final_profile(
    report: dict,
    feature: dict,
    rule: dict,
    approval_document: dict | None = None,
) -> dict:
    approval_fingerprints = (
        [approval_document["semantic_fingerprint"]] if approval_document else []
    )
    bindings = {
        "task_fingerprint": digest("task"),
        "input_fingerprint": report["bindings"]["input_fingerprint"],
        "feature_activation_fingerprint": feature["semantic_fingerprint"],
        "property_registry_fingerprint": report["bindings"]["property_registry_fingerprint"],
        "authority_contract_fingerprint": report["bindings"]["authority_contract_fingerprint"],
        "rule_asset_fingerprints": [rule["semantic_fingerprint"]],
        "structure_fingerprint": report["bindings"]["structure_fingerprint"],
        "approval_fingerprints": approval_fingerprints,
        "composition_report_fingerprint": report["semantic_fingerprint"],
    }
    inputs = [
        input_fingerprint("input:task", "task", bindings["task_fingerprint"]),
        input_fingerprint("input:source", "source_document", bindings["input_fingerprint"]),
        input_fingerprint("input:feature", "feature_activation", bindings["feature_activation_fingerprint"]),
        input_fingerprint("input:registry", "property_registry", bindings["property_registry_fingerprint"]),
        input_fingerprint("input:authority", "authority_contract", bindings["authority_contract_fingerprint"]),
        input_fingerprint("input:rule", "rule_asset", bindings["rule_asset_fingerprints"][0]),
        input_fingerprint("input:structure", "structure", bindings["structure_fingerprint"]),
        input_fingerprint("input:report", "conflict_report", bindings["composition_report_fingerprint"]),
    ]
    if approval_document:
        inputs.append(
            input_fingerprint(
                "input:approval",
                "approval",
                approval_document["semantic_fingerprint"],
            )
        )
    closure = []
    if approval_document:
        conflict = report["approval_required_conflicts"][0]
        closure = [
            {
                "conflict_id": conflict["conflict_id"],
                "proposed_resolution_id": conflict["proposed_resolution_id"],
                "qa_decision_id": approval_document["approval_id"],
            }
        ]
    resolved_properties = []
    for proposal in report["proposed_resolutions"]:
        resolved_properties.append(
            {
                "resolution_id": proposal["proposed_resolution_id"],
                "key": deepcopy(proposal["key"]),
                "resolved_binding": deepcopy(proposal["proposed_binding"]),
                "final_layer_kind": proposal["final_layer_kind"],
                "final_source": deepcopy(proposal["final_source"]),
                "candidate_chain": deepcopy(proposal["candidate_chain"]),
                "override_chain": list(proposal["override_chain"]),
                "excluded_candidates": [],
                "confidence": proposal["confidence"],
                "safety_check": {
                    "status": "pass",
                    "checked_invariant_ids": [
                        "test.safety-author-content-immutable"
                    ],
                },
                "execution_mode": proposal["execution_mode"],
            }
        )
    return stamp_test(
        {
            "artifact_kind": "final-execution-profile",
            "schema_version": "2.2",
            "registry_contract_version": "2.1",
            "authority_contract_version": "1.0",
            "artifact_id": "final-execution-profile:synthetic-h",
            "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
            "input_fingerprints": inputs,
            "semantic_fingerprint": digest("unstamped-final"),
            "task_id": "task:synthetic-h",
            "legacy_input": False,
            "activation": "disabled",
            "final_ready_eligible": False,
            "delivery_allowed": False,
            "safety_invariants": {
                "overridable": False,
                "author_content_mutation_allowed": False,
                "legacy_auto_activation_allowed": False,
                "delivery_evidence_allowed": False,
                "final_ready_allowed": False,
            },
            "bindings": bindings,
            "resolved_properties": resolved_properties,
            "closure_evidence": closure,
        }
    )


class P2bHContractTests(unittest.TestCase):
    maxDiff = None

    def setUp(self) -> None:
        self.registry = test_registry()
        self.rule = layered_asset()
        self.feature = feature_manifest(composer=True)

    def validate_test(self, document: dict) -> None:
        _validate_artifact_for_test(
            document,
            registry=self.registry,
            features={"profile_v2_schema": True},
        )

    def test_v0411_h_dag_001_012(self) -> None:
        with self.subTest(assertion_id="T411-H-DAG-001"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            result = _validate_artifact_dag_for_test(
                [final, report, self.feature, self.rule], registry=self.registry
            )
            self.assertFalse(result.runtime_eligible)
            self.assertLess(
                result.topological_order.index(report["artifact_id"]),
                result.topological_order.index(final["artifact_id"]),
            )

        with self.subTest(assertion_id="T411-H-DAG-002"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            qa = approval(report)
            final = final_profile(report, self.feature, self.rule, qa)
            result = _validate_artifact_dag_for_test(
                [final, qa, report, self.feature, self.rule], registry=self.registry
            )
            self.assertEqual(5, result.artifact_count)

        with self.subTest(assertion_id="T411-H-DAG-003"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            final = final_profile(report, self.feature, self.rule)
            with self.assertRaisesRegex(ArtifactDagError, "close every"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-004"):
            report = composition_report(self.feature, self.rule, status="fatal")
            final = final_profile(report, self.feature, self.rule)
            with self.assertRaisesRegex(ArtifactDagError, "fatal diagnostics"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )
            qa = approval(composition_report(self.feature, self.rule, status="awaiting_approval"))
            qa["bindings"]["composition_report_fingerprint"] = report["semantic_fingerprint"]
            next(item for item in qa["input_fingerprints"] if item["role"] == "conflict_report")["fingerprint"] = report["semantic_fingerprint"]
            qa["target"]["conflict_id"] = report["fatal_diagnostics"][0]["diagnostic_id"]
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "Fatal diagnostics cannot accept"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-005"):
            report = composition_report(self.feature, self.rule, status="unresolvable")
            final = final_profile(report, self.feature, self.rule)
            with self.assertRaisesRegex(ArtifactDagError, "unresolvable blockers"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )
            qa = approval(composition_report(self.feature, self.rule, status="awaiting_approval"))
            qa["bindings"]["composition_report_fingerprint"] = report["semantic_fingerprint"]
            next(item for item in qa["input_fingerprints"] if item["role"] == "conflict_report")["fingerprint"] = report["semantic_fingerprint"]
            qa["target"]["conflict_id"] = report["unresolvable_blockers"][0]["blocker_id"]
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "Unresolvable blockers cannot accept"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        report = composition_report(self.feature, self.rule, status="awaiting_approval")
        with self.subTest(assertion_id="T411-H-DAG-006"):
            qa = approval(report)
            qa["target"]["conflict_id"] = "conflict:unknown"
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "unknown conflict_id"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-007"):
            qa = approval(report)
            qa["target"]["proposed_resolution_id"] = "proposed-resolution:unknown"
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "unknown proposed_resolution_id"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-008"):
            qa = approval(report)
            qa["target"]["candidate_id"] = "candidate:unknown"
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "not part of the conflict"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-009"):
            qa = approval(report)
            qa["target"]["normalized_scope"] = document_scope("document:expanded")
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "scope differs"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-010"):
            qa = approval(report)
            stale = digest("stale-report")
            qa["bindings"]["composition_report_fingerprint"] = stale
            next(
                item for item in qa["input_fingerprints"] if item["role"] == "conflict_report"
            )["fingerprint"] = stale
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "missing internal conflict-report"):
                _validate_artifact_dag_for_test(
                    [qa, report, self.feature, self.rule], registry=self.registry
                )
            bad = feature_manifest()
            bad["semantic_fingerprint"] = "sha256:not-a-digest"
            with self.assertRaises(ArtifactContractError):
                self.validate_test(bad)
            self.assertTrue(_is_obvious_placeholder_fingerprint("sha256:" + ("0" * 64)))
            self.assertFalse(_is_obvious_placeholder_fingerprint(digest("not-placeholder")))

        with self.subTest(assertion_id="T411-H-DAG-011"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            final["bindings"]["structure_fingerprint"] = digest("other-structure")
            next(
                item for item in final["input_fingerprints"] if item["role"] == "structure"
            )["fingerprint"] = final["bindings"]["structure_fingerprint"]
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactDagError, "differs from its composition report"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-DAG-012"):
            with self.assertRaisesRegex(ArtifactDagError, "cycle"):
                _topological_artifact_order({"a": {"b"}, "b": {"a"}})

    def test_v0411_h_reference_integrity_001_024(self) -> None:
        with self.subTest(assertion_id="T411-H-REF-001"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            invalid_pairs = (
                ("keep_original", "select_candidate"),
                ("qa_exclusion", "adopt_proposed"),
                ("conflict_resolution", "keep_original"),
            )
            for decision_type, decision in invalid_pairs:
                qa = approval(report)
                qa["decision_type"] = decision_type
                qa["decision"] = decision
                qa = stamp_test(qa)
                with self.assertRaisesRegex(ArtifactContractError, "inconsistent"):
                    self.validate_test(qa)

        with self.subTest(assertion_id="T411-H-REF-002"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            report["approval_required_conflicts"][0]["key"] = property_key(
                document_scope("document:other")
            )
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "different normalized keys"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-003"):
            report = composition_report(self.feature, self.rule)
            excluded = deepcopy(candidate())
            excluded["candidate_id"] = "candidate:h-excluded"
            report["candidate_groups"][0]["excluded_candidates"].append(
                {
                    "candidate": excluded,
                    "exclusion_reason": "out_of_scope",
                    "reason_code": "OUT-OF-SCOPE",
                }
            )
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "cannot be applicable"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-004"):
            report = composition_report(self.feature, self.rule)
            report["scope_partitions"] = [
                {
                    "partition_id": "partition:h-failed",
                    "key": property_key(),
                    "source_scope": document_scope(),
                    "partition_scopes": [document_scope()],
                    "evidence_status": "failed",
                }
            ]
            report = stamp_test(report)
            with self.assertRaisesRegex(
                ArtifactContractError, "Failed or not-evaluated scope partitions"
            ):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-005"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            final["resolved_properties"] = []
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactDagError, "every and only"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-006"):
            report = composition_report(self.feature, self.rule)
            duplicate = deepcopy(report["candidate_groups"][0])
            duplicate["candidate_group_id"] = "candidate-group:h-duplicate"
            report["candidate_groups"].append(duplicate)
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "candidate_groups repeats"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-007"):
            report = composition_report(self.feature, self.rule)
            partition = {
                "partition_id": "partition:h-001",
                "key": property_key(),
                "source_scope": document_scope(),
                "partition_scopes": [document_scope()],
                "evidence_status": "verified",
            }
            duplicate = deepcopy(partition)
            duplicate["partition_id"] = "partition:h-002"
            report["scope_partitions"] = [partition, duplicate]
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "scope_partitions repeats"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-008"):
            report = composition_report(self.feature, self.rule)
            duplicate = deepcopy(report["proposed_resolutions"][0])
            duplicate["proposed_resolution_id"] = "proposed-resolution:h-duplicate"
            report["proposed_resolutions"].append(duplicate)
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "proposed_resolutions repeats"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-009"):
            report = composition_report(self.feature, self.rule)
            report["candidate_groups"] = []
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "no candidate group"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-010"):
            report = composition_report(self.feature, self.rule)
            extra = deepcopy(candidate())
            extra["candidate_id"] = "candidate:h-untracked"
            report["proposed_resolutions"][0]["candidate_chain"].append(extra)
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "does not exactly reference"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-011"):
            report = composition_report(self.feature, self.rule)
            report["proposed_resolutions"][0]["override_chain"].append(
                "candidate:h-unknown"
            )
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "non-active"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-012"):
            report = composition_report(self.feature, self.rule)
            report["proposed_resolutions"][0]["final_source"] = {
                "source_artifact_id": "layered-rule-asset:other",
                "source_rule_id": "RULE-OTHER",
            }
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "absent from its candidate chain"):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-013"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            final["resolved_properties"][0]["key"] = property_key(
                document_scope("document:other")
            )
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactDagError, "key differs"):
                _validate_artifact_dag_for_test(
                    [final, report, self.feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-014"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            excluded = deepcopy(candidate())
            excluded["candidate_id"] = "candidate:h-final-excluded"
            final["resolved_properties"][0]["excluded_candidates"].append(
                {
                    "candidate": excluded,
                    "exclusion_reason": "out_of_scope",
                    "reason_code": "OUT-OF-SCOPE",
                }
            )
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactContractError, "cannot be applicable"):
                self.validate_test(final)

        with self.subTest(assertion_id="T411-H-REF-015"):
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            final["resolved_properties"][0]["override_chain"].append(
                "candidate:h-unknown"
            )
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactContractError, "unknown active"):
                self.validate_test(final)

        with self.subTest(assertion_id="T411-H-REF-016"):
            report = composition_report(self.feature, self.rule)
            report["scope_partitions"] = [
                {
                    "partition_id": "partition:h-pending",
                    "key": property_key(),
                    "source_scope": document_scope(),
                    "partition_scopes": [document_scope()],
                    "evidence_status": "not_evaluated",
                }
            ]
            report = stamp_test(report)
            with self.assertRaisesRegex(
                ArtifactContractError, "Failed or not-evaluated scope partitions"
            ):
                self.validate_test(report)

        with self.subTest(assertion_id="T411-H-REF-017"):
            old_feature = deepcopy(MINIMAL_ARTIFACTS["feature-activation-manifest"])
            report = composition_report(self.feature, self.rule)
            report["bindings"]["feature_activation_fingerprint"] = old_feature[
                "semantic_fingerprint"
            ]
            next(
                item
                for item in report["input_fingerprints"]
                if item["role"] == "feature_activation"
            )["fingerprint"] = old_feature["semantic_fingerprint"]
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactDagError, "schema 2.1"):
                _validate_artifact_dag_for_test(
                    [report, old_feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-018"):
            disabled_feature = feature_manifest(composer=False)
            report = composition_report(disabled_feature, self.rule)
            with self.assertRaisesRegex(ArtifactDagError, "profile_v2_composer=true"):
                _validate_artifact_dag_for_test(
                    [report, disabled_feature, self.rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-019"):
            old_rule = deepcopy(MINIMAL_ARTIFACTS["layered-rule-asset"])
            report = composition_report(self.feature, self.rule)
            report["bindings"]["rule_asset_fingerprints"] = [
                old_rule["semantic_fingerprint"]
            ]
            next(
                item
                for item in report["input_fingerprints"]
                if item["role"] == "rule_asset"
            )["fingerprint"] = old_rule["semantic_fingerprint"]
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactDagError, "schema 2.1"):
                _validate_artifact_dag_for_test(
                    [report, self.feature, old_rule], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-020"):
            old_report = deepcopy(MINIMAL_ARTIFACTS["conflict-report"])
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            qa = approval(report)
            qa["bindings"]["composition_report_fingerprint"] = old_report[
                "semantic_fingerprint"
            ]
            next(
                item
                for item in qa["input_fingerprints"]
                if item["role"] == "conflict_report"
            )["fingerprint"] = old_report["semantic_fingerprint"]
            qa = stamp_test(qa)
            with self.assertRaisesRegex(ArtifactDagError, "schema 2.2"):
                _validate_artifact_dag_for_test(
                    [qa, old_report], registry=self.registry
                )

        with self.subTest(assertion_id="T411-H-REF-021"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            qa = approval(report)
            old_qa = deepcopy(MINIMAL_ARTIFACTS["qa-approval-artifact"])
            final = final_profile(report, self.feature, self.rule, qa)
            final["bindings"]["approval_fingerprints"] = [
                old_qa["semantic_fingerprint"]
            ]
            next(
                item
                for item in final["input_fingerprints"]
                if item["role"] == "approval"
            )["fingerprint"] = old_qa["semantic_fingerprint"]
            final = stamp_test(final)
            with self.assertRaisesRegex(ArtifactDagError, "schema 2.1"):
                _validate_artifact_dag_for_test(
                    [final, old_qa, report, self.feature, self.rule],
                    registry=self.registry,
                )

        with self.subTest(assertion_id="T411-H-REF-022"):
            report = composition_report(self.feature, self.rule, status="awaiting_approval")
            incomplete = deepcopy(report["approval_required_conflicts"][0])
            incomplete["candidates"][0]["candidate_id"] = "candidate:h-other"
            report["approval_required_conflicts"][0] = incomplete
            report = stamp_test(report)
            with self.assertRaisesRegex(ArtifactContractError, "complete candidate group"):
                self.validate_test(report)

        for assertion_id, evidence_status in (
            ("T411-H-REF-023", "failed"),
            ("T411-H-REF-024", "not_evaluated"),
        ):
            with self.subTest(assertion_id=assertion_id):
                report = composition_report(
                    self.feature, self.rule, status="unresolvable"
                )
                report["scope_partitions"] = [
                    {
                        "partition_id": f"partition:h-{evidence_status}",
                        "key": property_key(),
                        "source_scope": document_scope(),
                        "partition_scopes": [document_scope()],
                        "evidence_status": evidence_status,
                    }
                ]
                report = stamp_test(report)
                self.validate_test(report)
                final = final_profile(report, self.feature, self.rule)
                with self.assertRaisesRegex(
                    ArtifactDagError, "unresolvable blockers"
                ):
                    _validate_artifact_dag_for_test(
                        [final, report, self.feature, self.rule],
                        registry=self.registry,
                    )

    def test_v0411_h_route_001_018(self) -> None:
        matrix = load_artifact_contract_matrix()
        with self.subTest(assertion_id="T411-H-ROUTE-001"):
            self.assertEqual(EXPECTATIONS["route_count"], len(matrix["routes"]))
        with self.subTest(assertion_id="T411-H-ROUTE-002"):
            for key, expected in EXPECTATIONS["artifact_routes"].items():
                kind, version = key.split("@")
                route = next(
                    item
                    for item in matrix["routes"]
                    if item["artifact_kind"] == kind and item["schema_version"] == version
                )
                self.assertEqual(expected, [route["registry_contract_version"], route["authority_contract_version"]])
        with self.subTest(assertion_id="T411-H-ROUTE-003"):
            with mock.patch.object(artifacts, "load_registry") as loader:
                with self.assertRaises(ArtifactRouteError):
                    route_artifact_contract({"artifact_kind": "conflict-report", "schema_version": "9.9"})
                loader.assert_not_called()
        with self.subTest(assertion_id="T411-H-ROUTE-004"):
            document = feature_manifest()
            del document["registry_contract_version"]
            with self.assertRaisesRegex(ArtifactRouteError, "explicit"):
                route_artifact_contract(document)
        with self.subTest(assertion_id="T411-H-ROUTE-005"):
            document = feature_manifest()
            document["registry_contract_version"] = "2.0"
            with self.assertRaisesRegex(ArtifactRouteError, "four-part"):
                route_artifact_contract(document)
        with self.subTest(assertion_id="T411-H-ROUTE-006"):
            document = feature_manifest()
            document["authority_contract_version"] = "none"
            with self.assertRaisesRegex(ArtifactRouteError, "four-part"):
                route_artifact_contract(document)
        with self.subTest(assertion_id="T411-H-ROUTE-007"):
            legacy = {"artifact_kind": "feature-activation-manifest", "schema_version": "2.0", "registry_contract_version": "2.0"}
            with self.assertRaisesRegex(ArtifactRouteError, "cannot self-declare"):
                route_artifact_contract(legacy)
        with self.subTest(assertion_id="T411-H-ROUTE-008"):
            with self.assertRaises(ArtifactRouteError):
                route_artifact_contract({"artifact_kind": "conflict-report", "schema_version": "3.0"})
        with self.subTest(assertion_id="T411-H-ROUTE-009"):
            with self.assertRaises(ArtifactRouteError):
                route_artifact_contract({"artifact_kind": "conflict-report", "schema_version": "2.9"})
        with self.subTest(assertion_id="T411-H-ROUTE-010"):
            report = composition_report(self.feature, self.rule)
            route = route_artifact_contract(report)
            self.assertEqual("2.1", route.registry_contract_version)
            self.assertNotEqual("2.2", route.registry_contract_version)
            with mock.patch.object(
                artifacts,
                "load_registry",
                wraps=artifacts.load_registry,
            ) as loader:
                artifacts.load_routed_contracts(report)
                loader.assert_called_once_with(version="2.1")
        with self.subTest(assertion_id="T411-H-ROUTE-011"):
            verify_contract_matrix_alignment()
        with self.subTest(assertion_id="T411-H-ROUTE-012"):
            drift = dict(artifacts.ARTIFACT_SCHEMA_FILES)
            drift[("conflict-report", "2.2", "2.1", "1.0")] = "wrong.schema.json"
            with self.assertRaisesRegex(ArtifactRouteError, "differs"):
                verify_contract_matrix_alignment(drift)
        with self.subTest(assertion_id="T411-H-ROUTE-013"):
            drift = dict(artifacts.ARTIFACT_SCHEMA_FILES)
            drift[("capability-snapshot", "2.1", "2.1", "none")] = "capability-snapshot.schema.json"
            with self.assertRaisesRegex(ArtifactRouteError, "differs"):
                verify_contract_matrix_alignment(drift)
        with self.subTest(assertion_id="T411-H-ROUTE-014"):
            self.assertEqual(EXPECTATIONS["schema_document_count"], len(schema_documents()))
            for schema_id in schema_documents():
                schema_inventory_contract(schema_id)
        with self.subTest(assertion_id="T411-H-ROUTE-015"):
            for route in matrix["routes"]:
                schema = load_artifact_schema(route["artifact_kind"], version=route["schema_version"])
                self.assertEqual(route["schema_version"], schema["properties"]["schema_version"]["const"])
                Draft202012Validator.check_schema(schema)
            awaiting = composition_report(self.feature, self.rule, status="awaiting_approval")
            qa = approval(awaiting)
            report = composition_report(self.feature, self.rule)
            final = final_profile(report, self.feature, self.rule)
            new_documents = [feature_manifest(), qa, report, final]
            rejected = 0
            offline = artifacts._offline_schema_registry(
                _test_schema_overrides(self.registry)
            )
            for source in new_documents:
                for target in new_documents:
                    target_schema = load_artifact_schema(
                        target["artifact_kind"], version=target["schema_version"]
                    )
                    validator = Draft202012Validator(target_schema, registry=offline)
                    if source["artifact_kind"] == target["artifact_kind"]:
                        self.assertTrue(validator.is_valid(source))
                    else:
                        self.assertFalse(validator.is_valid(source))
                        rejected += 1
            self.assertEqual(12, rejected)
            report_schema = load_artifact_schema("conflict-report", version="2.2")
            exclusion_schema = schema_documents()[
                "https://schemas.format-monograph.local/v2.2/common.schema.json"
            ]["$defs"]["excluded_candidate"]["properties"]["exclusion_reason"]
            self.assertEqual(
                ["out_of_scope", "module_out_of_scope", "qa_exclusion"],
                exclusion_schema["enum"],
            )
            self.assertNotIn("invalid_binding", exclusion_schema["enum"])
            self.assertNotIn("lower_precedence", exclusion_schema["enum"])
            self.assertEqual(
                ["fatal", "unresolvable", "awaiting_approval", "resolvable"],
                report_schema["properties"]["proposal_status"]["enum"],
            )
        with self.subTest(assertion_id="T411-H-ROUTE-016"):
            self.assertEqual(("2.2", "2.1", True), schema_inventory_contract("https://schemas.format-monograph.local/v2.2/common.schema.json"))
        with self.subTest(assertion_id="T411-H-ROUTE-017"):
            signature = inspect.signature(validate_artifact)
            self.assertNotIn("registry", signature.parameters)
            self.assertNotIn("schema_override", signature.parameters)
        with self.subTest(assertion_id="T411-H-ROUTE-018"):
            legacy_routes = [item for item in matrix["routes"] if item["version_source"] == "legacy_matrix"]
            self.assertTrue(legacy_routes)
            self.assertTrue(all(item["authority_contract_version"] == "none" for item in legacy_routes))

    def test_v0411_h_route_019_023_four_part_identity(self) -> None:
        matrix = load_artifact_contract_matrix()
        alternate = deepcopy(
            next(
                item
                for item in matrix["routes"]
                if item["artifact_kind"] == "feature-activation-manifest"
                and item["schema_version"] == "2.1"
            )
        )
        alternate["route_id"] = "route:feature-activation-manifest:2.1:registry-2.0"
        alternate["registry_contract_version"] = "2.0"
        matrix["routes"].append(alternate)
        routes = _contract_routes_from_matrix(matrix)
        with self.subTest(assertion_id="T411-H-ROUTE-019"):
            document = feature_manifest()
            document["registry_contract_version"] = "2.0"
            selected = _route_artifact_contract_from_routes(document, routes)
            self.assertEqual(alternate["route_id"], selected.route_id)
        with self.subTest(assertion_id="T411-H-ROUTE-020"):
            document = feature_manifest()
            document["registry_contract_version"] = "2.0"
            document["authority_contract_version"] = "none"
            with self.assertRaisesRegex(ArtifactRouteError, "four-part"):
                _route_artifact_contract_from_routes(document, routes)
        with self.subTest(assertion_id="T411-H-ROUTE-021"):
            document = feature_manifest()
            del document["registry_contract_version"]
            with self.assertRaisesRegex(ArtifactRouteError, "explicit"):
                _route_artifact_contract_from_routes(document, routes)
        with self.subTest(assertion_id="T411-H-ROUTE-022"):
            document = feature_manifest()
            document["registry_contract_version"] = "9.9"
            with self.assertRaisesRegex(ArtifactRouteError, "four-part"):
                _route_artifact_contract_from_routes(document, routes)
        with self.subTest(assertion_id="T411-H-ROUTE-023"):
            duplicate_legacy = deepcopy(
                next(
                    item
                    for item in matrix["routes"]
                    if item["artifact_kind"] == "capability-snapshot"
                    and item["schema_version"] == "2.0"
                )
            )
            duplicate_legacy["route_id"] = "route:capability-snapshot:2.0:duplicate"
            duplicate_legacy["registry_contract_version"] = "2.1"
            duplicate_matrix = deepcopy(matrix)
            duplicate_matrix["routes"].append(duplicate_legacy)
            with self.assertRaisesRegex(ArtifactRouteError, "legacy artifact/version"):
                _contract_routes_from_matrix(duplicate_matrix)

    def test_v0411_h_route_024_025_exact_schema_validation(self) -> None:
        document = feature_manifest()
        route = route_artifact_contract(document)
        exact_schema = load_artifact_schema(
            route.artifact_kind,
            version=route.schema_version,
            registry_contract_version=route.registry_contract_version,
            authority_contract_version=route.authority_contract_version,
        )

        with self.subTest(assertion_id="T411-H-ROUTE-024"):
            sentinel = object()
            with mock.patch.object(
                artifacts,
                "load_routed_contracts",
                return_value=(route, exact_schema, self.registry, {}),
            ), mock.patch.object(
                artifacts,
                "_validate_artifact_contract",
                return_value=sentinel,
            ) as validator:
                result = validate_artifact(
                    document, features={"profile_v2_schema": True}
                )
            self.assertIs(sentinel, result)
            self.assertIs(
                exact_schema, validator.call_args.kwargs["resolved_schema"]
            )

        with self.subTest(assertion_id="T411-H-ROUTE-025"):
            with mock.patch.object(
                artifacts,
                "load_artifact_schema",
                side_effect=AssertionError("pair-only schema lookup was used"),
            ):
                result = validate_artifact(
                    document, features={"profile_v2_schema": True}
                )
            self.assertFalse(result.runtime_eligible)

    def test_v0411_h_authority_001_012(self) -> None:
        expected = tuple(EXPECTATIONS["authority_layers"])
        contract = load_authority_contract()
        with self.subTest(assertion_id="T411-H-AUTH-001"):
            self.assertEqual(expected, authority_layer_ids())
        with self.subTest(assertion_id="T411-H-AUTH-002"):
            self.assertEqual(tuple(range(6)), tuple(authority_rank(item) for item in expected))
        with self.subTest(assertion_id="T411-H-AUTH-003"):
            self.assertEqual("independent_non_overridable", contract["safety_policy"])
            self.assertNotIn("safety", expected)
        with self.subTest(assertion_id="T411-H-AUTH-004"):
            verify_authority_projection()
            self.assertEqual(list(expected), build_authority_layer_schema()["enum"])
        with self.subTest(assertion_id="T411-H-AUTH-005"):
            changed = deepcopy(contract)
            changed["layers"][0]["layer_id"] = changed["layers"][1]["layer_id"]
            with self.assertRaises(AuthorityContractError):
                validate_authority_contract_document(changed)
        with self.subTest(assertion_id="T411-H-AUTH-006"):
            changed = deepcopy(contract)
            changed["layers"][2]["rank"] = 3
            with self.assertRaisesRegex(AuthorityContractError, "contiguous"):
                validate_authority_contract_document(changed)
        with self.subTest(assertion_id="T411-H-AUTH-007"):
            changed = deepcopy(contract)
            changed["layers"][0]["layer_id"] = "safety"
            with self.assertRaisesRegex(AuthorityContractError, "Safety"):
                validate_authority_contract_document(changed)
        with self.subTest(assertion_id="T411-H-AUTH-008"):
            with self.assertRaisesRegex(AuthorityContractError, "Unknown"):
                authority_rank("unknown")
        with self.subTest(assertion_id="T411-H-AUTH-009"):
            self.assertEqual(authority_contract_fingerprint(), authority_contract_fingerprint())
        with self.subTest(assertion_id="T411-H-AUTH-010"):
            verify_legacy_layer_compatibility(load_json(SCHEMAS / "common.schema.json"))
            verify_legacy_layer_compatibility(load_json(SCHEMAS / "common.v2.1.schema.json"))
        with self.subTest(assertion_id="T411-H-AUTH-011"):
            feature_schema = load_artifact_schema("feature-activation-manifest", version="2.1")
            self.assertNotIn("layers", feature_schema["properties"]["features"]["properties"])
        with self.subTest(assertion_id="T411-H-AUTH-012"):
            self.assertTrue(all(item["authority_contract_version"] in {"none", "1.0"} for item in load_artifact_contract_matrix()["routes"]))

    def test_v0411_h_canonical_001_024(self) -> None:
        documents = schema_documents()
        inventory = fingerprint_field_inventory(documents)
        with self.subTest(assertion_id="T411-H-CAN-001"):
            audit_schema_composition(documents)
        with self.subTest(assertion_id="T411-H-CAN-002"):
            guards = sorted(record["field_path"] for record in inventory if "all_of_guard" in record["schema_features"])
            self.assertEqual(EXPECTATIONS["approved_all_of_guard_fields"], guards)
        with self.subTest(assertion_id="T411-H-CAN-003"):
            changed = deepcopy(documents)
            qa_id = "https://schemas.format-monograph.local/v2.1/qa-approval-artifact.schema.json"
            changed[qa_id]["$defs"]["approver"]["allOf"][0]["then"]["required"].append("actor_id")
            with self.assertRaisesRegex(CanonicalizationError, "unaudited allOf"):
                audit_schema_composition(changed)
        with self.subTest(assertion_id="T411-H-CAN-004"):
            changed = deepcopy(documents)
            root = "https://schemas.format-monograph.local/v2/capability-snapshot.schema.json"
            changed[root]["allOf"] = [{"type": "object"}]
            with self.assertRaisesRegex(CanonicalizationError, "unaudited allOf"):
                audit_schema_composition(changed)
        for assertion_id, location in (
            ("T411-H-CAN-005", "root"),
            ("T411-H-CAN-006", "defs"),
            ("T411-H-CAN-007", "property"),
        ):
            with self.subTest(assertion_id=assertion_id):
                changed = deepcopy(documents)
                root = "https://schemas.format-monograph.local/v2/capability-snapshot.schema.json"
                if location == "root":
                    changed[root]["anyOf"] = [{"type": "object"}]
                elif location == "defs":
                    changed[root].setdefault("$defs", {})["bad"] = {"anyOf": [{"type": "string"}]}
                else:
                    changed[root]["properties"]["bad"] = {"anyOf": [{"type": "string"}]}
                with self.assertRaisesRegex(CanonicalizationError, "does not support schema composition anyOf"):
                    audit_schema_composition(changed)
        local_a = "https://schemas.format-monograph.local/test/a.schema.json"
        local_b = "https://schemas.format-monograph.local/test/b.schema.json"
        local_c = "https://schemas.format-monograph.local/test/c.schema.json"
        ref_documents = {
            local_a: {"$id": local_a, "$defs": {"value": {"$ref": local_b + "#/$defs/value"}}},
            local_b: {"$id": local_b, "$defs": {"value": {"$ref": local_c + "#/$defs/value", "x-semantic-fingerprint": "exclude"}}},
            local_c: {"$id": local_c, "$defs": {"value": {"type": "string"}}},
        }
        with self.subTest(assertion_id="T411-H-CAN-008"):
            self.assertTrue(effective_fingerprint_exclusion(ref_documents[local_a]["$defs"]["value"], local_a, ref_documents))
        with self.subTest(assertion_id="T411-H-CAN-009"):
            ref_documents[local_c]["$defs"]["value"]["x-semantic-fingerprint"] = "include"
            self.assertTrue(effective_fingerprint_exclusion(ref_documents[local_a]["$defs"]["value"], local_a, ref_documents))
        with self.subTest(assertion_id="T411-H-CAN-010"):
            with self.assertRaisesRegex(CanonicalizationError, "Remote"):
                effective_fingerprint_exclusion({"$ref": "https://example.com/schema.json"}, local_a, ref_documents)
        with self.subTest(assertion_id="T411-H-CAN-011"):
            with self.assertRaisesRegex(CanonicalizationError, "unavailable"):
                effective_fingerprint_exclusion({"$ref": local_a + "/missing"}, local_a, ref_documents)
        with self.subTest(assertion_id="T411-H-CAN-012"):
            cyclic = deepcopy(ref_documents)
            cyclic[local_c]["$defs"]["value"] = {"$ref": local_a + "#/$defs/value"}
            with self.assertRaisesRegex(CanonicalizationError, "Cyclic"):
                effective_fingerprint_exclusion(cyclic[local_a]["$defs"]["value"], local_a, cyclic)
        with self.subTest(assertion_id="T411-H-CAN-013"):
            with self.assertRaisesRegex(CanonicalizationError, "pointer"):
                effective_fingerprint_exclusion({"$ref": local_b + "#/$defs/missing"}, local_a, ref_documents)
        with self.subTest(assertion_id="T411-H-CAN-014"):
            reduced = {key: value for key, value in documents.items() if "v2.1/qa-approval" not in key}
            with self.assertRaisesRegex(CanonicalizationError, "missing allOf"):
                audit_schema_composition(reduced)
        with self.subTest(assertion_id="T411-H-CAN-015"):
            contract_id = EXPECTATIONS["non_inventory_contract_schema_ids"][0]
            changed = deepcopy(documents)
            changed[contract_id]["anyOf"] = [{"type": "object"}]
            with self.assertRaises(CanonicalizationError):
                fingerprint_field_inventory(changed)
        with self.subTest(assertion_id="T411-H-CAN-016"):
            inventory_schema_ids = {urldefrag(item["field_path"])[0] for item in inventory}
            self.assertTrue(set(EXPECTATIONS["non_inventory_contract_schema_ids"]).isdisjoint(inventory_schema_ids))
        with self.subTest(assertion_id="T411-H-CAN-017"):
            counts = EXPECTATIONS["inventory_counts"]
            self.assertEqual(counts["schema_property_nodes"], len(inventory))
            self.assertEqual(counts["semantic_projected"], sum(item["classification"] == "semantic_projected" for item in inventory))
            self.assertEqual(counts["fingerprint_excluded"], sum(item["classification"] == "fingerprint_excluded" for item in inventory))
        with self.subTest(assertion_id="T411-H-CAN-018"):
            pairs = Counter(f"{item['schema_version']}->{item['registry_contract_version']}" for item in inventory)
            self.assertEqual(EXPECTATIONS["inventory_counts"]["schema_versions"], dict(sorted(pairs.items())))
        with self.subTest(assertion_id="T411-H-CAN-019"):
            common_records = [item for item in inventory if item["field_path"].startswith("https://schemas.format-monograph.local/v2.2/common.schema.json#")]
            self.assertTrue(common_records)
            self.assertTrue(all(item["registry_contract_version"] == "2.1" for item in common_records))
        with self.subTest(assertion_id="T411-H-CAN-020"):
            policy = load_canonical_composition_policy()
            for item in policy["approved_all_of_nodes"]:
                schema = documents[item["schema_id"]]
                node = schema["$defs"]["approver"]["allOf"]
                self.assertEqual(item["node_digest"], _node_digest(node))
        with self.subTest(assertion_id="T411-H-CAN-021"):
            schema = {"oneOf": [{"type": "string"}, {"type": "integer"}]}
            resolved_string, _ = _resolve_schema(schema, local_a, {}, "text")
            resolved_integer, _ = _resolve_schema(schema, local_a, {}, 2)
            self.assertEqual("string", resolved_string["type"])
            self.assertEqual("integer", resolved_integer["type"])
            with self.assertRaises(CanonicalizationError):
                _resolve_schema(schema, local_a, {}, False)
        with self.subTest(assertion_id="T411-H-CAN-022"):
            first = feature_manifest()
            second = {key: first[key] for key in reversed(first)}
            self.assertEqual(compute_semantic_fingerprint(first), compute_semantic_fingerprint(second))
        with self.subTest(assertion_id="T411-H-CAN-023"):
            first = feature_manifest()
            first["artifact_id"] = "feature-activation-manifest:caf\u00e9"
            second = deepcopy(first)
            second["artifact_id"] = "feature-activation-manifest:cafe\u0301"
            self.assertEqual(compute_semantic_fingerprint(first), compute_semantic_fingerprint(second))
        with self.subTest(assertion_id="T411-H-CAN-024"):
            signature = inspect.signature(compute_semantic_fingerprint)
            self.assertEqual(["document"], list(signature.parameters))

    def test_v0411_h_canonical_025_030_real_qa_branches(self) -> None:
        schema = load_artifact_schema("qa-approval-artifact", version="2.1")
        validator = Draft202012Validator(
            schema, registry=artifacts.offline_schema_registry()
        )
        report = composition_report(self.feature, self.rule, status="awaiting_approval")

        with self.subTest(assertion_id="T411-H-CAN-025"):
            qa = approval(report)
            qa["previous_approval_id"] = None
            self.assertTrue(validator.is_valid(qa))
        with self.subTest(assertion_id="T411-H-CAN-026"):
            qa = approval(report)
            qa["previous_approval_id"] = "approval:previous"
            qa = stamp_test(qa)
            self.assertTrue(validator.is_valid(qa))
        with self.subTest(assertion_id="T411-H-CAN-027"):
            qa = approval(report)
            qa["previous_approval_id"] = False
            self.assertFalse(validator.is_valid(qa))
        with self.subTest(assertion_id="T411-H-CAN-028"):
            qa = approval(report)
            qa["approver"] = {"actor_id": "actor:user", "actor_role": "user"}
            qa = stamp_test(qa)
            self.assertTrue(validator.is_valid(qa))
        with self.subTest(assertion_id="T411-H-CAN-029"):
            qa = approval(report)
            qa["approver"] = {
                "actor_id": "actor:publisher",
                "actor_role": "delegated_publisher",
                "authorization_reference": {
                    "authorization_id": "authorization:h-001",
                    "granted_by_actor_id": "actor:user",
                    "authority_scope": document_scope(),
                    "issued_at": "2026-08-18T00:00:00Z",
                },
            }
            qa = stamp_test(qa)
            self.assertTrue(validator.is_valid(qa))
        with self.subTest(assertion_id="T411-H-CAN-030"):
            qa = approval(report)
            qa["approver"] = {
                "actor_id": "actor:publisher",
                "actor_role": "delegated_publisher",
            }
            self.assertFalse(validator.is_valid(qa))

    def test_v0411_h_feature_001_008(self) -> None:
        with self.subTest(assertion_id="T411-H-FEAT-001"):
            schema = load_artifact_schema("feature-activation-manifest", version="2.1")
            self.assertIn("profile_v2_composer", schema["properties"]["features"]["required"])
            self.assertFalse(feature_manifest()["features"]["profile_v2_composer"])
        with self.subTest(assertion_id="T411-H-FEAT-002"):
            self.assertFalse(profile_v2_composer_contract_enabled(None))
        with self.subTest(assertion_id="T411-H-FEAT-003"):
            invalid = feature_manifest()
            del invalid["features"]["profile_v2_composer"]
            invalid = stamp_test(invalid)
            with self.assertRaises(ArtifactContractError):
                self.validate_test(invalid)
        with self.subTest(assertion_id="T411-H-FEAT-004"):
            self.assertFalse(profile_v2_composer_contract_enabled(feature_manifest(composer=False)))
        with self.subTest(assertion_id="T411-H-FEAT-005"):
            invalid = feature_manifest(composer=True, schema_enabled=False)
            with self.assertRaisesRegex(ArtifactContractError, "requires"):
                self.validate_test(invalid)
        with self.subTest(assertion_id="T411-H-FEAT-006"):
            positive = feature_manifest(composer=True, schema_enabled=True)
            self.assertTrue(profile_v2_composer_contract_enabled(positive))
            self.assertFalse(validate_artifact(positive, features={"profile_v2_schema": True}).runtime_eligible)
        with self.subTest(assertion_id="T411-H-FEAT-007"):
            invalid = feature_manifest()
            invalid["schema_version"] = "2.9"
            self.assertFalse(profile_v2_composer_contract_enabled(invalid))
        with self.subTest(assertion_id="T411-H-FEAT-008"):
            run_source = (SCRIPTS / "run_monograph.py").read_text(encoding="utf-8")
            self.assertNotIn("profile_v2_artifacts", run_source)
            self.assertFalse(feature_manifest()["features"]["final_ready_eligible"])


if __name__ == "__main__":
    unittest.main()
