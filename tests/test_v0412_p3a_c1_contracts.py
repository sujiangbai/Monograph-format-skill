import hashlib
import inspect
import json
import sys
import unittest
from copy import deepcopy
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
SCHEMAS = ROOT / "format-monograph" / "references" / "schemas" / "v2"
sys.path.insert(0, str(SCRIPTS))

import profile_v2_artifacts as artifacts
from profile_v2_artifacts import (
    ArtifactContractError,
    ArtifactRouteError,
    load_artifact_contract_matrix,
    load_artifact_schema,
    offline_schema_registry,
    route_artifact_contract,
    schema_documents,
    validate_intent_artifact_v041,
    verify_contract_matrix_alignment,
)
from profile_v2_canonical import (
    canonical_intent_semantic_bytes_v041,
    compute_intent_semantic_fingerprint_v041,
    fingerprint_field_inventory,
    stamp_intent_semantic_fingerprint_v041,
)
from profile_v2_composer import (
    ComposerContractError,
    ComposerDisabledError,
    IntentCompositionMetrics,
    apply_intent_resolutions_v041,
    apply_resolutions,
    compose_intent_profile_v041,
    compose_profile,
)
from profile_v2_registry import (
    RegistryContractError,
    build_property_catalog_schema,
    build_typed_value_schema,
    load_registry,
    validate_registry_document,
    verify_committed_catalog,
)


RANGES = {
    "ROUTE": 36,
    "CTX": 24,
    "FEATURE": 18,
    "COVER": 24,
    "FP": 18,
    "COMPAT": 24,
    "METRICS": 12,
}
ASSERTION_IDS = tuple(
    f"T412-C1-{family}-{index:03d}"
    for family, count in RANGES.items()
    for index in range(1, count + 1)
)


def digest(value: str) -> str:
    return "sha256:" + hashlib.sha256(value.encode("utf-8")).hexdigest()


def tool() -> dict:
    return {"tool_id": "format-monograph", "version": "0.4.1"}


def scope(document_id: str = "document:book") -> dict:
    return {
        "selectors": [
            {"selector_kind": "document", "selector_ids": [document_id]}
        ],
        "exclusions": [],
        "mutually_exclusive_conditions": [],
    }


def feature(**overrides: bool) -> dict:
    features = {
        "profile_v2_schema": True,
        "profile_v2_composer": True,
        "monograph_base_v041": True,
        "final_ready_eligible": False,
    }
    features.update(overrides)
    document = {
        "artifact_kind": "feature-activation-manifest",
        "schema_version": "2.2",
        "registry_contract_version": "2.2",
        "authority_contract_version": "1.0",
        "artifact_id": "feature-activation-manifest:c01",
        "created_by_tool": tool(),
        "input_fingerprints": [
            {
                "input_id": "input:feature-source",
                "role": "source_document",
                "fingerprint": digest("feature-source"),
            }
        ],
        "semantic_fingerprint": digest("unstamped-feature"),
        "features": features,
    }
    return stamp_intent_semantic_fingerprint_v041(document)


def asset(
    label: str = "base",
    *,
    value: str = "candidate",
    confidence: str = "high",
    activation: str = "approved",
    rule_status: str = "approved",
    document_id: str = "document:book",
) -> dict:
    rule_scope = scope(document_id)
    document = {
        "artifact_kind": "layered-rule-asset",
        "schema_version": "2.2",
        "registry_contract_version": "2.2",
        "authority_contract_version": "1.0",
        "artifact_id": f"layered-rule-asset:{label}01",
        "created_by_tool": tool(),
        "input_fingerprints": [
            {
                "input_id": f"input:asset-{label}",
                "role": "rule_asset",
                "fingerprint": digest(f"asset-source-{label}"),
            }
        ],
        "semantic_fingerprint": digest(f"unstamped-asset-{label}"),
        "layer_kind": "monograph_base",
        "can_override_safety_invariants": False,
        "activation": activation,
        "asset_scope": deepcopy(rule_scope),
        "allowed_semantic_object_kinds": ["paragraph"],
        "rules": [
            {
                "rule_id": f"RULE-{label.upper()}-01",
                "semantic_object_kind": "paragraph",
                "scope": deepcopy(rule_scope),
                "confidence": confidence,
                "status": rule_status,
                "properties": [
                    {
                        "property_id": "contract.intent-probe",
                        "value": {"type": "string", "value": value},
                        "unit_id": None,
                        "mode": "automatic",
                    }
                ],
            }
        ],
    }
    return stamp_intent_semantic_fingerprint_v041(document)


def compose(assets: list[dict], manifest: dict | None = None, *, metrics: bool = False):
    return compose_intent_profile_v041(
        assets,
        manifest or feature(),
        input_fingerprint=digest("source-document"),
        structure_fingerprint=digest("structure-map"),
        artifact_id="conflict-report:c01",
        created_by_tool=tool(),
        generated_at="2026-08-18T00:00:00Z",
        include_metrics=metrics,
    )


def approval(report: dict, *, actor_role: str = "user") -> dict:
    conflict = report["approval_required_conflicts"][0]
    approver = {"actor_id": "actor:user-01", "actor_role": actor_role}
    if actor_role == "delegated_publisher":
        approver["authorization_reference"] = {
            "authorization_id": "authorization:user-01",
            "granted_by_actor_id": "actor:user-01",
            "authority_scope": deepcopy(conflict["key"]["normalized_scope"]),
            "issued_at": "2026-08-18T00:00:00Z",
        }
    document = {
        "artifact_kind": "qa-approval-artifact",
        "schema_version": "2.2",
        "registry_contract_version": "2.2",
        "authority_contract_version": "1.0",
        "artifact_id": "qa-approval-artifact:c01",
        "created_by_tool": tool(),
        "input_fingerprints": [
            {"input_id": "input:approval-source", "role": "source_document", "fingerprint": report["bindings"]["input_fingerprint"]},
            {"input_id": "input:approval-structure", "role": "structure", "fingerprint": report["bindings"]["structure_fingerprint"]},
            {"input_id": "input:approval-report", "role": "conflict_report", "fingerprint": report["semantic_fingerprint"]},
        ],
        "semantic_fingerprint": digest("unstamped-approval"),
        "approval_id": "approval:c01",
        "approver": approver,
        "decision_type": "conflict_resolution",
        "decision": "adopt_proposed",
        "reason": "Synthetic contract approval.",
        "created_at": "2026-08-18T00:00:00Z",
        "bindings": {
            "input_fingerprint": report["bindings"]["input_fingerprint"],
            "structure_fingerprint": report["bindings"]["structure_fingerprint"],
            "composition_report_fingerprint": report["semantic_fingerprint"],
        },
        "target": {
            "conflict_id": conflict["conflict_id"],
            "proposed_resolution_id": conflict["proposed_resolution_id"],
            "normalized_scope": deepcopy(conflict["key"]["normalized_scope"]),
        },
        "previous_approval_id": None,
    }
    return stamp_intent_semantic_fingerprint_v041(document)


def legacy_manifest() -> dict:
    document = {
        "artifact_kind": "legacy-migration-manifest",
        "schema_version": "2.1",
        "registry_contract_version": "2.2",
        "authority_contract_version": "1.0",
        "artifact_id": "legacy-migration-manifest:c01",
        "created_by_tool": tool(),
        "input_fingerprints": [
            {
                "input_id": "input:legacy-profile",
                "role": "source_document",
                "fingerprint": digest("legacy-profile"),
            }
        ],
        "semantic_fingerprint": digest("unstamped-migration"),
        "legacy_input": True,
        "activation": "disabled",
        "source_schema_versions": ["1.0", "1.1"],
        "mappings": [],
    }
    return stamp_intent_semantic_fingerprint_v041(document)


class P3aC1ContractTests(unittest.TestCase):
    maxDiff = None

    def _record(self, assertion_id: str, condition: bool, message: str = "") -> None:
        self.assertIn(assertion_id, ASSERTION_IDS)
        self.assertTrue(condition, message or assertion_id)

    def test_assertion_manifest_is_exact(self) -> None:
        self.assertEqual(156, len(ASSERTION_IDS))
        self.assertEqual(len(ASSERTION_IDS), len(set(ASSERTION_IDS)))
        for family, count in RANGES.items():
            expected = [f"T412-C1-{family}-{index:03d}" for index in range(1, count + 1)]
            self.assertEqual(expected, [item for item in ASSERTION_IDS if f"-{family}-" in item])

    def test_route_contract_assertions(self) -> None:
        old = load_artifact_contract_matrix("1.0")
        new = load_artifact_contract_matrix("1.1")
        verify_contract_matrix_alignment()
        verify_contract_matrix_alignment(matrix_version="1.1")
        new_routes = [item for item in new["routes"] if item["registry_contract_version"] == "2.2"]
        report = compose([asset()]).report
        final_profile = apply_intent_resolutions_v041(
            report,
            [],
            feature(),
            task_id="task:route-cross-check",
            task_fingerprint=digest("route-cross-check"),
            artifact_id="final-execution-profile:route-cross-check",
            created_by_tool=tool(),
        ).final_profile
        approval_report = compose([asset(confidence="medium")]).report
        samples = [
            asset(),
            feature(),
            approval(approval_report),
            report,
            final_profile,
            legacy_manifest(),
        ]
        schemas_by_kind = {
            item["artifact_kind"]: load_artifact_schema(
                item["artifact_kind"],
                version=item["schema_version"],
                registry_contract_version=item["registry_contract_version"],
                authority_contract_version=item["authority_contract_version"],
                matrix_version="1.1",
            )
            for item in new_routes
        }
        schema_registry = offline_schema_registry(matrix_version="1.1")
        cross_artifact_contract_is_closed = all(
            not list(
                Draft202012Validator(
                    schemas_by_kind[sample["artifact_kind"]],
                    registry=schema_registry,
                    format_checker=FormatChecker(),
                ).iter_errors(sample)
            )
            and all(
                list(
                    Draft202012Validator(
                        schema,
                        registry=schema_registry,
                        format_checker=FormatChecker(),
                    ).iter_errors(sample)
                )
                for artifact_kind, schema in schemas_by_kind.items()
                if artifact_kind != sample["artifact_kind"]
            )
            for sample in samples
        )
        checks = [
            new["contract_version"] == "1.1",
            old["contract_version"] == "1.0",
            len(new["routes"]) == len(old["routes"]) + 6,
            all(
                item
                == {
                    key: value
                    for key, value in next(
                        route
                        for route in new["routes"]
                        if route["route_id"] == item["route_id"]
                    ).items()
                    if key != "registry_validation_context"
                }
                for item in old["routes"]
            ),
            len(new_routes) == 6,
            all(item["registry_validation_context"] == "declaration_intent" for item in new_routes),
            all(item["authority_contract_version"] == "1.0" for item in new_routes),
            len(schema_documents()) == 28,
            len(schema_documents(matrix_version="1.1")) == 39,
            route_artifact_contract(feature(), matrix_version="1.1").schema_version == "2.2",
            route_artifact_contract(asset(), matrix_version="1.1").registry_contract_version == "2.2",
            load_artifact_schema("conflict-report", version="2.3", registry_contract_version="2.2", authority_contract_version="1.0", matrix_version="1.1")["properties"]["schema_version"]["const"] == "2.3",
            load_artifact_schema("final-execution-profile", version="2.3", registry_contract_version="2.2", authority_contract_version="1.0", matrix_version="1.1")["properties"]["runtime_eligible"]["const"] is False,
            all(route["registry_contract_version"] != "2.2" for route in old["routes"]),
            all(route["schema_version"] != "2.3" for route in old["routes"]),
            all(route["version_source"] == "artifact_fields" for route in new_routes),
            any(route["artifact_kind"] == "legacy-migration-manifest" and route["schema_version"] == "2.1" for route in new_routes),
            any(item["schema_version"] == "2.3" and item["registry_contract_version"] == "2.2" for item in new["schema_resources"]),
        ]
        invalids = []
        for field, value in (
            ("schema_version", "2.9"),
            ("registry_contract_version", "2.1"),
            ("authority_contract_version", "none"),
        ):
            bad = feature()
            bad[field] = value
            try:
                route_artifact_contract(bad, matrix_version="1.1")
                invalids.append(False)
            except ArtifactRouteError:
                invalids.append(True)
        checks.extend(invalids)
        checks.extend(
            [
                cross_artifact_contract_is_closed,
                len({item["route_id"] for item in new["routes"]}) == len(new["routes"]),
                len({(item["artifact_kind"], item["schema_version"], item["registry_contract_version"], item["authority_contract_version"]) for item in new["routes"]}) == len(new["routes"]),
                set(artifacts.ARTIFACT_KINDS) == {item["artifact_kind"] for item in new["routes"]},
                all(Path(SCHEMAS / item["schema_file"]).is_file() for item in new["routes"]),
                all(Path(SCHEMAS / item["schema_file"]).is_file() for item in new["schema_resources"]),
                all(item["registry_validation_context"] == "strict_execution" for item in new["routes"] if item["registry_contract_version"] != "2.2"),
                load_registry(version="2.2", validation_context="declaration_intent")["schema_version"] == "2.2",
                load_registry(version="2.1")["schema_version"] == "2.1",
                json.loads((SCHEMAS / "artifact-contract-matrix.v1.0.json").read_text(encoding="utf-8")) == old,
                all(route["artifact_kind"] in artifacts.ARTIFACT_KINDS for route in new["routes"]),
                all(route["registry_contract_version"] in {"2.0", "2.1", "2.2"} for route in new["routes"]),
                all(route["authority_contract_version"] in {"none", "1.0"} for route in new["routes"]),
                all(route["schema_id"].startswith("https://schemas.format-monograph.local/") for route in new["routes"]),
                all(item["fingerprint_inventory"] is (item["resource_kind"] == "fingerprint_shared") for item in new["schema_resources"]),
            ]
        )
        self.assertEqual(36, len(checks))
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-ROUTE-{index:03d}"):
                self._record(f"T412-C1-ROUTE-{index:03d}", condition)

    def test_registry_context_assertions(self) -> None:
        registry = load_registry(version="2.2", validation_context="declaration_intent")
        checks = []
        checks.append(registry["schema_version"] == "2.2")
        checks.append(any(item["availability"] == "reserved" for item in registry["executor_capabilities"]))
        checks.append(any(item["availability"] == "reserved" for item in registry["auditor_capabilities"]))
        checks.append(any("automatic" in item["modes"] for item in registry["properties"]))
        checks.append(build_property_catalog_schema(registry, validation_context="declaration_intent") == json.loads((SCHEMAS / "property-catalog.v2.2.generated.schema.json").read_text(encoding="utf-8")))
        checks.append(build_typed_value_schema(registry, validation_context="declaration_intent") == json.loads((SCHEMAS / "typed-value.v2.2.generated.schema.json").read_text(encoding="utf-8")))
        checks.append(not _raises(lambda: verify_committed_catalog(registry, version="2.2", validation_context="declaration_intent")))
        checks.append(_raises(lambda: validate_registry_document(registry), RegistryContractError))
        checks.append(_raises(lambda: validate_registry_document(registry, validation_context="unknown"), RegistryContractError))
        checks.append(_raises(lambda: validate_registry_document(load_registry(version="2.1"), validation_context="declaration_intent"), RegistryContractError))
        planned = deepcopy(registry); planned["executor_capabilities"][1]["availability"] = "planned"
        checks.append(_raises(lambda: validate_registry_document(planned, validation_context="declaration_intent"), RegistryContractError))
        unavailable = deepcopy(registry); unavailable["executor_capabilities"][1]["availability"] = "unavailable"
        checks.append(_raises(lambda: validate_registry_document(unavailable, validation_context="declaration_intent"), RegistryContractError))
        no_audit = deepcopy(registry); no_audit["auditor_capabilities"][1]["availability"] = "unavailable"
        checks.append(_raises(lambda: validate_registry_document(no_audit, validation_context="declaration_intent"), RegistryContractError))
        implemented = deepcopy(registry); implemented["executor_capabilities"][1]["availability"] = "implemented"; implemented["auditor_capabilities"][1]["availability"] = "implemented"
        checks.append(not _raises(lambda: validate_registry_document(implemented)))
        checks.extend([
            set(item["availability"] for item in registry["executor_capabilities"]) <= {"unavailable", "reserved", "implemented"},
            set(item["availability"] for item in registry["auditor_capabilities"]) <= {"unavailable", "reserved", "implemented"},
            registry["registry_id"] == "registry:profile-v2-intent-contract-core",
            len(registry["properties"]) == 2,
            any(item["safety_invariant"] for item in registry["properties"]),
            any(item["property_id"] == "contract.intent-probe" for item in registry["properties"]),
            all(not item.get("test_only") for item in registry["properties"]),
            all(not item["capability_id"].endswith("stub") for item in registry["executor_capabilities"]),
            all(not item["capability_id"].endswith("stub") for item in registry["auditor_capabilities"]),
            canonical_intent_semantic_bytes_v041(feature()) == canonical_intent_semantic_bytes_v041(deepcopy(feature())),
        ])
        self.assertEqual(24, len(checks))
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-CTX-{index:03d}"):
                self._record(f"T412-C1-CTX-{index:03d}", condition)

    def test_feature_gate_and_qa_branch_assertions(self) -> None:
        checks = [not _raises(lambda: compose([asset()], feature()))]
        for flag in ("profile_v2_schema", "profile_v2_composer", "monograph_base_v041"):
            checks.append(_raises(lambda flag=flag: compose([asset()], feature(**{flag: False})), ComposerDisabledError))
        for flag in ("profile_v2_schema", "profile_v2_composer", "monograph_base_v041"):
            missing = feature(); del missing["features"][flag]; missing = stamp_intent_semantic_fingerprint_v041(missing)
            checks.append(_raises(lambda missing=missing: compose([asset()], missing), ComposerDisabledError))
        stale = feature(); stale["semantic_fingerprint"] = digest("stale")
        checks.append(_raises(lambda: compose([asset()], stale), ComposerDisabledError))
        checks.append(
            _raises(
                lambda: compose_intent_profile_v041(
                    [asset()],
                    None,
                    input_fingerprint=digest("source-document"),
                    structure_fingerprint=digest("structure-map"),
                    artifact_id="conflict-report:c01",
                    created_by_tool=tool(),
                    generated_at="2026-08-18T00:00:00Z",
                ),
                ComposerDisabledError,
            )
        )
        disabled_report = compose([asset()]).report
        checks.append(
            disabled_report["activation"] == "disabled"
            and disabled_report["runtime_eligible"] is False
            and disabled_report["final_ready_eligible"] is False
            and disabled_report["delivery_allowed"] is False
        )
        medium = compose([asset(confidence="medium")]).report
        user = approval(medium, actor_role="user")
        delegated = approval(medium, actor_role="delegated_publisher")
        checks.append(not _raises(lambda: validate_intent_artifact_v041(user)))
        checks.append(not _raises(lambda: validate_intent_artifact_v041(delegated)))
        user_with_reference = deepcopy(user)
        user_with_reference["approver"]["authorization_reference"] = deepcopy(
            delegated["approver"]["authorization_reference"]
        )
        user_with_reference = stamp_intent_semantic_fingerprint_v041(user_with_reference)
        checks.append(not _raises(lambda: validate_intent_artifact_v041(user_with_reference)))
        old_qa = load_artifact_schema("qa-approval-artifact", version="2.0")
        new_qa = load_artifact_schema(
            "qa-approval-artifact",
            version="2.2",
            registry_contract_version="2.2",
            authority_contract_version="1.0",
            matrix_version="1.1",
        )
        checks.append(
            set(old_qa["$defs"]["approver"]["properties"]["actor_role"]["enum"])
            == set(new_qa["$defs"]["approver"]["properties"]["actor_role"]["enum"])
            == {"user", "delegated_publisher"}
        )
        checks.append(
            {
                branch["properties"]["actor_role"]["const"]
                for branch in new_qa["$defs"]["approver"]["oneOf"]
            }
            == {"user", "delegated_publisher"}
        )
        missing_auth = deepcopy(delegated); del missing_auth["approver"]["authorization_reference"]
        checks.append(_raises(lambda: validate_intent_artifact_v041(missing_auth), ArtifactContractError))
        double_branch = deepcopy(user); double_branch["approver"]["actor_role"] = ["user", "delegated_publisher"]
        checks.append(_raises(lambda: validate_intent_artifact_v041(double_branch), ArtifactContractError))
        unknown_actor = deepcopy(user); unknown_actor["approver"]["actor_role"] = "developer"
        checks.append(_raises(lambda: validate_intent_artifact_v041(unknown_actor), ArtifactContractError))
        self.assertEqual(18, len(checks))
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-FEATURE-{index:03d}"):
                self._record(f"T412-C1-FEATURE-{index:03d}", condition)

    def test_coverage_assertions(self) -> None:
        result = compose([asset()], metrics=True)
        checks = [
            result.status == "resolvable",
            len(result.report["candidate_groups"]) == 1,
            len(result.report["proposed_resolutions"]) == 1,
            result.metrics.input_asset_count == 1,
            result.metrics.input_rule_count == 1,
            result.metrics.input_binding_count == 1,
            _raises(lambda: compose([], feature()), ComposerContractError),
            _raises(lambda: compose([asset(activation="draft")]), ComposerContractError),
            _raises(lambda: compose([asset(activation="disabled")]), ComposerContractError),
            _raises(lambda: compose([asset(activation="rejected")]), (ComposerContractError, ArtifactContractError)),
            _raises(lambda: compose([asset(rule_status="draft")]), ComposerContractError),
            _raises(lambda: compose([asset(rule_status="disabled")]), (ComposerContractError, ArtifactContractError)),
        ]
        duplicate = asset(); checks.append(_raises(lambda: compose([duplicate, deepcopy(duplicate)]), ComposerContractError))
        no_rules = asset(); no_rules["rules"] = []; no_rules = stamp_intent_semantic_fingerprint_v041(no_rules); checks.append(_raises(lambda: compose([no_rules]), (ComposerContractError, ArtifactContractError)))
        no_bindings = asset(); no_bindings["rules"][0]["properties"] = []; no_bindings = stamp_intent_semantic_fingerprint_v041(no_bindings); checks.append(_raises(lambda: compose([no_bindings]), (ComposerContractError, ArtifactContractError)))
        bad_source = asset(); bad_source["semantic_fingerprint"] = digest("bad"); checks.append(_raises(lambda: compose([bad_source]), ArtifactContractError))
        two = compose([asset("one", document_id="document:one"), asset("two", document_id="document:two")], metrics=True)
        checks.extend([
            two.metrics.input_asset_count == 2,
            two.metrics.input_binding_count == 2,
            two.metrics.candidate_count == 2,
            two.metrics.expected_key_count == 2,
            len(two.report["scope_partitions"]) == 2,
            all(group["candidates"] for group in two.report["candidate_groups"]),
            all(proposal["candidate_chain"] for proposal in two.report["proposed_resolutions"]),
            result.report["fatal_diagnostics"] == [],
        ])
        self.assertEqual(24, len(checks))
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-COVER-{index:03d}"):
                self._record(f"T412-C1-COVER-{index:03d}", condition)

    def test_fingerprint_assertions(self) -> None:
        first = compose([asset()], metrics=False)
        second = compose([deepcopy(asset())], metrics=True)
        old_docs = schema_documents()
        new_docs = schema_documents(matrix_version="1.1")
        inventory = fingerprint_field_inventory(new_docs, matrix_version="1.1")
        source_assets = [asset()]
        source_manifest = feature()
        source_assets_before = deepcopy(source_assets)
        source_manifest_before = deepcopy(source_manifest)
        immutable_result = compose_intent_profile_v041(
            source_assets,
            source_manifest,
            input_fingerprint=digest("source-document"),
            structure_fingerprint=digest("structure-map"),
            artifact_id="conflict-report:immutable-inputs",
            created_by_tool=tool(),
            generated_at="2026-08-18T00:00:00Z",
        )
        report_before_apply = deepcopy(immutable_result.report)
        apply_intent_resolutions_v041(
            immutable_result.report,
            [],
            source_manifest,
            task_id="task:immutable-inputs",
            task_fingerprint=digest("immutable-inputs"),
            artifact_id="final-execution-profile:immutable-inputs",
            created_by_tool=tool(),
        )
        checks = [
            first.report == second.report,
            first.report["semantic_fingerprint"] == second.report["semantic_fingerprint"],
            compute_intent_semantic_fingerprint_v041(first.report) == first.report["semantic_fingerprint"],
            canonical_intent_semantic_bytes_v041(first.report) == canonical_intent_semantic_bytes_v041(second.report),
            len(old_docs) == 28,
            len(new_docs) == 39,
            len(inventory) > 625,
            all(record["schema_version"] for record in inventory),
            all(record["registry_contract_version"] for record in inventory),
            any(record["schema_version"] == "2.3" and record["registry_contract_version"] == "2.2" for record in inventory),
            json.loads((SCHEMAS / "property-catalog.v2.2.generated.schema.json").read_text(encoding="utf-8")) == build_property_catalog_schema(load_registry(version="2.2", validation_context="declaration_intent"), validation_context="declaration_intent"),
            json.loads((SCHEMAS / "typed-value.v2.2.generated.schema.json").read_text(encoding="utf-8")) == build_typed_value_schema(load_registry(version="2.2", validation_context="declaration_intent"), validation_context="declaration_intent"),
            source_assets == source_assets_before
            and source_manifest == source_manifest_before,
            immutable_result.report == report_before_apply,
            compose([asset("a"), asset("b")]).report == compose([asset("b"), asset("a")]).report,
            first.metrics is None,
            second.metrics is not None,
            b"metrics" not in canonical_intent_semantic_bytes_v041(second.report),
        ]
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-FP-{index:03d}"):
                self._record(f"T412-C1-FP-{index:03d}", condition)

    def test_compatibility_assertions(self) -> None:
        old = load_artifact_contract_matrix("1.0")
        new = load_artifact_contract_matrix("1.1")
        run_text = (SCRIPTS / "run_monograph.py").read_text(encoding="utf-8")
        old_routes = {(item["artifact_kind"], item["schema_version"], item["registry_contract_version"], item["authority_contract_version"], item["schema_file"]) for item in old["routes"]}
        new_routes = {(item["artifact_kind"], item["schema_version"], item["registry_contract_version"], item["authority_contract_version"], item["schema_file"]) for item in new["routes"]}
        checks = [
            old_routes <= new_routes,
            "intent" not in inspect.signature(compose_profile).parameters,
            "intent" not in inspect.signature(apply_resolutions).parameters,
            "compose_intent_profile_v041" not in run_text,
            "profile_v2_composer" not in run_text,
            artifacts.CONTRACT_MATRIX_PATH.name == "artifact-contract-matrix.v1.0.json",
            artifacts.ARTIFACT_SCHEMA_FILES == {(item["artifact_kind"], item["schema_version"], item["registry_contract_version"], item["authority_contract_version"]): item["schema_file"] for item in old["routes"]},
            load_registry(version="2.1")["registry_id"] == "registry:profile-v2-core",
            load_registry(version="2.0")["schema_version"] == "2.0",
            load_artifact_schema("conflict-report", version="2.2")["properties"]["schema_version"]["const"] == "2.2",
            load_artifact_schema("final-execution-profile", version="2.2")["properties"]["schema_version"]["const"] == "2.2",
            load_artifact_schema("qa-approval-artifact", version="2.1")["properties"]["schema_version"]["const"] == "2.1",
            load_artifact_contract_matrix()["contract_version"] == "1.0",
            len(schema_documents()) == 28,
            not _raises(lambda: verify_contract_matrix_alignment()),
            not _raises(lambda: verify_contract_matrix_alignment(matrix_version="1.1")),
            all(item["schema_version"] != "2.3" for item in old["routes"]),
            all(item["registry_contract_version"] != "2.2" for item in old["routes"]),
            all("monograph_base_v041" not in json.dumps(load_artifact_schema("feature-activation-manifest", version="2.1")) for _ in [0]),
            load_registry(version="2.2", validation_context="declaration_intent")["registry_id"] != load_registry(version="2.1")["registry_id"],
            not any(path.name.startswith("monograph-base") for path in SCHEMAS.iterdir()),
            "argparse" not in (SCRIPTS / "profile_v2_composer.py").read_text(encoding="utf-8"),
            "if __name__" not in (SCRIPTS / "profile_v2_composer.py").read_text(encoding="utf-8"),
            all(item["version_source"] == "artifact_fields" for item in new["routes"] if item["registry_contract_version"] == "2.2"),
        ]
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-COMPAT-{index:03d}"):
                self._record(f"T412-C1-COMPAT-{index:03d}", condition)

    def test_metrics_assertions(self) -> None:
        without = compose([asset()], metrics=False)
        with_metrics = compose([asset()], metrics=True)
        metrics = with_metrics.metrics
        applied = apply_intent_resolutions_v041(
            with_metrics.report,
            [],
            feature(),
            task_id="task:c01",
            task_fingerprint=digest("task"),
            artifact_id="final-execution-profile:c01",
            created_by_tool=tool(),
            metrics=metrics,
        )
        checks = [
            without.metrics is None,
            isinstance(metrics, IntentCompositionMetrics),
            json.loads(json.dumps(metrics.as_dict())) == metrics.as_dict(),
            metrics.input_asset_count == 1,
            metrics.input_rule_count == 1,
            metrics.input_binding_count == 1,
            metrics.candidate_count == 1,
            metrics.expected_key_count == 1,
            metrics.max_candidates_per_key == 1,
            without.report == with_metrics.report,
            without.report["semantic_fingerprint"] == with_metrics.report["semantic_fingerprint"],
            applied.metrics == metrics and applied.final_profile["runtime_eligible"] is False,
        ]
        for index, condition in enumerate(checks, 1):
            with self.subTest(assertion_id=f"T412-C1-METRICS-{index:03d}"):
                self._record(f"T412-C1-METRICS-{index:03d}", condition)


def _raises(callable_value, error_type=Exception) -> bool:
    try:
        callable_value()
    except error_type:
        return True
    except Exception:
        return False
    return False


if __name__ == "__main__":
    unittest.main()
