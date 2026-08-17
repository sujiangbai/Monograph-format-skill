from __future__ import annotations

import json
import re
import sys
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path
from unittest import mock

from jsonschema import Draft202012Validator


REPO = Path(__file__).resolve().parents[1]
SKILL = REPO / "format-monograph"
SCRIPTS = SKILL / "scripts"
SCHEMA_DIR = SKILL / "references" / "schemas" / "v2"
FIXTURES = REPO / "tests" / "fixtures" / "v041"
sys.path.insert(0, str(SCRIPTS))

from profile_v2_artifacts import (  # noqa: E402
    ARTIFACT_KINDS,
    ArtifactContractError,
    ProfileV2DisabledError,
    load_artifact_schema,
    offline_schema_registry,
    read_profile_document,
    require_v2_nonlegacy_contract,
    schema_documents,
    schema_errors,
    schema_for_requested_minor,
    validate_artifact,
)
from profile_v2_registry import (  # noqa: E402
    GENERATED_CATALOG_PATH,
    GENERATED_TYPED_VALUE_PATH,
    RegistryContractError,
    build_property_catalog_schema,
    build_typed_value_schema,
    catalog_differences,
    load_registry,
    typed_value_differences,
    validate_registry_document,
    verify_committed_catalog,
)
from validate_profile import validate as validate_legacy_path  # noqa: E402


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def minimal_artifacts() -> dict[str, dict]:
    return load_json(FIXTURES / "minimal-artifacts.json")


def resolved_final_profile() -> dict:
    return load_json(FIXTURES / "resolved-final-profile.json")


def legacy_profile(version: str) -> dict:
    profile = {
        "schema_version": version,
        "profile_id": f"synthetic-legacy-{version.replace('.', '-')}",
        "name": "Synthetic legacy profile",
        "locale": "zh-CN",
        "scope": {"document_type": "monograph", "input_format": "DOCX"},
        "target_applications": ["Microsoft 365"],
        "source_precedence": ["user_requirement", "written_requirement"],
        "sources": [
            {
                "id": "SRC-001",
                "type": "user_requirement",
                "label": "Synthetic requirement",
                "summary": "Synthetic test-only source.",
                "public": True,
            }
        ],
        "rules": [
            {
                "id": "FMT-BODY-001",
                "category": "body",
                "selector": {"kind": "paragraph_role", "value": "body"},
                "properties": {"font_size_pt": 10.5},
                "source_ids": ["SRC-001"],
                "evidence_summary": "Synthetic test-only rule.",
                "confidence": "high",
                "status": "approved",
                "application": "automatic",
            }
        ],
        "conflicts": [],
        "open_questions": [],
        "approval": {"status": "draft"},
    }
    if version == "1.1":
        profile["runtime_policy"] = {
            "caller_requirements_highest": True,
            "editable_equations_required": True,
            "formula_image_policy": "block",
        }
    return profile


class ArtifactSchemaTests(unittest.TestCase):
    """T41-SCH-001..024 and V041-ART-001."""

    def setUp(self) -> None:
        self.artifacts = minimal_artifacts()

    def test_t41_sch_001_all_eight_artifact_schemas_are_valid(self) -> None:
        self.assertEqual(set(ARTIFACT_KINDS), set(self.artifacts))
        for kind in ARTIFACT_KINDS:
            with self.subTest(kind=kind):
                schema = load_artifact_schema(kind)
                Draft202012Validator.check_schema(schema)
                self.assertFalse(schema_errors(kind, self.artifacts[kind]))

    def test_t41_sch_002_all_schema_documents_pass_metaschema(self) -> None:
        documents = schema_documents()
        self.assertEqual(12, len(documents))
        for schema_id, schema in documents.items():
            with self.subTest(schema_id=schema_id):
                Draft202012Validator.check_schema(schema)

    def test_t41_sch_003_cross_artifact_rejection_is_complete_8_by_7(self) -> None:
        rejected = 0
        for source_kind, artifact in self.artifacts.items():
            for target_kind in ARTIFACT_KINDS:
                if source_kind == target_kind:
                    continue
                with self.subTest(source=source_kind, target=target_kind):
                    self.assertTrue(schema_errors(target_kind, artifact))
                    rejected += 1
        self.assertEqual(56, rejected)

    def test_t41_sch_004_top_level_and_nested_unknown_fields_are_rejected(self) -> None:
        for kind, artifact in self.artifacts.items():
            top_level = deepcopy(artifact)
            top_level["unexpected"] = True
            nested = deepcopy(artifact)
            nested["created_by_tool"]["unexpected"] = True
            with self.subTest(kind=kind, level="top"):
                self.assertTrue(schema_errors(kind, top_level))
            with self.subTest(kind=kind, level="nested"):
                self.assertTrue(schema_errors(kind, nested))

        def object_contracts(value: object) -> list[dict]:
            contracts: list[dict] = []
            if isinstance(value, dict):
                if value.get("type") == "object":
                    contracts.append(value)
                for child in value.values():
                    contracts.extend(object_contracts(child))
            elif isinstance(value, list):
                for child in value:
                    contracts.extend(object_contracts(child))
            return contracts

        for schema_id, schema in schema_documents().items():
            for contract in object_contracts(schema):
                with self.subTest(schema_id=schema_id, contract=contract.get("title")):
                    self.assertIs(False, contract.get("additionalProperties"))

    def test_t41_sch_005_common_identifiers_and_fingerprints_are_strict(self) -> None:
        for field, value in (
            ("artifact_kind", "wrong-kind"),
            ("artifact_id", "not valid"),
            ("schema_version", "v2"),
            ("semantic_fingerprint", "sha256:not-a-digest"),
        ):
            artifact = deepcopy(self.artifacts["capability-snapshot"])
            artifact[field] = value
            with self.subTest(field=field):
                self.assertTrue(schema_errors("capability-snapshot", artifact))

        duplicate = deepcopy(self.artifacts["capability-snapshot"])
        duplicate["input_fingerprints"].append(deepcopy(duplicate["input_fingerprints"][0]))
        self.assertTrue(schema_errors("capability-snapshot", duplicate))

    def test_t41_sch_006_semantic_fingerprint_excludes_nonsemantic_metadata(self) -> None:
        for kind in ARTIFACT_KINDS:
            schema = load_artifact_schema(kind)
            excluded = set(schema["x-semantic-fingerprint-excludes"])
            self.assertTrue({"created_at", "display_name", "absolute_path"} <= excluded)

    def test_t41_sch_007_local_ref_closure_is_offline(self) -> None:
        with mock.patch("urllib.request.urlopen", side_effect=AssertionError("network access")):
            registry = offline_schema_registry()
            self.assertIsNotNone(registry)
            for kind, artifact in self.artifacts.items():
                with self.subTest(kind=kind):
                    self.assertFalse(schema_errors(kind, artifact))

    def test_t41_sch_008_feature_gate_defaults_to_closed(self) -> None:
        artifact = self.artifacts["capability-snapshot"]
        with self.assertRaises(ProfileV2DisabledError):
            validate_artifact(artifact)
        with self.assertRaises(ProfileV2DisabledError):
            validate_artifact(artifact, features={"profile_v2_schema": False})
        result = validate_artifact(artifact, features={"profile_v2_schema": True})
        self.assertFalse(result.runtime_eligible)
        self.assertEqual("disabled", result.activation)

    def test_t41_sch_009_artifact_categories_cannot_be_mixed(self) -> None:
        foreign_fields = {
            "qa-approval-artifact": ("producer", {"actor_id": "actor:test", "actor_role": "developer"}),
            "execution-evidence-artifact": ("approval_id", "approval:test"),
            "capability-snapshot": ("decision", "approve"),
            "conflict-report": ("merge_authorization", True),
        }
        for kind, (field, value) in foreign_fields.items():
            artifact = deepcopy(self.artifacts[kind])
            artifact[field] = value
            with self.subTest(kind=kind):
                self.assertTrue(schema_errors(kind, artifact))


class RuntimeAuthorityAndEvidenceTests(unittest.TestCase):
    """T41-SCH-010..016, V041-RUNTIME-AUTH-001, and V041-EVIDENCE-001."""

    def setUp(self) -> None:
        self.artifacts = minimal_artifacts()

    def test_t41_sch_010_qa_approver_enum_rejects_three_nonapprovers(self) -> None:
        for role in ("project_manager", "developer", "automated_qa"):
            artifact = deepcopy(self.artifacts["qa-approval-artifact"])
            artifact["approver"]["actor_role"] = role
            with self.subTest(role=role):
                self.assertTrue(schema_errors("qa-approval-artifact", artifact))

    def test_t41_sch_011_delegated_publisher_requires_structured_authorization(self) -> None:
        artifact = deepcopy(self.artifacts["qa-approval-artifact"])
        artifact["approver"]["actor_role"] = "delegated_publisher"
        self.assertTrue(schema_errors("qa-approval-artifact", artifact))
        artifact["approver"]["authorization_reference"] = {}
        self.assertTrue(schema_errors("qa-approval-artifact", artifact))
        artifact["approver"]["authorization_reference"] = {
            "authorization_id": "authorization:test",
            "granted_by_actor_id": "actor:user",
            "authority_scope": {"scope_kind": "document", "scope_ids": ["document:test"]},
            "issued_at": "2026-01-01T00:00:00Z",
        }
        self.assertFalse(schema_errors("qa-approval-artifact", artifact))

    def test_t41_sch_012_evidence_producer_rejects_approval_roles(self) -> None:
        for role in ("user", "delegated_publisher"):
            artifact = deepcopy(self.artifacts["execution-evidence-artifact"])
            artifact["producer"]["actor_role"] = role
            with self.subTest(role=role):
                self.assertTrue(schema_errors("execution-evidence-artifact", artifact))

    def test_t41_sch_013_delivery_and_final_ready_cannot_be_encoded(self) -> None:
        mutations = (
            ("evidence_class", "delivery"),
            ("non_delivery", False),
            ("final_ready_eligible", True),
            ("evidence_scope", "v0.4.3"),
        )
        for field, value in mutations:
            artifact = deepcopy(self.artifacts["execution-evidence-artifact"])
            artifact[field] = value
            with self.subTest(field=field):
                self.assertTrue(schema_errors("execution-evidence-artifact", artifact))

    def test_t41_sch_014_evidence_history_and_supersedes_are_closed(self) -> None:
        artifact = deepcopy(self.artifacts["execution-evidence-artifact"])
        artifact["supersedes"] = ["evidence:older"]
        artifact["history"].append(
            {
                "entry_id": "history:superseded",
                "event_type": "superseded",
                "created_at": "2026-01-02T00:00:00Z",
                "reason": "Synthetic superseding relationship.",
            }
        )
        self.assertFalse(schema_errors("execution-evidence-artifact", artifact))
        artifact["history"][0]["mutable_payload"] = True
        self.assertTrue(schema_errors("execution-evidence-artifact", artifact))

    def test_t41_sch_015_final_profile_keeps_p1_safety_contract_disabled(self) -> None:
        artifact = deepcopy(self.artifacts["final-execution-profile"])
        for field, value in (
            ("final_ready_eligible", True),
            ("delivery_allowed", True),
            ("activation", "enabled"),
        ):
            mutated = deepcopy(artifact)
            mutated[field] = value
            with self.subTest(field=field):
                self.assertTrue(schema_errors("final-execution-profile", mutated))
        mutated = deepcopy(artifact)
        mutated["safety_invariants"]["author_content_mutation_allowed"] = True
        self.assertTrue(schema_errors("final-execution-profile", mutated))

    def test_t41_sch_016_ordinary_layer_cannot_claim_safety_override(self) -> None:
        artifact = deepcopy(self.artifacts["layered-rule-asset"])
        artifact["can_override_safety_invariants"] = True
        self.assertTrue(schema_errors("layered-rule-asset", artifact))

        artifact = deepcopy(self.artifacts["layered-rule-asset"])
        artifact["rules"] = [
            {
                "rule_id": "RULE-TEST-001",
                "semantic_object_kind": "document",
                "scope": {"scope_kind": "document", "scope_ids": ["document:test"]},
                "confidence": "high",
                "status": "draft",
                "properties": [
                    {
                        "property_id": "security.author-content-immutable",
                        "value": {"type": "boolean", "value": True},
                        "unit_id": None,
                        "mode": "block",
                    }
                ],
            }
        ]
        with self.assertRaisesRegex(ArtifactContractError, "not allowed in layer"):
            validate_artifact(artifact, features={"profile_v2_schema": True})


class RegistryFoundationTests(unittest.TestCase):
    """T41-REG-CORE-001..010 and V041-REGISTRY-CORE-001."""

    def setUp(self) -> None:
        self.test_registry_path = FIXTURES / "property-registry.test.json"

    def test_t41_reg_core_001_registry_schema_and_core_registry_are_valid(self) -> None:
        registry = load_registry()
        self.assertEqual("production", registry["registry_scope"])
        self.assertEqual(1, len(registry["properties"]))

    def test_t41_reg_core_002_committed_catalog_is_mechanically_generated(self) -> None:
        registry = load_registry()
        committed = load_json(GENERATED_CATALOG_PATH)
        committed_typed_values = load_json(GENERATED_TYPED_VALUE_PATH)
        self.assertEqual(build_property_catalog_schema(registry), committed)
        self.assertEqual(build_typed_value_schema(registry), committed_typed_values)
        self.assertEqual(
            {"registry_only": [], "schema_only": []},
            catalog_differences(registry, committed),
        )
        self.assertEqual(
            {"registry_only": [], "schema_only": []},
            typed_value_differences(registry, committed_typed_values),
        )
        verify_committed_catalog(registry)

    def test_t41_reg_core_003_second_property_list_is_detected(self) -> None:
        registry = load_registry()
        catalog = load_json(GENERATED_CATALOG_PATH)
        rogue = deepcopy(catalog["oneOf"][0])
        rogue["properties"]["property_id"]["const"] = "rogue.duplicate-source"
        catalog["oneOf"].append(rogue)
        self.assertEqual(
            ["rogue.duplicate-source"],
            catalog_differences(registry, catalog)["schema_only"],
        )

    def test_t41_reg_core_004_test_properties_are_fixture_only(self) -> None:
        with self.assertRaisesRegex(RegistryContractError, "refuses test-only"):
            load_registry(self.test_registry_path)
        registry = load_registry(self.test_registry_path, allow_test=True)
        self.assertEqual(["test.paragraph-font-size"], [p["property_id"] for p in registry["properties"]])
        catalog = build_property_catalog_schema(registry)
        binding = {
            "property_id": "test.paragraph-font-size",
            "value": {"type": "decimal", "value": "10.50"},
            "unit_id": "unit.pt",
            "mode": "report",
        }
        self.assertFalse(
            list(
                Draft202012Validator(
                    catalog, registry=offline_schema_registry()
                ).iter_errors(binding)
            )
        )

    def test_t41_reg_core_005_unregistered_references_are_rejected(self) -> None:
        mutations = {
            "data_type_id": "string",
            "canonical_unit_id": "unit.mm",
            "normalizer_id": "normalizer.missing",
            "comparator_id": "comparator.missing",
            "executor_capability_id": "executor.missing",
            "auditor_capability_id": "auditor.missing",
        }
        for field, value in mutations.items():
            registry = load_json(self.test_registry_path)
            registry["properties"][0][field] = value
            with self.subTest(field=field), self.assertRaises(RegistryContractError):
                validate_registry_document(registry)

    def test_t41_reg_core_006_callable_and_dynamic_import_ids_are_rejected(self) -> None:
        for value in ("os.system", "normalizer.import", "normalizer.module-loader", "lambda"):
            registry = load_json(self.test_registry_path)
            registry["normalizers"][0]["normalizer_id"] = value
            registry["properties"][0]["normalizer_id"] = value
            with self.subTest(value=value), self.assertRaises(RegistryContractError):
                validate_registry_document(registry)

    def test_t41_reg_core_007_safety_property_is_nonoverridable_and_isolated(self) -> None:
        registry = load_registry()
        property_entry = registry["properties"][0]
        self.assertTrue(property_entry["safety_invariant"])
        self.assertFalse(property_entry["overridable"])
        self.assertEqual(["safety"], property_entry["allowed_layers"])
        mutated = deepcopy(registry)
        mutated["properties"][0]["overridable"] = True
        with self.assertRaises(RegistryContractError):
            validate_registry_document(mutated)

    def test_t41_reg_core_008_duplicate_catalog_ids_are_rejected(self) -> None:
        registry = load_json(self.test_registry_path)
        registry["units"].append(deepcopy(registry["units"][0]))
        with self.assertRaisesRegex(RegistryContractError, "duplicate"):
            validate_registry_document(registry)

    def test_t41_reg_core_009_nested_registry_fields_are_closed(self) -> None:
        registry = load_json(self.test_registry_path)
        registry["properties"][0]["callable"] = "module.function"
        with self.assertRaises(RegistryContractError):
            validate_registry_document(registry)

    def test_t41_reg_core_010_decimal_values_use_strings_not_binary_float(self) -> None:
        registry = load_registry(self.test_registry_path, allow_test=True)
        schema = build_property_catalog_schema(registry)
        binding = {
            "property_id": "test.paragraph-font-size",
            "value": {"type": "decimal", "value": 10.5},
            "unit_id": "unit.pt",
            "mode": "report",
        }
        self.assertTrue(
            list(
                Draft202012Validator(
                    schema, registry=offline_schema_registry()
                ).iter_errors(binding)
            )
        )

    def test_t41_reg_core_002a_typed_value_drift_is_detected_both_directions(self) -> None:
        registry = load_registry()
        committed = load_json(GENERATED_TYPED_VALUE_PATH)

        changed_registry = deepcopy(registry)
        changed_registry["data_types"].append(
            {"data_type_id": "token-string", "json_type": "string", "max_length": 12}
        )
        self.assertEqual(
            ["token-string"],
            typed_value_differences(changed_registry, committed)["registry_only"],
        )
        with self.assertRaisesRegex(RegistryContractError, "data type difference"):
            verify_committed_catalog(changed_registry)

        changed_schema = deepcopy(committed)
        changed_schema["$defs"]["rogue-type"] = deepcopy(changed_schema["$defs"]["string"])
        changed_schema["oneOf"].append({"$ref": "#/$defs/rogue-type"})
        self.assertEqual(
            ["rogue-type"],
            typed_value_differences(registry, changed_schema)["schema_only"],
        )

    def test_t41_reg_core_009a_property_policy_fields_are_required_and_closed(self) -> None:
        required_fields = (
            "missing_strategy",
            "unknown_object_strategy",
            "value_constraints",
            "constraint_ids",
            "test_ids",
        )
        for field in required_fields:
            registry = load_json(self.test_registry_path)
            del registry["properties"][0][field]
            with self.subTest(field=field), self.assertRaises(RegistryContractError):
                validate_registry_document(registry)

        registry = load_json(self.test_registry_path)
        registry["properties"][0]["value_constraints"]["callable"] = "module.function"
        with self.assertRaises(RegistryContractError):
            validate_registry_document(registry)

    def test_t41_reg_core_005a_constraint_references_and_values_are_typed(self) -> None:
        registry = load_json(self.test_registry_path)
        registry["properties"][0]["constraint_ids"] = ["constraint.missing"]
        with self.assertRaisesRegex(RegistryContractError, "unregistered constraints"):
            validate_registry_document(registry)

        registry = load_json(self.test_registry_path)
        second = deepcopy(registry["properties"][0])
        second["property_id"] = "test.paragraph-line-height"
        registry["properties"].append(second)
        registry["constraints"] = [
            {
                "constraint_id": "constraint.import-handler",
                "constraint_kind": "requires",
                "property_ids": [
                    "test.paragraph-font-size",
                    "test.paragraph-line-height",
                ],
                "enforcement": "validator",
                "description": "Synthetic declarative relationship.",
            }
        ]
        with self.assertRaisesRegex(RegistryContractError, "callable-like"):
            validate_registry_document(registry)

        registry = load_json(self.test_registry_path)
        registry["properties"][0]["value_constraints"]["enum_values"] = [True]
        with self.assertRaisesRegex(RegistryContractError, "outside its registered data type"):
            validate_registry_document(registry)


class ResolvedProfileAndConflictContractTests(unittest.TestCase):
    """T41-SCH-005/015 and T41-REG-CORE-002/005 composition data contracts."""

    def setUp(self) -> None:
        self.profile = resolved_final_profile()
        self.artifacts = minimal_artifacts()

    def test_t41_sch_015a_complete_resolved_property_contract_is_valid(self) -> None:
        self.assertFalse(schema_errors("final-execution-profile", self.profile))
        result = validate_artifact(
            self.profile, features={"profile_v2_schema": True}
        )
        self.assertFalse(result.runtime_eligible)

    def test_t41_sch_015b_final_profile_requires_all_fingerprint_bindings(self) -> None:
        for field in (
            "feature_activation_fingerprint",
            "property_registry_fingerprint",
            "profile_fingerprints",
            "approval_fingerprints",
        ):
            profile = deepcopy(self.profile)
            del profile["bindings"][field]
            with self.subTest(field=field):
                self.assertTrue(schema_errors("final-execution-profile", profile))

        for field in ("profile_fingerprints", "approval_fingerprints"):
            profile = deepcopy(self.profile)
            profile["bindings"][field].append(profile["bindings"][field][0])
            with self.subTest(field=field, condition="duplicate"):
                self.assertTrue(schema_errors("final-execution-profile", profile))

    def test_t41_sch_015c_resolved_source_chain_and_nested_fields_are_closed(self) -> None:
        profile = deepcopy(self.profile)
        profile["resolved_properties"][0]["candidate_chain"] = []
        self.assertTrue(schema_errors("final-execution-profile", profile))

        profile = deepcopy(self.profile)
        profile["resolved_properties"][0]["key"]["normalized_scope"]["callable"] = "x.y"
        self.assertTrue(schema_errors("final-execution-profile", profile))

        profile = deepcopy(self.profile)
        profile["resolved_properties"][0]["final_source"]["source_rule_id"] = "OTHER-RULE"
        with self.assertRaisesRegex(ArtifactContractError, "final source"):
            validate_artifact(profile, features={"profile_v2_schema": True})

    def _candidate(self, candidate_id: str, *, layer: str = "safety", scope: str = "applicable") -> dict:
        return {
            "candidate_id": candidate_id,
            "property_binding": {
                "property_id": "security.author-content-immutable",
                "value": {"type": "boolean", "value": True},
                "unit_id": None,
                "mode": "block",
            },
            "source": {
                "source_artifact_id": "layered-rule-asset:safety",
                "source_rule_id": "SAFETY-AUTHOR-CONTENT",
            },
            "layer_kind": layer,
            "confidence": "high",
            "scope_status": scope,
        }

    def _key(self) -> dict:
        return deepcopy(self.profile["resolved_properties"][0]["key"])

    def _conflict(self, *, reason: str = "same_layer") -> dict:
        artifact = deepcopy(self.artifacts["conflict-report"])
        artifact["conflicts"] = [
            {
                "conflict_id": "conflict:test",
                "key": self._key(),
                "reason": reason,
                "status": "blocked_qa",
                "candidates": [
                    self._candidate("candidate:first"),
                    self._candidate("candidate:second"),
                ],
                "excluded_candidates": [],
            }
        ]
        return artifact

    def test_t41_sch_016a_same_layer_conflict_records_sources_and_bindings(self) -> None:
        artifact = self._conflict()
        self.assertFalse(schema_errors("conflict-report", artifact))
        validate_artifact(artifact, features={"profile_v2_schema": True})

    def test_t41_sch_016b_conflict_rejects_unregistered_property_and_cross_layer_confusion(self) -> None:
        artifact = self._conflict()
        artifact["conflicts"][0]["candidates"] = artifact["conflicts"][0]["candidates"][:1]
        with self.assertRaisesRegex(ArtifactContractError, "at least two"):
            validate_artifact(artifact, features={"profile_v2_schema": True})

        artifact = self._conflict()
        artifact["conflicts"][0]["key"]["property_id"] = "unknown.property"
        with self.assertRaisesRegex(ArtifactContractError, "unregistered property"):
            validate_artifact(artifact, features={"profile_v2_schema": True})

        artifact = self._conflict()
        artifact["conflicts"][0]["candidates"][1]["layer_kind"] = "monograph_base"
        with self.assertRaisesRegex(ArtifactContractError, "same_layer"):
            validate_artifact(artifact, features={"profile_v2_schema": True})

    def test_t41_sch_016c_scope_violation_requires_a_declarative_exclusion(self) -> None:
        artifact = self._conflict(reason="scope_violation")
        artifact["conflicts"][0]["candidates"] = [self._candidate("candidate:active")]
        with self.assertRaisesRegex(ArtifactContractError, "scope-violation exclusion"):
            validate_artifact(artifact, features={"profile_v2_schema": True})

        artifact["conflicts"][0]["excluded_candidates"] = [
            {
                "candidate": self._candidate(
                    "candidate:excluded", scope="out_of_scope"
                ),
                "exclusion_reason": "scope_violation",
                "reason_code": "SCOPE-OUTSIDE-SECTION",
            }
        ]
        validate_artifact(artifact, features={"profile_v2_schema": True})


class VersionDispatchAndRegressionTests(unittest.TestCase):
    """T41-SCH-017..024, V041-LEGACY-001, and V041-GOV-001."""

    def test_t41_sch_017_legacy_10_and_11_are_read_only_disabled(self) -> None:
        for version in ("1.0", "1.1"):
            result = read_profile_document(legacy_profile(version))
            with self.subTest(version=version):
                self.assertTrue(result.legacy_input)
                self.assertTrue(result.read_only)
                self.assertEqual("disabled", result.activation)
                self.assertFalse(result.runtime_eligible)

    def test_t41_sch_018_legacy_cannot_enter_v2_contract(self) -> None:
        result = read_profile_document(legacy_profile("1.1"))
        with self.assertRaisesRegex(ArtifactContractError, "cannot enter V2"):
            require_v2_nonlegacy_contract(result)

    def test_t41_sch_019_unknown_major_and_minor_versions_are_rejected(self) -> None:
        artifact = minimal_artifacts()["capability-snapshot"]
        for version in ("3.0", "2.1"):
            mutated = deepcopy(artifact)
            mutated["schema_version"] = version
            with self.subTest(version=version), self.assertRaises(ArtifactContractError):
                validate_artifact(mutated, features={"profile_v2_schema": True})
        with self.assertRaises(ArtifactContractError):
            read_profile_document(legacy_profile("1.0") | {"schema_version": "1.2"})

    def test_t41_sch_020_declared_minor_is_read_only_and_still_closed(self) -> None:
        artifact = minimal_artifacts()["capability-snapshot"]
        artifact["schema_version"] = "2.1"
        schema = load_artifact_schema("capability-snapshot")
        schema["x-read-compatible-minor-versions"] = ["2.1"]
        effective, read_only = schema_for_requested_minor(schema, "2.1")
        self.assertTrue(read_only)
        result = validate_artifact(
            artifact,
            features={"profile_v2_schema": True},
            schema_override=schema,
        )
        self.assertTrue(result.read_only)
        artifact["unknown_minor_field"] = True
        self.assertTrue(schema_errors("capability-snapshot", artifact, schema_override=effective))

    def test_t41_sch_021_legacy_validator_behavior_is_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            for version in ("1.0", "1.1"):
                path = Path(temp_dir) / f"legacy-{version}.json"
                path.write_text(json.dumps(legacy_profile(version)), encoding="utf-8")
                errors, parsed = validate_legacy_path(path)
                with self.subTest(version=version):
                    self.assertFalse(errors)
                    self.assertEqual(version, parsed["schema_version"])

    def test_t41_sch_022_v2_reader_does_not_enable_runtime(self) -> None:
        artifact = minimal_artifacts()["final-execution-profile"]
        result = read_profile_document(
            artifact, features={"profile_v2_schema": True}
        )
        self.assertFalse(result.runtime_eligible)
        self.assertEqual("disabled", result.activation)

    def test_t41_sch_023_no_p2_to_p7_feature_is_declared(self) -> None:
        artifact = minimal_artifacts()["feature-activation-manifest"]
        self.assertEqual(
            {"profile_v2_schema", "final_ready_eligible"},
            set(artifact["features"]),
        )
        for forbidden in (
            "profile_v2_composer",
            "monograph_base_v041",
            "v041_basic_execution",
            "v041_word_finalize",
        ):
            mutated = deepcopy(artifact)
            mutated["features"][forbidden] = False
            with self.subTest(feature=forbidden):
                self.assertTrue(schema_errors("feature-activation-manifest", mutated))

    def test_t41_sch_024_p1_files_contain_no_private_paths_or_manuscript_markers(self) -> None:
        files = list(SCHEMA_DIR.glob("*")) + [
            SCRIPTS / "profile_v2_registry.py",
            SCRIPTS / "profile_v2_artifacts.py",
            FIXTURES / "minimal-artifacts.json",
            FIXTURES / "property-registry.test.json",
            FIXTURES / "resolved-final-profile.json",
        ]
        forbidden_patterns = (
            re.compile(r"(?<![A-Za-z0-9])[A-Za-z]:[\\/]"),
            re.compile(r"/Users/|/home/"),
            re.compile(r"BEGIN (?:RSA |OPENSSH )?PRIVATE KEY"),
            re.compile(r"(?:ghp_|github_pat_)[A-Za-z0-9_]+"),
        )
        for path in files:
            if not path.is_file():
                continue
            value = path.read_text(encoding="utf-8")
            for pattern in forbidden_patterns:
                with self.subTest(path=path.name, pattern=pattern.pattern):
                    self.assertIsNone(pattern.search(value))


if __name__ == "__main__":
    unittest.main()
