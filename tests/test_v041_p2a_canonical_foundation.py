from __future__ import annotations

import json
import random
import re
import sys
import unittest
from copy import deepcopy
from pathlib import Path
from urllib.parse import urldefrag, urljoin

from jsonschema import Draft202012Validator, FormatChecker


REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "format-monograph" / "scripts"
FIXTURES = REPO / "tests" / "fixtures" / "v041"
sys.path.insert(0, str(SCRIPTS))

import profile_v2_artifacts as artifacts  # noqa: E402
from profile_v2_artifacts import (  # noqa: E402
    ArtifactContractError,
    _validate_artifact_for_test,
    artifact_semantic_errors,
    load_artifact_schema,
    offline_schema_registry,
    read_profile_document,
    schema_documents,
    validate_artifact,
)
from profile_v2_canonical import (  # noqa: E402
    CanonicalizationError,
    _merge_schema,
    _semantic_projection,
    _stamp_semantic_fingerprint_for_test,
    canonical_semantic_bytes,
    compute_semantic_fingerprint,
    fingerprint_field_inventory,
    fingerprint_field_node,
    schema_node_fingerprint,
    semantic_projection,
    stamp_semantic_fingerprint,
    unclassified_array_paths,
    verify_semantic_fingerprint,
)
from profile_v2_registry import (  # noqa: E402
    RegistryContractError,
    build_property_catalog_schema,
    build_typed_value_schema,
    load_registry,
    validate_registry_document,
    verify_committed_catalog,
)
from profile_v2_scope import (  # noqa: E402
    ScopeContractError,
    normalize_scope,
    normalized_property_scope_key,
    scope_disjoint,
    scope_equal,
    scope_overlap_state,
    scope_subset,
    validate_module_asset_scope,
)
from profile_v2_values import (  # noqa: E402
    ValueNormalizationError,
    compare_property_bindings,
    normalize_property_binding,
)


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def minimal_artifacts() -> dict[str, dict]:
    return load_json(FIXTURES / "minimal-artifacts.json")


def test_registry() -> dict:
    return load_registry(
        FIXTURES / "property-registry.v2.1.test.json",
        allow_test=True,
        version="2.1",
    )


def document_scope(*ids: str) -> dict:
    return {
        "selectors": [{"selector_kind": "document", "selector_ids": list(ids)}],
        "exclusions": [],
        "mutually_exclusive_conditions": [],
    }


def layered_v21(*, value: str = "1", unit_id: str = "unit.mm", layer: str = "monograph_base") -> dict:
    return {
        "artifact_kind": "layered-rule-asset",
        "schema_version": "2.1",
        "artifact_id": "layered-rule-asset:test-p2a",
        "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
        "input_fingerprints": [
            {
                "input_id": "input:test",
                "role": "rule_asset",
                "fingerprint": "sha256:" + ("0" * 64),
            }
        ],
        "semantic_fingerprint": "sha256:" + ("0" * 64),
        "layer_kind": layer,
        "can_override_safety_invariants": False,
        "activation": "disabled",
        "asset_scope": document_scope("document:test"),
        "allowed_semantic_object_kinds": ["paragraph"],
        "rules": [
            {
                "rule_id": "RULE-P2A-001",
                "semantic_object_kind": "paragraph",
                "scope": document_scope("document:test"),
                "confidence": "high",
                "status": "draft",
                "properties": [
                    {
                        "property_id": "test.paragraph-font-size",
                        "value": {"type": "decimal", "value": value},
                        "unit_id": unit_id,
                        "mode": "report",
                    }
                ],
            }
        ],
    }


def stamp_test(document: dict, registry: dict | None = None) -> dict:
    registry = registry or test_registry()
    schema = load_artifact_schema(
        document["artifact_kind"], version=document["schema_version"]
    )
    documents = artifacts._schema_documents(artifacts._test_schema_overrides(registry))
    return _stamp_semantic_fingerprint_for_test(
        document, schema=schema, documents=documents, registry=registry
    )


def _schema_example(
    schema: dict,
    *,
    schema_id: str,
    documents: dict[str, dict],
    variant: int = 0,
    depth: int = 0,
) -> object:
    """Build a small deterministic value for an actual committed schema node."""

    if depth > 24:
        raise AssertionError(f"Schema example recursion exceeded at {schema_id}")
    current = deepcopy(schema)
    if "$ref" in current:
        reference = urljoin(schema_id, current.pop("$ref"))
        target_uri, fragment = urldefrag(reference)
        target = fingerprint_field_node(documents, f"{target_uri}#{fragment}")
        merged = _merge_schema(target, current)
        return _schema_example(
            merged,
            schema_id=target_uri,
            documents=documents,
            variant=variant,
            depth=depth + 1,
        )
    if "oneOf" in current:
        branches = current.pop("oneOf")
        branch = deepcopy(branches[variant % len(branches)])
        branch = _merge_schema(branch, current)
        return _schema_example(
            branch,
            schema_id=schema_id,
            documents=documents,
            variant=variant,
            depth=depth + 1,
        )
    if "const" in current:
        return deepcopy(current["const"])
    enum = current.get("enum")
    if isinstance(enum, list) and enum:
        return deepcopy(enum[variant % len(enum)])

    value_type = current.get("type")
    if isinstance(value_type, list):
        value_type = next((item for item in value_type if item != "null"), "null")
    if value_type == "null":
        return None
    if value_type == "boolean":
        return bool(variant % 2)
    if value_type == "integer":
        minimum = current.get("minimum", 0)
        if "exclusiveMinimum" in current:
            minimum = max(minimum, current["exclusiveMinimum"] + 1)
        value = int(minimum) + variant
        maximum = current.get("maximum")
        if maximum is not None:
            value = min(value, int(maximum))
        return value
    if value_type == "string" or value_type is None:
        if current.get("format") == "date-time":
            return f"2026-01-{(variant % 9) + 1:02d}T00:00:00Z"
        pattern = current.get("pattern")
        candidates = [
            f"test.value{variant}",
            f"RULE-TEST-{variant:03d}",
            f"test_value{variant}",
            f"1.{variant}",
            f"1.0.{variant}",
            f"test-value{variant}",
            "sha256:" + (f"{variant:x}"[-1] * 64),
            f"item:test{variant}",
            "scope:" + (f"{variant:x}"[-1] * 64),
            f"unit.test{variant}",
            f"constraint.test{variant}",
            f"auditor.test{variant}",
            f"executor.test{variant}",
            f"comparator.test{variant}",
            f"normalizer.test{variant}",
            f"T41-TEST-{variant}",
            f"registry:test{variant}",
            str(variant),
        ]
        if isinstance(pattern, str):
            for candidate in candidates:
                if re.fullmatch(pattern, candidate):
                    return candidate
            raise AssertionError(f"No deterministic example matches {pattern!r}")
        minimum = int(current.get("minLength", 1))
        value = f"value-{variant}"
        return value if len(value) >= minimum else ("x" * minimum)
    if value_type == "array":
        count = int(current.get("minItems", 0))
        item_schema = current.get("items", {})
        return [
            _schema_example(
                item_schema,
                schema_id=schema_id,
                documents=documents,
                variant=variant + index,
                depth=depth + 1,
            )
            for index in range(count)
        ]
    if value_type == "object":
        properties = current.get("properties", {})
        result = {}
        for index, name in enumerate(current.get("required", [])):
            result[name] = _schema_example(
                properties[name],
                schema_id=schema_id,
                documents=documents,
                variant=variant + index,
                depth=depth + 1,
            )
        return result
    raise AssertionError(f"Unsupported schema example type {value_type!r} at {schema_id}")


def _set_pointer(document: dict, pointer: str, value: object) -> None:
    current: object = document
    parts = pointer.removeprefix("/").split("/") if pointer else []
    for raw in parts[:-1]:
        token = raw.replace("~1", "/").replace("~0", "~")
        current = current[int(token)] if isinstance(current, list) else current[token]
    final = parts[-1].replace("~1", "/").replace("~0", "~")
    if isinstance(current, list):
        current[int(final)] = value
    else:
        current[final] = value


def candidate_v21() -> dict:
    return {
        "candidate_id": "candidate:p2a-font-size",
        "property_binding": {
            "property_id": "test.paragraph-font-size",
            "value": {"type": "decimal", "value": "2.83"},
            "unit_id": "unit.pt",
            "mode": "report",
        },
        "source": {
            "source_artifact_id": "layered-rule-asset:test-p2a",
            "source_rule_id": "RULE-P2A-001",
        },
        "layer_kind": "monograph_base",
        "confidence": "high",
        "scope_status": "applicable",
    }


def conflict_v21() -> dict:
    scope = normalize_scope(document_scope("document:test"))
    return {
        "artifact_kind": "conflict-report",
        "schema_version": "2.1",
        "artifact_id": "conflict-report:test-p2a",
        "created_by_tool": {"tool_id": "format-monograph", "version": "0.4.1"},
        "input_fingerprints": [
            {"input_id": "input:test", "role": "conflict_report", "fingerprint": "sha256:" + ("0" * 64)}
        ],
        "semantic_fingerprint": "sha256:" + ("0" * 64),
        "generated_at": "2026-01-01T00:00:00Z",
        "conflicts": [
            {
                "conflict_id": "conflict:p2a-font-size",
                "key": {
                    "semantic_object_kind": "paragraph",
                    "property_id": "test.paragraph-font-size",
                    "normalized_scope": scope,
                },
                "reason": "low_confidence",
                "status": "blocked_qa",
                "candidates": [candidate_v21()],
                "excluded_candidates": [],
            }
        ],
    }


def final_profile_v21() -> dict:
    document = deepcopy(minimal_artifacts()["final-execution-profile"])
    document["schema_version"] = "2.1"
    scope = normalize_scope(document_scope("document:test"))
    candidate = candidate_v21()
    document["resolved_properties"] = [
        {
            "resolution_id": "resolution:p2a-font-size",
            "key": {
                "semantic_object_kind": "paragraph",
                "property_id": "test.paragraph-font-size",
                "normalized_scope": scope,
            },
            "resolved_binding": deepcopy(candidate["property_binding"]),
            "final_layer_kind": "monograph_base",
            "final_source": deepcopy(candidate["source"]),
            "candidate_chain": [candidate],
            "override_chain": ["candidate:p2a-font-size"],
            "excluded_candidates": [],
            "confidence": "high",
            "conflict_id": None,
            "qa_decision_id": None,
            "safety_check": {
                "status": "pass",
                "checked_invariant_ids": ["test.safety-author-content-immutable"],
            },
            "execution_mode": "report",
        }
    ]
    return document


class VersionedContractTests(unittest.TestCase):
    def test_t41_p2a_ver_001_explicit_routes_do_not_fake_minor_versions(self) -> None:
        for kind in artifacts.ARTIFACT_KINDS:
            self.assertEqual(
                "2.0",
                load_artifact_schema(kind, version="2.0")["properties"]["schema_version"]["const"],
            )
        for kind in ("layered-rule-asset", "conflict-report", "final-execution-profile"):
            self.assertEqual(
                "2.1",
                load_artifact_schema(kind, version="2.1")["properties"]["schema_version"]["const"],
            )
        with self.assertRaises(ArtifactContractError):
            load_artifact_schema("capability-snapshot", version="2.1")
        with self.assertRaises(ArtifactContractError):
            load_artifact_schema("layered-rule-asset", version="2.2")

    def test_t41_p2a_ver_002_v20_and_v21_remain_disabled(self) -> None:
        v20 = minimal_artifacts()["final-execution-profile"]
        result20 = read_profile_document(v20, features={"profile_v2_schema": True})
        self.assertFalse(result20.runtime_eligible)
        self.assertEqual("disabled", result20.activation)

        v21 = deepcopy(v20)
        v21["schema_version"] = "2.1"
        v21 = stamp_semantic_fingerprint(v21)
        result21 = validate_artifact(v21, features={"profile_v2_schema": True})
        self.assertFalse(result21.runtime_eligible)
        self.assertEqual("disabled", result21.activation)

    def test_t41_p2a_ver_003_p1_placeholder_fingerprint_is_not_readable(self) -> None:
        artifact = deepcopy(minimal_artifacts()["capability-snapshot"])
        artifact["semantic_fingerprint"] = "sha256:" + ("0" * 64)
        with self.assertRaisesRegex(ArtifactContractError, "semantic_fingerprint"):
            read_profile_document(artifact, features={"profile_v2_schema": True})

    def test_t41_p2a_sch_001_every_supported_array_is_classified(self) -> None:
        self.assertEqual([], unclassified_array_paths(schema_documents()))

    def test_t41_p2a_reg_001_both_generated_catalogs_match_registry(self) -> None:
        verify_committed_catalog(load_registry(), version="2.0")
        verify_committed_catalog(load_registry(version="2.1"), version="2.1")
        self.assertEqual(
            "https://schemas.format-monograph.local/v2.1/property-catalog.generated.schema.json",
            build_property_catalog_schema(load_registry(version="2.1"))["$id"],
        )
        self.assertEqual(
            "https://schemas.format-monograph.local/v2.1/typed-value.generated.schema.json",
            build_typed_value_schema(load_registry(version="2.1"))["$id"],
        )


class TestRegistryIsolationTests(unittest.TestCase):
    def test_t41_p2a_reg_002_test_registry_reuses_every_nonproperty_catalog(self) -> None:
        registry = test_registry()
        document = stamp_test(layered_v21(), registry)
        mutations = []

        executor_drift = deepcopy(registry)
        executor_drift["executor_capabilities"][0]["description"] = "Test drift."
        mutations.append(("executor_capabilities", executor_drift))

        auditor_drift = deepcopy(registry)
        auditor_drift["auditor_capabilities"][0]["description"] = "Test drift."
        mutations.append(("auditor_capabilities", auditor_drift))

        constraint_drift = deepcopy(registry)
        constraint_drift["constraints"].append(
            {
                "constraint_id": "constraint.test-drift",
                "constraint_kind": "requires",
                "property_ids": [
                    "test.paragraph-font-size",
                    "test.table-column-count",
                ],
                "enforcement": "qa",
                "description": "Synthetic catalog drift.",
            }
        )
        mutations.append(("constraints", constraint_drift))

        for collection, mutated in mutations:
            with self.subTest(collection=collection), self.assertRaisesRegex(
                ArtifactContractError, f"production {collection} catalog exactly"
            ):
                _validate_artifact_for_test(
                    document,
                    registry=mutated,
                    features={"profile_v2_schema": True},
                )


class CanonicalFingerprintTests(unittest.TestCase):
    def setUp(self) -> None:
        self.artifacts = minimal_artifacts()
        self.golden = load_json(FIXTURES / "canonical-golden-vectors.json")

    def test_t41_p2a_can_001_golden_vector_is_stable(self) -> None:
        artifact = self.artifacts["feature-activation-manifest"]
        self.assertEqual(
            self.golden["feature_activation_fingerprint"],
            compute_semantic_fingerprint(artifact),
        )
        payload = canonical_semantic_bytes(artifact)
        self.assertFalse(payload.startswith(b"\xef\xbb\xbf"))
        self.assertFalse(payload.endswith(b"\n"))
        self.assertEqual(payload, canonical_semantic_bytes(semantic_projection(artifact)))

    def test_t41_p2a_can_002_set_order_is_stable_and_ordered_history_is_not(self) -> None:
        profile = deepcopy(self.artifacts["final-execution-profile"])
        shuffled = deepcopy(profile)
        random.Random(41).shuffle(shuffled["input_fingerprints"])
        self.assertEqual(
            compute_semantic_fingerprint(profile), compute_semantic_fingerprint(shuffled)
        )

        evidence = deepcopy(self.artifacts["execution-evidence-artifact"])
        evidence["history"].append(
            {
                "entry_id": "history:reviewed",
                "event_type": "reviewed",
                "created_at": "2026-01-02T00:00:00Z",
                "reason": "Second synthetic event.",
            }
        )
        reversed_history = deepcopy(evidence)
        reversed_history["history"].reverse()
        self.assertNotEqual(
            compute_semantic_fingerprint(evidence),
            compute_semantic_fingerprint(reversed_history),
        )

        reversed_keys = {key: profile[key] for key in reversed(list(profile))}
        self.assertEqual(
            canonical_semantic_bytes(profile), canonical_semantic_bytes(reversed_keys)
        )

    def test_t41_p2a_can_003_nfc_strings_converge_and_key_collisions_fail(self) -> None:
        composed = deepcopy(self.artifacts["qa-approval-artifact"])
        decomposed = deepcopy(composed)
        composed["reason"] = "Caf\u00e9"
        decomposed["reason"] = "Cafe\u0301"
        self.assertEqual(
            compute_semantic_fingerprint(composed),
            compute_semantic_fingerprint(decomposed),
        )
        collision = deepcopy(composed)
        collision["created_by_tool"]["\u00e9"] = "one"
        collision["created_by_tool"]["e\u0301"] = "two"
        with self.assertRaisesRegex(CanonicalizationError, "key collision"):
            compute_semantic_fingerprint(collision)

    def test_t41_p2a_can_004_only_schema_exclusions_are_ignored(self) -> None:
        snapshot = deepcopy(self.artifacts["capability-snapshot"])
        changed_time = deepcopy(snapshot)
        changed_time["captured_at"] = "2030-12-31T23:59:59Z"
        self.assertEqual(
            compute_semantic_fingerprint(snapshot),
            compute_semantic_fingerprint(changed_time),
        )
        changed_semantics = deepcopy(snapshot)
        changed_semantics["environment"]["python_version"] = "3.13.0"
        self.assertNotEqual(
            compute_semantic_fingerprint(snapshot),
            compute_semantic_fingerprint(changed_semantics),
        )

        excluded_time_fields = [
            ("qa-approval-artifact", ("created_at",)),
            ("conflict-report", ("generated_at",)),
            ("execution-evidence-artifact", ("created_at",)),
            ("execution-evidence-artifact", ("history", 0, "created_at")),
        ]
        for kind, path in excluded_time_fields:
            original = deepcopy(self.artifacts[kind])
            changed = deepcopy(original)
            target = changed
            for part in path[:-1]:
                target = target[part]
            target[path[-1]] = "2031-01-01T00:00:00Z"
            with self.subTest(kind=kind, path=path):
                self.assertEqual(
                    compute_semantic_fingerprint(original),
                    compute_semantic_fingerprint(changed),
                )

        delegated = deepcopy(self.artifacts["qa-approval-artifact"])
        delegated["approver"] = {
            "actor_id": "actor:publisher",
            "actor_role": "delegated_publisher",
            "authorization_reference": {
                "authorization_id": "authorization:test",
                "granted_by_actor_id": "actor:user",
                "authority_scope": {
                    "scope_kind": "document",
                    "scope_ids": ["document:test"],
                },
                "issued_at": "2026-01-01T00:00:00Z",
            },
        }
        changed_issued = deepcopy(delegated)
        changed_issued["approver"]["authorization_reference"]["issued_at"] = (
            "2032-01-01T00:00:00Z"
        )
        self.assertEqual(
            compute_semantic_fingerprint(delegated),
            compute_semantic_fingerprint(changed_issued),
        )

    def test_t41_p2a_can_005_self_reference_is_excluded_but_verify_detects_stale(self) -> None:
        artifact = deepcopy(self.artifacts["feature-activation-manifest"])
        original = compute_semantic_fingerprint(artifact)
        artifact["semantic_fingerprint"] = "sha256:" + ("f" * 64)
        self.assertEqual(original, compute_semantic_fingerprint(artifact))
        with self.assertRaises(CanonicalizationError):
            verify_semantic_fingerprint(artifact)

        unstamped = deepcopy(self.artifacts["feature-activation-manifest"])
        unstamped["semantic_fingerprint"] = "sha256:" + ("0" * 64)
        stamped = stamp_semantic_fingerprint(unstamped)
        self.assertEqual("sha256:" + ("0" * 64), unstamped["semantic_fingerprint"])
        verify_semantic_fingerprint(stamped)

        stale = deepcopy(self.artifacts["capability-snapshot"])
        stale["capabilities"][0]["available"] = False
        with self.assertRaisesRegex(ArtifactContractError, "semantic_fingerprint"):
            validate_artifact(stale, features={"profile_v2_schema": True})

    def test_t41_p2a_can_006_semantic_identity_and_bindings_are_sensitive(self) -> None:
        base = self.artifacts["final-execution-profile"]
        mutations = [
            ("artifact_id", "final-execution-profile:other"),
            ("task_id", "task:other"),
        ]
        for field, value in mutations:
            changed = deepcopy(base)
            changed[field] = value
            with self.subTest(field=field):
                self.assertNotEqual(
                    compute_semantic_fingerprint(base),
                    compute_semantic_fingerprint(changed),
                )
        changed_input = deepcopy(base)
        changed_input["input_fingerprints"][0]["fingerprint"] = "sha256:" + ("a" * 64)
        changed_binding = deepcopy(base)
        changed_binding["bindings"]["input_fingerprint"] = "sha256:" + ("b" * 64)
        self.assertNotEqual(compute_semantic_fingerprint(base), compute_semantic_fingerprint(changed_input))
        self.assertNotEqual(compute_semantic_fingerprint(base), compute_semantic_fingerprint(changed_binding))

    def test_t41_p2a_can_007_binary_float_is_rejected(self) -> None:
        artifact = deepcopy(self.artifacts["capability-snapshot"])
        artifact["capabilities"][0]["available"] = 1.0
        with self.assertRaisesRegex(CanonicalizationError, "floating-point"):
            compute_semantic_fingerprint(artifact)

    def test_t41_p2a_can_008_schema_field_inventory_is_machine_locked(self) -> None:
        expected = load_json(FIXTURES / "semantic-fingerprint-field-inventory.json")
        self.assertEqual("2.0", expected["schema_version"])
        self.assertEqual(
            {
                "semantic_projected": "T41-P2A-CAN-009",
                "semantic_guarded": "T41-P2A-CAN-010",
                "fingerprint_excluded": "T41-P2A-CAN-011",
                "validation_guard": "T41-P2A-CAN-010",
            },
            expected["classification_tests"],
        )
        documents = schema_documents()
        actual = fingerprint_field_inventory(documents)
        self.assertEqual(expected["inventory"], actual)
        self.assertEqual(expected["counts"]["schema_property_nodes"], len(actual))
        self.assertEqual(len(actual), len({record["field_path"] for record in actual}))
        for classification in expected["classification_tests"]:
            self.assertEqual(
                expected["counts"][classification],
                sum(
                    record["classification"] == classification for record in actual
                ),
            )

        for record in actual:
            with self.subTest(field_path=record["field_path"]):
                node = fingerprint_field_node(documents, record["field_path"])
                self.assertEqual(
                    record["schema_node_fingerprint"], schema_node_fingerprint(node)
                )
                self.assertEqual(
                    record["field_name"],
                    record["field_path"].rsplit("/properties/", 1)[-1]
                    .replace("~1", "/")
                    .replace("~0", "~"),
                )

        bad_path = deepcopy(actual[0])
        bad_path["field_path"] += "/missing"
        with self.assertRaises(CanonicalizationError):
            fingerprint_field_node(documents, bad_path["field_path"])

        added = deepcopy(documents)
        root_id = "https://schemas.format-monograph.local/v2/capability-snapshot.schema.json"
        added[root_id]["properties"]["unclassified_new_field"] = {"type": "string"}
        self.assertNotEqual(expected["inventory"], fingerprint_field_inventory(added))

        reclassified = deepcopy(documents)
        projected = next(
            record for record in actual if record["classification"] == "semantic_projected"
        )
        fingerprint_field_node(reclassified, projected["field_path"])[
            "x-semantic-fingerprint"
        ] = "exclude"
        self.assertNotEqual(
            expected["inventory"], fingerprint_field_inventory(reclassified)
        )

    def test_t41_p2a_can_009_actual_included_nodes_project_their_field(self) -> None:
        documents = schema_documents()
        inventory = fingerprint_field_inventory(documents)
        for record in inventory:
            if record["classification"] != "semantic_projected":
                continue
            with self.subTest(field_path=record["field_path"]):
                schema_id, _ = urldefrag(record["field_path"])
                container = deepcopy(
                    fingerprint_field_node(documents, record["container_path"])
                )
                container["$schema"] = "https://json-schema.org/draft/2020-12/schema"
                container["$id"] = schema_id
                node = fingerprint_field_node(documents, record["field_path"])
                sample = _schema_example(
                    node, schema_id=schema_id, documents=documents
                )
                registry = load_registry(version=record["schema_version"])
                first = _semantic_projection(
                    {}, schema=container, documents=documents, registry=registry
                )
                second = _semantic_projection(
                    {record["field_name"]: sample},
                    schema=container,
                    documents=documents,
                    registry=registry,
                )
                self.assertNotEqual(first, second)
                self.assertIn(record["field_name"], second)

    def test_t41_p2a_can_010_guarded_nodes_use_their_real_schema_constraint(self) -> None:
        documents = schema_documents()
        inventory = fingerprint_field_inventory(documents)
        registry = offline_schema_registry()
        guarded = [
            record
            for record in inventory
            if record["classification"] in {"semantic_guarded", "validation_guard"}
        ]
        mechanism_map = load_json(
            FIXTURES / "semantic-fingerprint-field-inventory.json"
        )["guard_mechanism_tests"]
        for record in guarded:
            with self.subTest(field_path=record["field_path"]):
                schema_id, _ = urldefrag(record["field_path"])
                node = deepcopy(fingerprint_field_node(documents, record["field_path"]))
                for reason in record["guard_reasons"]:
                    self.assertIn(reason, mechanism_map)
                if record["classification"] == "validation_guard":
                    self.assertIn("validation_guard", record["schema_features"])
                    continue
                if not ({"const", "single_enum"} & set(record["guard_reasons"])):
                    self.assertTrue(
                        {
                            "normalized_scope_context",
                            "property_binding_context",
                        }
                        & set(record["guard_reasons"])
                    )
                    continue
                validation_schema = {
                    "$schema": "https://json-schema.org/draft/2020-12/schema",
                    "$id": schema_id,
                    **node,
                }
                baseline = _schema_example(
                    node, schema_id=schema_id, documents=documents
                )
                if "const" in node:
                    invalid = not baseline if isinstance(baseline, bool) else f"{baseline}-invalid"
                else:
                    invalid = f"{baseline}-invalid"
                validator = Draft202012Validator(
                    validation_schema,
                    registry=registry,
                    format_checker=FormatChecker(),
                )
                self.assertTrue(validator.is_valid(baseline))
                self.assertFalse(validator.is_valid(invalid))

        approval = deepcopy(self.artifacts["qa-approval-artifact"])
        delegated_without_authorization = deepcopy(approval)
        delegated_without_authorization["approver"]["actor_role"] = "delegated_publisher"
        validator = Draft202012Validator(
            load_artifact_schema("qa-approval-artifact"),
            registry=offline_schema_registry(),
            format_checker=FormatChecker(),
        )
        self.assertTrue(validator.is_valid(approval))
        self.assertFalse(validator.is_valid(delegated_without_authorization))

    def test_t41_p2a_can_011_actual_excluded_nodes_do_not_project_their_field(self) -> None:
        documents = schema_documents()
        inventory = fingerprint_field_inventory(documents)
        excluded = [
            record
            for record in inventory
            if record["classification"] == "fingerprint_excluded"
        ]
        for record in excluded:
            with self.subTest(field_path=record["field_path"]):
                schema_id, _ = urldefrag(record["field_path"])
                container = deepcopy(
                    fingerprint_field_node(documents, record["container_path"])
                )
                container["$schema"] = "https://json-schema.org/draft/2020-12/schema"
                container["$id"] = schema_id
                node = fingerprint_field_node(documents, record["field_path"])
                sample = _schema_example(
                    node, schema_id=schema_id, documents=documents
                )
                registry = load_registry(version=record["schema_version"])
                first = _semantic_projection(
                    {}, schema=container, documents=documents, registry=registry
                )
                second = _semantic_projection(
                    {record["field_name"]: sample},
                    schema=container,
                    documents=documents,
                    registry=registry,
                )
                self.assertEqual(first, second)

    def test_t41_p2a_can_012_projection_mechanisms_are_real_and_closed(self) -> None:
        expected = load_json(FIXTURES / "semantic-fingerprint-field-inventory.json")
        inventory = fingerprint_field_inventory(schema_documents())
        actual_features = sorted(
            {feature for record in inventory for feature in record["schema_features"]}
        )
        self.assertEqual(
            sorted(expected["projection_mechanism_tests"]), actual_features
        )
        self.assertTrue(any("ref" in record["schema_features"] for record in inventory))
        self.assertTrue(any("one_of" in record["schema_features"] for record in inventory))
        self.assertTrue(
            any("generated_schema" in record["schema_features"] for record in inventory)
        )
        unsupported = deepcopy(schema_documents())
        root_id = "https://schemas.format-monograph.local/v2/capability-snapshot.schema.json"
        unsupported[root_id]["properties"]["unsupported_composition"] = {
            "anyOf": [{"type": "string"}]
        }
        with self.assertRaisesRegex(CanonicalizationError, "does not support"):
            fingerprint_field_inventory(unsupported)

    def test_t41_p2a_can_013_key_artifact_end_to_end_mutations_are_real(self) -> None:
        evidence = load_json(FIXTURES / "semantic-fingerprint-field-inventory.json")
        inventory_paths = {
            record["field_path"] for record in evidence["inventory"]
        }
        factories = {
            "layered_v21": layered_v21,
            "conflict_v21": conflict_v21,
            "final_profile_v21": final_profile_v21,
        }
        for case in evidence["end_to_end_cases"]:
            with self.subTest(case_id=case["case_id"]):
                self.assertIn(case["schema_field_path"], inventory_paths)
                if case["document_factory"] == "minimal_artifact":
                    baseline = deepcopy(self.artifacts[case["artifact_kind"]])
                    changed = deepcopy(baseline)
                    _set_pointer(changed, case["instance_path"], case["replacement"])
                    baseline = stamp_semantic_fingerprint(baseline)
                    changed = stamp_semantic_fingerprint(changed)
                    validate_artifact(baseline, features={"profile_v2_schema": True})
                    validate_artifact(changed, features={"profile_v2_schema": True})
                else:
                    baseline = factories[case["document_factory"]]()
                    changed = deepcopy(baseline)
                    _set_pointer(changed, case["instance_path"], case["replacement"])
                    baseline = stamp_test(baseline)
                    changed = stamp_test(changed)
                    _validate_artifact_for_test(
                        baseline,
                        registry=test_registry(),
                        features={"profile_v2_schema": True},
                    )
                    _validate_artifact_for_test(
                        changed,
                        registry=test_registry(),
                        features={"profile_v2_schema": True},
                    )
                self.assertNotEqual(
                    baseline["semantic_fingerprint"], changed["semantic_fingerprint"]
                )


class ExactUnitTests(unittest.TestCase):
    def setUp(self) -> None:
        self.registry = test_registry()
        self.entry_id = "test.paragraph-font-size"

    def binding(self, value: str, unit_id: str = "unit.pt") -> dict:
        return {
            "property_id": self.entry_id,
            "value": {"type": "decimal", "value": value},
            "unit_id": unit_id,
            "mode": "report",
        }

    def test_t41_p2a_unit_001_exact_mm_conversion_and_half_even(self) -> None:
        one_mm = normalize_property_binding(self.binding("1", "unit.mm"), self.registry)
        self.assertEqual("2.83", one_mm["value"]["value"])
        self.assertEqual("unit.pt", one_mm["unit_id"])
        self.assertEqual(
            "2.34",
            normalize_property_binding(self.binding("2.345"), self.registry)["value"]["value"],
        )
        self.assertEqual(
            "2.36",
            normalize_property_binding(self.binding("2.355"), self.registry)["value"]["value"],
        )

    def test_t41_p2a_unit_002_equivalent_units_have_identical_artifact_fingerprint(self) -> None:
        millimetres = stamp_test(layered_v21(value="1", unit_id="unit.mm"), self.registry)
        points = stamp_test(layered_v21(value="2.83", unit_id="unit.pt"), self.registry)
        self.assertEqual(
            millimetres["semantic_fingerprint"], points["semantic_fingerprint"]
        )
        self.assertTrue(
            compare_property_bindings(
                millimetres["rules"][0]["properties"][0],
                points["rules"][0]["properties"][0],
                self.registry,
            )
        )

    def test_t41_p2a_unit_003_negative_zero_and_trailing_zero_are_canonical(self) -> None:
        negative_zero = normalize_property_binding(self.binding("-0"), self.registry)
        self.assertEqual("0.00", negative_zero["value"]["value"])
        self.assertEqual(
            normalize_property_binding(self.binding("2.500"), self.registry),
            normalize_property_binding(self.binding("2.5"), self.registry),
        )
        long_decimal = "123.123456789012345678901234567890123456789"
        self.assertEqual(
            "123.12",
            normalize_property_binding(self.binding(long_decimal), self.registry)["value"]["value"],
        )

    def test_t41_p2a_unit_004_registry_rejects_bad_unit_graphs(self) -> None:
        core = load_registry(version="2.1")
        mutations = []
        bad_target = deepcopy(core)
        bad_target["units"][2]["canonical_unit_id"] = "unit.percent"
        mutations.append(bad_target)
        bad_ratio = deepcopy(core)
        bad_ratio["units"][2]["to_canonical_denominator"] = 0
        mutations.append(bad_ratio)
        bad_canonical = deepcopy(core)
        bad_canonical["units"][1]["to_canonical_numerator"] = 2
        mutations.append(bad_canonical)
        chain = deepcopy(core)
        chain["units"][1]["canonical_unit_id"] = "unit.mm"
        mutations.append(chain)
        for registry in mutations:
            with self.subTest(registry=registry["units"]), self.assertRaises(RegistryContractError):
                validate_registry_document(registry)

    def test_t41_p2a_unit_005_cross_dimension_unknown_and_float_fail_closed(self) -> None:
        with self.assertRaises(ValueNormalizationError):
            normalize_property_binding(self.binding("10", "unit.percent"), self.registry)
        with self.assertRaises(ValueNormalizationError):
            normalize_property_binding(self.binding("10", "unit.unknown"), self.registry)
        float_binding = self.binding("10")
        float_binding["value"]["value"] = 10.0
        with self.assertRaises(ValueNormalizationError):
            normalize_property_binding(float_binding, self.registry)

    def test_t41_p2a_unit_006_integer_conversion_must_stay_exact(self) -> None:
        registry = deepcopy(self.registry)
        entry = next(
            item for item in registry["properties"] if item["property_id"] == "test.table-column-count"
        )
        entry.update(
            {
                "canonical_unit_id": "unit.pt",
                "allowed_unit_ids": ["unit.mm", "unit.pt"],
                "comparison_precision": {"kind": "decimal_places", "value": 0},
                "normalizer_id": "normalizer.decimal",
            }
        )
        validate_registry_document(registry)
        exact = {
            "property_id": entry["property_id"],
            "value": {"type": "integer", "value": 127},
            "unit_id": "unit.mm",
            "mode": "report",
        }
        self.assertEqual(360, normalize_property_binding(exact, registry)["value"]["value"])
        inexact = deepcopy(exact)
        inexact["value"]["value"] = 1
        with self.assertRaisesRegex(ValueNormalizationError, "exact integer"):
            normalize_property_binding(inexact, registry)

    def test_t41_p2a_unit_007_precision_and_builtin_catalogs_fail_closed(self) -> None:
        core = load_registry(version="2.1")
        missing_precision = deepcopy(self.registry)
        missing_precision["properties"][0]["comparison_precision"] = None
        with self.assertRaisesRegex(RegistryContractError, "comparison_precision"):
            validate_registry_document(missing_precision)

        missing_builtin = deepcopy(core)
        missing_builtin["normalizers"] = missing_builtin["normalizers"][:-1]
        with self.assertRaisesRegex(RegistryContractError, "built-in"):
            validate_registry_document(missing_builtin)
        extra_builtin = deepcopy(core)
        extra_builtin["comparators"].append(
            {
                "comparator_id": "comparator.synthetic",
                "description": "No implementation exists for this synthetic ID.",
            }
        )
        with self.assertRaisesRegex(RegistryContractError, "built-in"):
            validate_registry_document(extra_builtin)


class ScopeContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.golden = load_json(FIXTURES / "canonical-golden-vectors.json")

    def test_t41_p2a_scope_001_scope_id_and_set_order_are_deterministic(self) -> None:
        scope = document_scope("document:test")
        normalized = normalize_scope(scope)
        self.assertEqual(self.golden["document_scope_id"], normalized["scope_id"])
        reordered = {
            "mutually_exclusive_conditions": [],
            "exclusions": [],
            "selectors": [{"selector_ids": ["document:test"], "selector_kind": "document"}],
        }
        self.assertTrue(scope_equal(normalized, normalize_scope(reordered)))
        with self.assertRaisesRegex(ScopeContractError, "scope_id"):
            normalize_scope({**scope, "scope_id": "scope:" + ("0" * 64)})

    def test_t41_p2a_scope_002_positive_objects_are_not_exceptions(self) -> None:
        scope = {
            "selectors": [{"selector_kind": "object", "selector_ids": ["object:one"]}],
            "exclusions": [{"selector_kind": "object", "selector_ids": ["object:two"]}],
            "mutually_exclusive_conditions": [],
        }
        normalized = normalize_scope(scope)
        self.assertEqual("object:one", normalized["selectors"][0]["selector_ids"][0])
        self.assertEqual("object:two", normalized["exclusions"][0]["selector_ids"][0])

    def test_t41_p2a_scope_002a_every_selector_kind_is_lossless(self) -> None:
        kinds = [
            "document",
            "section",
            "chapter",
            "semantic_role",
            "object",
            "property",
            "rule",
            "conflict",
        ]
        scope = {
            "selectors": [
                {"selector_kind": kind, "selector_ids": [f"{kind}:one"]}
                for kind in reversed(kinds)
            ],
            "exclusions": [
                {"selector_kind": "object", "selector_ids": ["object:excluded"]},
                {"selector_kind": "chapter", "selector_ids": ["chapter:excluded"]},
            ],
            "mutually_exclusive_conditions": [
                {
                    "condition_id": "condition:two",
                    "condition_kind": "excludes_selector",
                    "target": {"selector_kind": "object", "selector_ids": ["object:two"]},
                },
                {
                    "condition_id": "condition:one",
                    "condition_kind": "requires_selector",
                    "target": {"selector_kind": "section", "selector_ids": ["section:one"]},
                },
            ],
        }
        normalized = normalize_scope(scope)
        self.assertEqual(
            sorted(kinds),
            [item["selector_kind"] for item in normalized["selectors"]],
        )
        permuted = deepcopy(scope)
        permuted["selectors"].reverse()
        permuted["exclusions"].reverse()
        permuted["mutually_exclusive_conditions"].reverse()
        self.assertEqual(normalized, normalize_scope(permuted))

        for index, kind in enumerate(kinds):
            changed = deepcopy(scope)
            changed["selectors"][len(kinds) - index - 1]["selector_ids"] = [
                f"{kind}:changed"
            ]
            with self.subTest(kind=kind):
                self.assertNotEqual(
                    normalized["scope_id"], normalize_scope(changed)["scope_id"]
                )

    def test_t41_p2a_scope_003_relations_are_conservative(self) -> None:
        whole = normalize_scope(document_scope("document:test"))
        chapter = normalize_scope(
            {
                "selectors": [
                    {"selector_kind": "document", "selector_ids": ["document:test"]},
                    {"selector_kind": "chapter", "selector_ids": ["chapter:one"]},
                ],
                "exclusions": [],
                "mutually_exclusive_conditions": [],
            }
        )
        other = normalize_scope(document_scope("document:other"))
        conditional = normalize_scope(
            {
                **document_scope("document:test"),
                "mutually_exclusive_conditions": [
                    {
                        "condition_id": "condition:one",
                        "condition_kind": "requires_selector",
                        "target": {"selector_kind": "chapter", "selector_ids": ["chapter:one"]},
                    }
                ],
            }
        )
        self.assertTrue(scope_subset(chapter, whole))
        self.assertFalse(scope_subset(whole, chapter))
        self.assertTrue(scope_disjoint(whole, other))
        self.assertEqual("overlap", scope_overlap_state(whole, chapter))
        self.assertEqual("unknown", scope_overlap_state(whole, conditional))

    def test_t41_p2a_scope_004_module_boundary_is_fail_closed(self) -> None:
        asset = layered_v21(layer="module")
        validate_module_asset_scope(asset)
        out_of_scope = deepcopy(asset)
        out_of_scope["rules"][0]["semantic_object_kind"] = "table"
        with self.assertRaisesRegex(ScopeContractError, "semantic object"):
            validate_module_asset_scope(out_of_scope)
        unknown = deepcopy(asset)
        unknown["rules"][0]["scope"]["mutually_exclusive_conditions"] = [
            {
                "condition_id": "condition:unknown",
                "condition_kind": "requires_selector",
                "target": {"selector_kind": "chapter", "selector_ids": ["chapter:one"]},
            }
        ]
        with self.assertRaisesRegex(ScopeContractError, "not provably"):
            validate_module_asset_scope(unknown)

    def test_t41_p2a_scope_005_scope_change_changes_id_and_artifact_fingerprint(self) -> None:
        first = layered_v21()
        second = deepcopy(first)
        second["rules"][0]["scope"] = document_scope("document:other")
        first = stamp_test(first)
        second = stamp_test(second)
        self.assertNotEqual(first["semantic_fingerprint"], second["semantic_fingerprint"])


class NormalizedCompositionKeyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.registry = test_registry()
        self.scope = normalize_scope(
            {
                "selectors": [
                    {
                        "selector_kind": "document",
                        "selector_ids": ["document:alpha", "document:beta"],
                    },
                    {
                        "selector_kind": "section",
                        "selector_ids": ["section:alpha", "section:beta"],
                    },
                ],
                "exclusions": [],
                "mutually_exclusive_conditions": [],
            }
        )

    def _profile_with_duplicate(self, scope: dict) -> dict:
        profile = final_profile_v21()
        profile["resolved_properties"][0]["key"]["normalized_scope"] = deepcopy(
            self.scope
        )
        duplicate = deepcopy(profile["resolved_properties"][0])
        duplicate["resolution_id"] = "resolution:p2a-font-size-duplicate"
        duplicate["key"]["normalized_scope"] = deepcopy(scope)
        profile["resolved_properties"].append(duplicate)
        return profile

    def _conflicts_with_duplicate(self, scope: dict) -> dict:
        report = conflict_v21()
        report["conflicts"][0]["key"]["normalized_scope"] = deepcopy(self.scope)
        duplicate = deepcopy(report["conflicts"][0])
        duplicate["conflict_id"] = "conflict:p2a-font-size-duplicate"
        duplicate["key"]["normalized_scope"] = deepcopy(scope)
        report["conflicts"].append(duplicate)
        return report

    def _assert_duplicate_rejected(self, document: dict) -> None:
        document = stamp_test(document, self.registry)
        with self.assertRaisesRegex(ArtifactContractError, "composition key"):
            _validate_artifact_for_test(
                document,
                registry=self.registry,
                features={"profile_v2_schema": True},
            )

    def test_t41_p2a_scope_008_selector_order_cannot_split_composition_key(self) -> None:
        reordered = deepcopy(self.scope)
        reordered["selectors"].reverse()
        self._assert_duplicate_rejected(self._profile_with_duplicate(reordered))
        self._assert_duplicate_rejected(self._conflicts_with_duplicate(reordered))

    def test_t41_p2a_scope_009_selector_id_order_cannot_split_composition_key(self) -> None:
        reordered = deepcopy(self.scope)
        for selector in reordered["selectors"]:
            selector["selector_ids"].reverse()
        self._assert_duplicate_rejected(self._profile_with_duplicate(reordered))
        self._assert_duplicate_rejected(self._conflicts_with_duplicate(reordered))

    def test_t41_p2a_scope_010_nfc_equivalence_cannot_split_composition_key(self) -> None:
        composed = normalize_scope(document_scope("document:Caf\u00e9"))
        decomposed = {
            **deepcopy(composed),
            "selectors": [
                {
                    "selector_kind": "document",
                    "selector_ids": ["document:Cafe\u0301"],
                }
            ],
        }
        self.assertEqual(
            normalized_property_scope_key(
                "paragraph", "test.paragraph-font-size", composed
            ),
            normalized_property_scope_key(
                "paragraph", "test.paragraph-font-size", decomposed
            ),
        )
        for kind, document in (
            ("final-execution-profile", self._profile_with_duplicate(decomposed)),
            ("conflict-report", self._conflicts_with_duplicate(decomposed)),
        ):
            if kind == "final-execution-profile":
                document["resolved_properties"][0]["key"]["normalized_scope"] = composed
            else:
                document["conflicts"][0]["key"]["normalized_scope"] = composed
            with self.subTest(kind=kind):
                errors = artifact_semantic_errors(kind, document, self.registry)
                self.assertTrue(any("composition key" in error for error in errors))

    def test_t41_p2a_scope_011_different_real_scopes_can_coexist(self) -> None:
        other = normalize_scope(document_scope("document:other"))
        profile = stamp_test(self._profile_with_duplicate(other), self.registry)
        _validate_artifact_for_test(
            profile,
            registry=self.registry,
            features={"profile_v2_schema": True},
        )
        report = stamp_test(self._conflicts_with_duplicate(other), self.registry)
        _validate_artifact_for_test(
            report,
            registry=self.registry,
            features={"profile_v2_schema": True},
        )


class V21ArtifactIntegrationTests(unittest.TestCase):
    def test_t41_p2a_art_001_private_v21_contract_remains_non_runtime(self) -> None:
        registry = test_registry()
        artifact = stamp_test(layered_v21(), registry)
        result = _validate_artifact_for_test(
            artifact,
            registry=registry,
            features={"profile_v2_schema": True},
        )
        self.assertFalse(result.runtime_eligible)
        self.assertEqual("disabled", result.activation)

    def test_t41_p2a_art_002_stale_v21_fingerprint_is_rejected(self) -> None:
        registry = test_registry()
        artifact = stamp_test(layered_v21(), registry)
        artifact["rules"][0]["confidence"] = "low"
        with self.assertRaisesRegex(ArtifactContractError, "semantic_fingerprint"):
            _validate_artifact_for_test(
                artifact,
                registry=registry,
                features={"profile_v2_schema": True},
            )

    def test_t41_p2a_art_003_test_catalog_drift_is_rejected(self) -> None:
        registry = test_registry()
        registry["normalizers"][0]["description"] = "Drifted private catalog."
        artifact = layered_v21()
        with self.assertRaisesRegex(ArtifactContractError, "reuse the production"):
            _validate_artifact_for_test(
                artifact,
                registry=registry,
                features={"profile_v2_schema": True},
            )

    def test_t41_p2a_art_004_conflict_and_final_use_full_normalized_scope(self) -> None:
        registry = test_registry()
        for document in (conflict_v21(), final_profile_v21()):
            stamped = stamp_test(document, registry)
            with self.subTest(kind=document["artifact_kind"]):
                result = _validate_artifact_for_test(
                    stamped,
                    registry=registry,
                    features={"profile_v2_schema": True},
                )
                self.assertFalse(result.runtime_eligible)

                stale_scope = deepcopy(stamped)
                if document["artifact_kind"] == "conflict-report":
                    scope = stale_scope["conflicts"][0]["key"]["normalized_scope"]
                else:
                    scope = stale_scope["resolved_properties"][0]["key"]["normalized_scope"]
                scope["scope_id"] = "scope:" + ("0" * 64)
                with self.assertRaisesRegex(ArtifactContractError, "scope_id"):
                    _validate_artifact_for_test(
                        stale_scope,
                        registry=registry,
                        features={"profile_v2_schema": True},
                    )


if __name__ == "__main__":
    unittest.main()
