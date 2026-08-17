#!/usr/bin/env python3
"""Offline V0.4.1 P1 artifact schema dispatch and read-only version routing."""

from __future__ import annotations

import json
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from jsonschema import Draft202012Validator, FormatChecker
from referencing import Registry, Resource

from profile_v2_registry import (
    RegistryContractError,
    build_property_catalog_schema,
    build_typed_value_schema,
    load_registry,
    property_index,
    validate_binding_for_layer,
    validate_registry_document,
    verify_committed_catalog,
)


SCHEMA_DIR = Path(__file__).resolve().parent.parent / "references" / "schemas" / "v2"
LEGACY_SCHEMA_PATH = (
    Path(__file__).resolve().parent.parent / "references" / "format-profile.schema.json"
)

ARTIFACT_KINDS = (
    "layered-rule-asset",
    "final-execution-profile",
    "qa-approval-artifact",
    "legacy-migration-manifest",
    "feature-activation-manifest",
    "capability-snapshot",
    "conflict-report",
    "execution-evidence-artifact",
)

ARTIFACT_SCHEMA_FILES = {
    kind: f"{kind}.schema.json" for kind in ARTIFACT_KINDS
}
VERSION_PATTERN = re.compile(r"^(?P<major>[0-9]+)\.(?P<minor>[0-9]+)$")


class ArtifactContractError(ValueError):
    """Raised when an artifact or version violates the P1 contract."""


class ProfileV2DisabledError(ArtifactContractError):
    """Raised when the default-off V2 schema path is not explicitly enabled."""


@dataclass(frozen=True)
class ProfileReadResult:
    schema_version: str
    artifact_kind: str
    legacy_input: bool
    activation: str
    read_only: bool
    runtime_eligible: bool
    document: dict[str, Any]


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ArtifactContractError(f"Schema or artifact file does not exist: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ArtifactContractError(f"Invalid JSON in {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ArtifactContractError(f"Expected a JSON object in {path}.")
    return value


def profile_v2_schema_enabled(features: Mapping[str, Any] | None) -> bool:
    return bool(features and features.get("profile_v2_schema") is True)


def load_artifact_schema(artifact_kind: str) -> dict[str, Any]:
    try:
        filename = ARTIFACT_SCHEMA_FILES[artifact_kind]
    except KeyError as exc:
        raise ArtifactContractError(f"Unknown V2 artifact_kind: {artifact_kind}") from exc
    return _load_json(SCHEMA_DIR / filename)


def _schema_documents(
    schema_overrides: Mapping[str, dict[str, Any]] | None = None,
) -> dict[str, dict[str, Any]]:
    documents: dict[str, dict[str, Any]] = {}
    for path in sorted(SCHEMA_DIR.glob("*.schema.json")):
        schema = _load_json(path)
        Draft202012Validator.check_schema(schema)
        schema_id = schema.get("$id")
        if not isinstance(schema_id, str):
            raise ArtifactContractError(f"Schema is missing a stable $id: {path.name}")
        if schema_id in documents:
            raise ArtifactContractError(f"Duplicate schema $id: {schema_id}")
        documents[schema_id] = schema
    for schema_id, schema in (schema_overrides or {}).items():
        if schema_id not in documents:
            raise ArtifactContractError(f"Schema override has no registered target: {schema_id}")
        if schema.get("$id") != schema_id:
            raise ArtifactContractError(f"Schema override ID does not match its target: {schema_id}")
        Draft202012Validator.check_schema(schema)
        documents[schema_id] = deepcopy(schema)
    return documents


def schema_documents() -> dict[str, dict[str, Any]]:
    """Return only the repository's committed offline schema set."""

    return _schema_documents()


def _offline_schema_registry(
    schema_overrides: Mapping[str, dict[str, Any]] | None = None,
) -> Registry:
    registry = Registry()
    for schema_id, schema in _schema_documents(schema_overrides).items():
        registry = registry.with_resource(schema_id, Resource.from_contents(schema))
    return registry


def offline_schema_registry() -> Registry:
    """Build the production resolver from committed local schemas only."""

    return _offline_schema_registry()


def _format_errors(validator: Draft202012Validator, value: Any) -> list[str]:
    return [
        f"{'/'.join(str(item) for item in error.absolute_path) or '<root>'}: {error.message}"
        for error in sorted(validator.iter_errors(value), key=lambda item: list(item.path))
    ]


def _schema_errors(
    artifact_kind: str,
    document: dict[str, Any],
    *,
    schema_override: dict[str, Any] | None = None,
    schema_documents_override: Mapping[str, dict[str, Any]] | None = None,
) -> list[str]:
    schema = schema_override or load_artifact_schema(artifact_kind)
    Draft202012Validator.check_schema(schema)
    validator = Draft202012Validator(
        schema,
        registry=_offline_schema_registry(schema_documents_override),
        format_checker=FormatChecker(),
    )
    return _format_errors(validator, document)


def schema_errors(artifact_kind: str, document: dict[str, Any]) -> list[str]:
    """Validate shape against committed schemas without accepting overrides."""

    return _schema_errors(artifact_kind, document)


def _require_test_registry(registry: dict[str, Any]) -> None:
    validate_registry_document(registry)
    if registry.get("registry_scope") != "test":
        raise ArtifactContractError(
            "The internal test contract requires a validated test registry."
        )


def _test_schema_overrides(registry: dict[str, Any]) -> dict[str, dict[str, Any]]:
    typed_values = build_typed_value_schema(registry)
    property_catalog = build_property_catalog_schema(registry)
    return {
        typed_values["$id"]: typed_values,
        property_catalog["$id"]: property_catalog,
    }


def _schema_errors_for_test(
    artifact_kind: str,
    document: dict[str, Any],
    *,
    registry: dict[str, Any],
    schema_override: dict[str, Any] | None = None,
) -> list[str]:
    """Test-only schema path for registry-derived synthetic contracts."""

    _require_test_registry(registry)
    return _schema_errors(
        artifact_kind,
        document,
        schema_override=schema_override,
        schema_documents_override=_test_schema_overrides(registry),
    )


def _parse_version(version: Any) -> tuple[int, int]:
    if not isinstance(version, str):
        raise ArtifactContractError("schema_version must be a major.minor string.")
    match = VERSION_PATTERN.fullmatch(version)
    if not match:
        raise ArtifactContractError(f"Invalid schema_version: {version}")
    return int(match.group("major")), int(match.group("minor"))


def schema_for_requested_minor(
    schema: dict[str, Any], requested_version: str
) -> tuple[dict[str, Any], bool]:
    """Return a read schema and whether it is a compatible-minor read."""

    requested_major, _ = _parse_version(requested_version)
    current_version = schema.get("properties", {}).get("schema_version", {}).get("const")
    current_major, _ = _parse_version(current_version)
    if requested_major != current_major:
        raise ArtifactContractError(
            f"Unknown major schema version {requested_version}; expected {current_major}.x."
        )
    if requested_version == current_version:
        return schema, False
    compatible = schema.get("x-read-compatible-minor-versions", [])
    if requested_version not in compatible:
        raise ArtifactContractError(
            f"Schema minor version {requested_version} is not declared read-compatible."
        )
    result = deepcopy(schema)
    result["properties"]["schema_version"]["const"] = requested_version
    return result, True


def artifact_semantic_errors(
    artifact_kind: str,
    document: dict[str, Any],
    registry: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    properties = property_index(registry)

    seen_input_ids: dict[str, tuple[str, str]] = {}
    for item in document.get("input_fingerprints", []):
        input_id = item["input_id"]
        value = (item["role"], item["fingerprint"])
        if input_id in seen_input_ids:
            errors.append(
                f"Duplicate input_id {input_id} cannot represent multiple fingerprint bindings."
            )
        else:
            seen_input_ids[input_id] = value

    def validate_key_and_candidate(
        key: dict[str, Any], candidate: dict[str, Any], context: str
    ) -> None:
        property_id = key["property_id"]
        binding = candidate["property_binding"]
        if binding["property_id"] != property_id:
            errors.append(f"{context} candidate property does not match its normalized key.")
        entry = properties.get(property_id)
        if entry is None:
            errors.append(f"{context} references unregistered property {property_id}.")
        elif entry.get("safety_invariant"):
            errors.append(
                f"{context} cannot place safety invariant {property_id} in a rule candidate chain."
            )
        elif key["semantic_object_kind"] not in entry["semantic_object_kinds"]:
            errors.append(
                f"{context} uses property {property_id} for an unsupported semantic object."
            )
        try:
            validate_binding_for_layer(binding, candidate["layer_kind"], registry)
        except RegistryContractError as exc:
            errors.append(str(exc))

    if artifact_kind == "layered-rule-asset":
        layer_kind = document["layer_kind"]
        seen_rules: set[str] = set()
        for rule in document.get("rules", []):
            rule_id = rule["rule_id"]
            if rule_id in seen_rules:
                errors.append(f"Duplicate rule ID: {rule_id}.")
            seen_rules.add(rule_id)
            seen_properties: set[str] = set()
            for binding in rule.get("properties", []):
                property_id = binding["property_id"]
                if property_id in seen_properties:
                    errors.append(f"Rule {rule_id} repeats property {property_id}.")
                seen_properties.add(property_id)
                try:
                    validate_binding_for_layer(binding, layer_kind, registry)
                except RegistryContractError as exc:
                    errors.append(str(exc))
    elif artifact_kind == "final-execution-profile":
        if document.get("legacy_input") is not False:
            errors.append("V2 final execution profiles cannot be legacy inputs.")
        if document.get("activation") != "disabled":
            errors.append("P1 final execution profiles must remain disabled.")

        inputs_by_role: dict[str, list[str]] = {}
        for item in document.get("input_fingerprints", []):
            inputs_by_role.setdefault(item["role"], []).append(item["fingerprint"])
        bindings = document["bindings"]
        singleton_bindings = {
            "source_document": "input_fingerprint",
            "feature_activation": "feature_activation_fingerprint",
            "property_registry": "property_registry_fingerprint",
            "structure": "structure_fingerprint",
            "conflict_report": "conflict_report_fingerprint",
        }
        for role, field in singleton_bindings.items():
            values = inputs_by_role.get(role, [])
            if len(values) != 1:
                errors.append(f"Final profile requires exactly one {role} input binding.")
            elif values[0] != bindings[field]:
                errors.append(f"Final profile {role} fingerprint does not match bindings.{field}.")
        for role, field in (
            ("profile", "profile_fingerprints"),
            ("approval", "approval_fingerprints"),
        ):
            values = inputs_by_role.get(role, [])
            if not values:
                errors.append(f"Final profile requires at least one {role} input binding.")
            elif len(values) != len(set(values)):
                errors.append(f"Final profile has duplicate {role} input fingerprints.")
            elif sorted(values) != sorted(bindings[field]):
                errors.append(f"Final profile {role} fingerprints do not match bindings.{field}.")

        seen_resolution_ids: set[str] = set()
        seen_composition_keys: set[str] = set()
        for resolved in document.get("resolved_properties", []):
            resolution_id = resolved["resolution_id"]
            if resolution_id in seen_resolution_ids:
                errors.append(f"Duplicate resolution_id: {resolution_id}.")
            seen_resolution_ids.add(resolution_id)
            key = resolved["key"]
            composition_key = json.dumps(key, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
            if composition_key in seen_composition_keys:
                errors.append("Final profile repeats a semantic object/property/scope composition key.")
            seen_composition_keys.add(composition_key)
            binding = resolved["resolved_binding"]
            if binding["property_id"] != key["property_id"]:
                errors.append("Resolved property binding does not match its normalized key.")
            if resolved["execution_mode"] != binding["mode"]:
                errors.append("Resolved property execution_mode does not match its binding mode.")
            entry = properties.get(key["property_id"])
            if entry is not None and entry.get("safety_invariant"):
                errors.append(
                    f"Safety invariant {key['property_id']} cannot be a normal resolved property."
                )
            synthetic_candidate = {
                "property_binding": binding,
                "layer_kind": resolved["final_layer_kind"],
            }
            validate_key_and_candidate(key, synthetic_candidate, "Resolved property")
            applicable_ids: set[str] = set()
            active_ids: set[str] = set()
            excluded_ids: set[str] = set()
            final_source_matches = False
            for candidate in resolved["candidate_chain"]:
                candidate_id = candidate["candidate_id"]
                if candidate_id in active_ids:
                    errors.append(f"Resolved property repeats active candidate_id {candidate_id}.")
                active_ids.add(candidate_id)
                validate_key_and_candidate(key, candidate, "Resolved property")
                if candidate["scope_status"] == "applicable":
                    applicable_ids.add(candidate_id)
                else:
                    errors.append(
                        "Active resolved property candidates must be applicable; "
                        "exclusions are separate."
                    )
                if (
                    candidate["source"] == resolved["final_source"]
                    and candidate["layer_kind"] == resolved["final_layer_kind"]
                    and candidate["property_binding"] == binding
                ):
                    final_source_matches = True
            for excluded in resolved["excluded_candidates"]:
                candidate = excluded["candidate"]
                candidate_id = candidate["candidate_id"]
                if candidate_id in excluded_ids:
                    errors.append(f"Resolved property repeats excluded candidate_id {candidate_id}.")
                excluded_ids.add(candidate_id)
                validate_key_and_candidate(key, candidate, "Excluded resolved property")
                if candidate["scope_status"] == "applicable":
                    errors.append("Excluded resolved candidates cannot be marked applicable.")
            overlapping_ids = active_ids & excluded_ids
            if overlapping_ids:
                errors.append("Resolved property candidate IDs cannot be both active and excluded.")
            if not final_source_matches:
                errors.append("Resolved property final source is absent from its candidate chain.")
            unknown_override_ids = set(resolved["override_chain"]) - applicable_ids
            if unknown_override_ids:
                errors.append("Resolved property override chain references non-applicable candidates.")
            for invariant_id in resolved["safety_check"]["checked_invariant_ids"]:
                invariant = properties.get(invariant_id)
                if invariant is None or not invariant.get("safety_invariant"):
                    errors.append(
                        f"Safety check references non-safety property {invariant_id}."
                    )
    elif artifact_kind == "conflict-report":
        for conflict in document.get("conflicts", []):
            key = conflict["key"]
            layers: set[str] = set()
            for candidate in conflict["candidates"]:
                validate_key_and_candidate(key, candidate, "Conflict")
                layers.add(candidate["layer_kind"])
                if candidate["scope_status"] != "applicable":
                    errors.append("Conflict candidates must be applicable; exclusions are separate.")
            for excluded in conflict["excluded_candidates"]:
                candidate = excluded["candidate"]
                validate_key_and_candidate(key, candidate, "Excluded conflict")
                if candidate["scope_status"] == "applicable":
                    errors.append("Excluded conflict candidates cannot be marked applicable.")
            if conflict["reason"] == "same_layer":
                if len(conflict["candidates"]) < 2 or len(layers) != 1:
                    errors.append("same_layer conflicts require at least two candidates in one layer.")
            if conflict["reason"] == "scope_violation" and not any(
                item["exclusion_reason"] == "scope_violation"
                for item in conflict["excluded_candidates"]
            ):
                errors.append("scope_violation conflicts require a scope-violation exclusion.")
    return errors


def _validate_artifact_contract(
    document: dict[str, Any],
    *,
    features: Mapping[str, Any] | None,
    registry: dict[str, Any],
    schema_override: dict[str, Any] | None = None,
    schema_documents_override: Mapping[str, dict[str, Any]] | None = None,
) -> ProfileReadResult:
    if not profile_v2_schema_enabled(features):
        raise ProfileV2DisabledError(
            "profile_v2_schema is disabled by default; enable it only for the P1 schema path."
        )
    artifact_kind = document.get("artifact_kind")
    if artifact_kind not in ARTIFACT_KINDS:
        raise ArtifactContractError(f"Unknown V2 artifact_kind: {artifact_kind}")
    schema = schema_override or load_artifact_schema(artifact_kind)
    effective_schema, compatible_minor = schema_for_requested_minor(
        schema, document.get("schema_version")
    )
    errors = _schema_errors(
        artifact_kind,
        document,
        schema_override=effective_schema,
        schema_documents_override=schema_documents_override,
    )
    if not errors:
        errors.extend(artifact_semantic_errors(artifact_kind, document, registry))
    if errors:
        raise ArtifactContractError("Invalid V2 artifact: " + " | ".join(errors))
    return ProfileReadResult(
        schema_version=document["schema_version"],
        artifact_kind=artifact_kind,
        legacy_input=False,
        activation="disabled",
        read_only=compatible_minor,
        runtime_eligible=False,
        document=deepcopy(document),
    )


def validate_artifact(
    document: dict[str, Any],
    *,
    features: Mapping[str, Any] | None = None,
) -> ProfileReadResult:
    """Validate against the committed production registry and schema contracts."""

    effective_registry = load_registry()
    validate_registry_document(effective_registry)
    verify_committed_catalog(effective_registry)
    return _validate_artifact_contract(
        document,
        features=features,
        registry=effective_registry,
    )


def _validate_artifact_for_test(
    document: dict[str, Any],
    *,
    registry: dict[str, Any],
    features: Mapping[str, Any] | None = None,
    schema_override: dict[str, Any] | None = None,
) -> ProfileReadResult:
    """Validate synthetic contracts without making them production-eligible."""

    _require_test_registry(registry)
    return _validate_artifact_contract(
        document,
        features=features,
        registry=registry,
        schema_override=schema_override,
        schema_documents_override=_test_schema_overrides(registry),
    )


def _validate_legacy_profile(document: dict[str, Any]) -> None:
    from validate_profile import semantic_errors

    schema = _load_json(LEGACY_SCHEMA_PATH)
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = _format_errors(validator, document)
    if not errors:
        errors.extend(semantic_errors(document))
    if errors:
        raise ArtifactContractError("Invalid legacy profile: " + " | ".join(errors))


def read_profile_document(
    document: dict[str, Any],
    *,
    features: Mapping[str, Any] | None = None,
) -> ProfileReadResult:
    """Dispatch 1.0/1.1 legacy reads and gated V2 artifact reads explicitly."""

    major, minor = _parse_version(document.get("schema_version"))
    if major == 1:
        if minor not in {0, 1}:
            raise ArtifactContractError(
                f"Legacy schema minor version 1.{minor} is not supported."
            )
        _validate_legacy_profile(document)
        return ProfileReadResult(
            schema_version=document["schema_version"],
            artifact_kind="legacy-format-profile",
            legacy_input=True,
            activation="disabled",
            read_only=True,
            runtime_eligible=False,
            document=deepcopy(document),
        )
    if major == 2:
        return validate_artifact(
            document,
            features=features,
        )
    raise ArtifactContractError(f"Unknown major schema version: {major}.")


def require_v2_nonlegacy_contract(result: ProfileReadResult) -> dict[str, Any]:
    if result.legacy_input or result.artifact_kind == "legacy-format-profile":
        raise ArtifactContractError(
            "Legacy profiles are read-only migration/report inputs and cannot enter V2 contracts."
        )
    return deepcopy(result.document)
