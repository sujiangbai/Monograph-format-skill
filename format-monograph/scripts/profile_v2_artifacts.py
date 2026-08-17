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
    load_registry,
    validate_binding_for_layer,
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


def schema_documents() -> dict[str, dict[str, Any]]:
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
    return documents


def offline_schema_registry() -> Registry:
    registry = Registry()
    for schema_id, schema in schema_documents().items():
        registry = registry.with_resource(schema_id, Resource.from_contents(schema))
    return registry


def _format_errors(validator: Draft202012Validator, value: Any) -> list[str]:
    return [
        f"{'/'.join(str(item) for item in error.absolute_path) or '<root>'}: {error.message}"
        for error in sorted(validator.iter_errors(value), key=lambda item: list(item.path))
    ]


def schema_errors(
    artifact_kind: str,
    document: dict[str, Any],
    *,
    schema_override: dict[str, Any] | None = None,
) -> list[str]:
    schema = schema_override or load_artifact_schema(artifact_kind)
    Draft202012Validator.check_schema(schema)
    validator = Draft202012Validator(
        schema,
        registry=offline_schema_registry(),
        format_checker=FormatChecker(),
    )
    return _format_errors(validator, document)


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
    return errors


def validate_artifact(
    document: dict[str, Any],
    *,
    features: Mapping[str, Any] | None = None,
    registry: dict[str, Any] | None = None,
    schema_override: dict[str, Any] | None = None,
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
    errors = schema_errors(artifact_kind, document, schema_override=effective_schema)
    registry = registry or load_registry()
    verify_committed_catalog() if registry.get("registry_scope") == "production" else None
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
    schema_override: dict[str, Any] | None = None,
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
            schema_override=schema_override,
        )
    raise ArtifactContractError(f"Unknown major schema version: {major}.")


def require_v2_nonlegacy_contract(result: ProfileReadResult) -> dict[str, Any]:
    if result.legacy_input or result.artifact_kind == "legacy-format-profile":
        raise ArtifactContractError(
            "Legacy profiles are read-only migration/report inputs and cannot enter V2 contracts."
        )
    return deepcopy(result.document)
