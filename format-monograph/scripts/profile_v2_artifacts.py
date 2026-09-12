#!/usr/bin/env python3
"""Offline V0.4.1 P1 artifact schema dispatch and read-only version routing."""

from __future__ import annotations

import hashlib
import json
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping

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
from profile_v2_authority import (
    AuthorityContractError,
    authority_contract_fingerprint,
    load_authority_contract,
    verify_authority_projection,
    verify_legacy_layer_compatibility,
)


SCHEMA_DIR = Path(__file__).resolve().parent.parent / "references" / "schemas" / "v2"
CONTRACT_MATRIX_PATH = SCHEMA_DIR / "artifact-contract-matrix.v1.0.json"
CONTRACT_MATRIX_SCHEMA_PATH = SCHEMA_DIR / "artifact-contract-matrix.schema.json"
CONTRACT_MATRIX_PATHS = {
    "1.0": CONTRACT_MATRIX_PATH,
    "1.1": SCHEMA_DIR / "artifact-contract-matrix.v1.1.json",
}
CONTRACT_MATRIX_SCHEMA_PATHS = {
    "1.0": CONTRACT_MATRIX_SCHEMA_PATH,
    "1.1": SCHEMA_DIR / "artifact-contract-matrix.v1.1.schema.json",
}
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

_MATRIX_SOURCE = json.loads(CONTRACT_MATRIX_PATH.read_text(encoding="utf-8"))
ARTIFACT_SCHEMA_FILES = {
    (
        route["artifact_kind"],
        route["schema_version"],
        route["registry_contract_version"],
        route["authority_contract_version"],
    ): route["schema_file"]
    for route in _MATRIX_SOURCE["routes"]
}
VERSION_PATTERN = re.compile(r"^(?P<major>[0-9]+)\.(?P<minor>[0-9]+)$")
TEST_REGISTRY_SHARED_CATALOGS = (
    "data_types",
    "units",
    "normalizers",
    "comparators",
    "executor_capabilities",
    "auditor_capabilities",
    "constraints",
)


class ArtifactContractError(ValueError):
    """Raised when an artifact or version violates the P1 contract."""


class ProfileV2DisabledError(ArtifactContractError):
    """Raised when the default-off V2 schema path is not explicitly enabled."""


class ArtifactRouteError(ArtifactContractError):
    """Raised before registry loading when a four-part contract route is invalid."""


class ArtifactDagError(ArtifactContractError):
    """Raised when artifact bindings do not form the approved acyclic graph."""


@dataclass(frozen=True)
class ArtifactEnvelope:
    artifact_kind: str
    schema_version: str
    registry_contract_version: str | None
    authority_contract_version: str | None


@dataclass(frozen=True)
class ContractRoute:
    route_id: str
    artifact_kind: str
    schema_version: str
    registry_contract_version: str
    authority_contract_version: str
    schema_file: str
    schema_id: str
    version_source: str
    registry_validation_context: str = "strict_execution"


@dataclass(frozen=True)
class SchemaResourceContract:
    schema_file: str
    schema_id: str
    schema_version: str
    registry_contract_version: str
    resource_kind: str
    fingerprint_inventory: bool
    version_source: str


@dataclass(frozen=True)
class DagValidationResult:
    topological_order: tuple[str, ...]
    artifact_count: int
    runtime_eligible: bool = False


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


def load_artifact_contract_matrix(version: str = "1.0") -> dict[str, Any]:
    """Load the sole editable artifact/schema/registry/authority route table."""

    try:
        schema_path = CONTRACT_MATRIX_SCHEMA_PATHS[version]
        matrix_path = CONTRACT_MATRIX_PATHS[version]
    except KeyError as exc:
        raise ArtifactRouteError(f"Unsupported artifact contract matrix version: {version}") from exc
    schema = _load_json(schema_path)
    Draft202012Validator.check_schema(schema)
    matrix = _load_json(matrix_path)
    errors = _format_errors(Draft202012Validator(schema), matrix)
    if errors:
        raise ArtifactRouteError("Invalid artifact contract matrix: " + " | ".join(errors))
    return matrix


def _contract_routes_from_matrix(matrix: Mapping[str, Any]) -> tuple[ContractRoute, ...]:
    routes = tuple(ContractRoute(**item) for item in matrix["routes"])
    identities = [
        (
            item.artifact_kind,
            item.schema_version,
            item.registry_contract_version,
            item.authority_contract_version,
        )
        for item in routes
    ]
    legacy_pairs = [
        (item.artifact_kind, item.schema_version)
        for item in routes
        if item.version_source == "legacy_matrix"
    ]
    route_ids = [item.route_id for item in routes]
    if len(identities) != len(set(identities)):
        raise ArtifactRouteError("Artifact contract matrix repeats a four-part route identity.")
    if len(legacy_pairs) != len(set(legacy_pairs)):
        raise ArtifactRouteError("Artifact contract matrix repeats a legacy artifact/version pair.")
    if len(route_ids) != len(set(route_ids)):
        raise ArtifactRouteError("Artifact contract matrix repeats a route_id.")
    return routes


def _contract_routes(matrix_version: str = "1.0") -> tuple[ContractRoute, ...]:
    return _contract_routes_from_matrix(load_artifact_contract_matrix(matrix_version))


def _schema_resource_contracts(
    matrix_version: str = "1.0",
) -> tuple[SchemaResourceContract, ...]:
    resources = tuple(
        SchemaResourceContract(**item)
        for item in load_artifact_contract_matrix(matrix_version)["schema_resources"]
    )
    schema_ids = [item.schema_id for item in resources]
    schema_files = [item.schema_file for item in resources]
    if len(schema_ids) != len(set(schema_ids)):
        raise ArtifactRouteError("Artifact contract matrix repeats a schema resource ID.")
    if len(schema_files) != len(set(schema_files)):
        raise ArtifactRouteError("Artifact contract matrix repeats a schema resource file.")
    return resources


def schema_inventory_contract(
    schema_id: str, *, matrix_version: str = "1.0"
) -> tuple[str, str, bool]:
    """Return explicit schema/registry versions and inventory eligibility."""

    route_matches = [
        item for item in _contract_routes(matrix_version) if item.schema_id == schema_id
    ]
    resource_matches = [
        item
        for item in _schema_resource_contracts(matrix_version)
        if item.schema_id == schema_id
    ]
    if len(route_matches) + len(resource_matches) != 1:
        raise ArtifactRouteError(
            f"Schema ID has no unique contract-matrix classification: {schema_id}"
        )
    if route_matches:
        route = route_matches[0]
        return route.schema_version, route.registry_contract_version, True
    resource = resource_matches[0]
    return (
        resource.schema_version,
        resource.registry_contract_version,
        resource.fingerprint_inventory,
    )


def read_minimal_artifact_envelope(document: Mapping[str, Any]) -> ArtifactEnvelope:
    """Read only routing fields; this function never loads a schema or registry."""

    if not isinstance(document, Mapping):
        raise ArtifactRouteError("Artifact envelope must be a JSON object.")
    artifact_kind = document.get("artifact_kind")
    schema_version = document.get("schema_version")
    if artifact_kind not in ARTIFACT_KINDS:
        raise ArtifactRouteError(f"Unknown V2 artifact_kind: {artifact_kind}")
    _parse_version(schema_version)
    registry_version = document.get("registry_contract_version")
    authority_version = document.get("authority_contract_version")
    for field_name, value in (
        ("registry_contract_version", registry_version),
        ("authority_contract_version", authority_version),
    ):
        if value is not None and not isinstance(value, str):
            raise ArtifactRouteError(f"{field_name} must be a string when present.")
    return ArtifactEnvelope(
        artifact_kind=str(artifact_kind),
        schema_version=str(schema_version),
        registry_contract_version=registry_version,
        authority_contract_version=authority_version,
    )


def _route_artifact_contract_from_routes(
    document: Mapping[str, Any], routes: tuple[ContractRoute, ...]
) -> ContractRoute:
    envelope = read_minimal_artifact_envelope(document)
    pair_matches = [
        route
        for route in routes
        if route.artifact_kind == envelope.artifact_kind
        and route.schema_version == envelope.schema_version
    ]
    if not pair_matches:
        raise ArtifactRouteError(
            f"No route for {envelope.artifact_kind}@{envelope.schema_version}."
        )
    sources = {route.version_source for route in pair_matches}
    if sources == {"artifact_fields"}:
        if envelope.registry_contract_version is None or envelope.authority_contract_version is None:
            raise ArtifactRouteError(
                "New artifact contracts require explicit registry_contract_version and "
                "authority_contract_version fields."
            )
        identity = (
            envelope.artifact_kind,
            envelope.schema_version,
            envelope.registry_contract_version,
            envelope.authority_contract_version,
        )
        exact = [
            route
            for route in pair_matches
            if (
                route.artifact_kind,
                route.schema_version,
                route.registry_contract_version,
                route.authority_contract_version,
            )
            == identity
        ]
        if len(exact) != 1:
            raise ArtifactRouteError("Artifact four-part contract does not match the route matrix.")
        return exact[0]
    if sources == {"legacy_matrix"}:
        if envelope.registry_contract_version is not None or envelope.authority_contract_version is not None:
            raise ArtifactRouteError("Legacy artifact contracts cannot self-declare route versions.")
        if len(pair_matches) != 1:
            raise ArtifactRouteError(
                "Legacy artifact contracts require one exact matrix row per artifact/version pair."
            )
        return pair_matches[0]
    raise ArtifactRouteError(
        "Artifact/version pair mixes legacy and explicit route sources."
    )


def route_artifact_contract(
    document: Mapping[str, Any], *, matrix_version: str = "1.0"
) -> ContractRoute:
    """Resolve the complete route before any registry or authority is loaded."""

    return _route_artifact_contract_from_routes(
        document, _contract_routes(matrix_version)
    )


def verify_contract_matrix_alignment(
    route_index: Mapping[tuple[str, str, str, str], str] | None = None,
    *,
    matrix_version: str = "1.0",
) -> None:
    """Reject drift in either direction between the matrix, schemas, and Python index."""

    routes = _contract_routes(matrix_version)
    expected = {
        (
            item.artifact_kind,
            item.schema_version,
            item.registry_contract_version,
            item.authority_contract_version,
        ): item.schema_file
        for item in routes
    }
    if route_index is None:
        if matrix_version == "1.0":
            actual = dict(ARTIFACT_SCHEMA_FILES)
        else:
            actual = {
                (
                    item.artifact_kind,
                    item.schema_version,
                    item.registry_contract_version,
                    item.authority_contract_version,
                ): item.schema_file
                for item in routes
            }
    else:
        actual = dict(route_index)
    if expected != actual:
        raise ArtifactRouteError("Python artifact route index differs from the contract matrix.")
    resources = _schema_resource_contracts(matrix_version)
    routed_schema_ids = {item.schema_id for item in routes}
    resource_schema_ids = {item.schema_id for item in resources}
    if routed_schema_ids & resource_schema_ids:
        raise ArtifactRouteError("A schema cannot be both an artifact route and a resource.")
    expected_files = {item.schema_file for item in routes} | {
        item.schema_file for item in resources
    }
    actual_files = {path.name for path in SCHEMA_DIR.glob("*.schema.json")}
    if matrix_version == "1.1" and expected_files != actual_files:
        raise ArtifactRouteError(
            "Artifact contract matrix does not classify the complete offline schema set."
        )
    if matrix_version == "1.0" and not expected_files.issubset(actual_files):
        raise ArtifactRouteError("The immutable 1.0 schema subset is incomplete.")
    for route in routes:
        path = SCHEMA_DIR / route.schema_file
        schema = _load_json(path)
        if schema.get("$id") != route.schema_id:
            raise ArtifactRouteError(f"Route {route.route_id} has a mismatched schema_id.")
        properties = schema.get("properties", {})
        if properties.get("artifact_kind", {}).get("const") != route.artifact_kind:
            raise ArtifactRouteError(f"Route {route.route_id} has a mismatched artifact_kind.")
        if properties.get("schema_version", {}).get("const") != route.schema_version:
            raise ArtifactRouteError(f"Route {route.route_id} has a mismatched schema_version.")
        if route.version_source == "artifact_fields":
            required = set(schema.get("required", []))
            for field_name, expected_value in (
                ("registry_contract_version", route.registry_contract_version),
                ("authority_contract_version", route.authority_contract_version),
            ):
                if field_name not in required:
                    raise ArtifactRouteError(
                        f"Route {route.route_id} schema does not require {field_name}."
                    )
                if properties.get(field_name, {}).get("const") != expected_value:
                    raise ArtifactRouteError(
                        f"Route {route.route_id} schema disagrees on {field_name}."
                    )
    for resource in resources:
        schema = _load_json(SCHEMA_DIR / resource.schema_file)
        if schema.get("$id") != resource.schema_id:
            raise ArtifactRouteError(
                f"Schema resource {resource.schema_file} has a mismatched schema_id."
            )
        if resource.fingerprint_inventory != (
            resource.resource_kind == "fingerprint_shared"
        ):
            raise ArtifactRouteError(
                f"Schema resource {resource.schema_file} has an inconsistent inventory class."
            )
        if resource.version_source == "schema_metadata":
            if schema.get("x-profile-schema-version") != resource.schema_version:
                raise ArtifactRouteError(
                    f"Schema resource {resource.schema_file} has a mismatched profile version."
                )
            if (
                schema.get("x-registry-contract-version")
                != resource.registry_contract_version
            ):
                raise ArtifactRouteError(
                    f"Schema resource {resource.schema_file} has a mismatched registry version."
                )
        if resource.resource_kind == "contract_metadata":
            if resource.registry_contract_version != "none" or resource.fingerprint_inventory:
                raise ArtifactRouteError(
                    f"Contract metadata {resource.schema_file} cannot use property inventory."
                )


def load_routed_contracts(
    document: Mapping[str, Any],
    *,
    matrix_version: str = "1.0",
) -> tuple[ContractRoute, dict[str, Any], dict[str, Any], dict[str, Any] | None]:
    """Resolve the route first, then and only then load schema/registry/authority."""

    route = route_artifact_contract(document, matrix_version=matrix_version)
    schema = _load_json(SCHEMA_DIR / route.schema_file)
    try:
        if route.registry_validation_context == "strict_execution":
            registry = load_registry(version=route.registry_contract_version)
            validate_registry_document(registry)
            verify_committed_catalog(
                registry, version=route.registry_contract_version
            )
        else:
            registry = load_registry(
                version=route.registry_contract_version,
                validation_context=route.registry_validation_context,
            )
            validate_registry_document(
                registry, validation_context=route.registry_validation_context
            )
            verify_committed_catalog(
                registry,
                version=route.registry_contract_version,
                validation_context=route.registry_validation_context,
            )
    except RegistryContractError as exc:
        raise ArtifactRouteError(str(exc)) from exc
    authority: dict[str, Any] | None = None
    if route.authority_contract_version != "none":
        try:
            authority = load_authority_contract(route.authority_contract_version)
            verify_authority_projection(route.authority_contract_version)
        except AuthorityContractError as exc:
            raise ArtifactRouteError(str(exc)) from exc
    return route, schema, registry, authority


def profile_v2_schema_enabled(features: Mapping[str, Any] | None) -> bool:
    return bool(features and features.get("profile_v2_schema") is True)


def profile_v2_composer_contract_enabled(
    manifest: dict[str, Any] | None,
) -> bool:
    """Return H contract eligibility without connecting it to a runtime entry point."""

    if manifest is None:
        return False
    try:
        result = validate_artifact(
            manifest,
            features={"profile_v2_schema": True},
        )
    except ArtifactContractError:
        return False
    features = result.document.get("features", {})
    return bool(
        result.artifact_kind == "feature-activation-manifest"
        and result.schema_version == "2.1"
        and features.get("profile_v2_schema") is True
        and features.get("profile_v2_composer") is True
    )


def profile_v2_intent_contract_enabled(
    manifest: dict[str, Any] | None,
) -> bool:
    """Return C1 intent eligibility without connecting it to a runtime entry point."""

    if manifest is None:
        return False
    try:
        result = validate_intent_artifact_v041(manifest)
    except ArtifactContractError:
        return False
    features = result.document.get("features", {})
    return bool(
        result.artifact_kind == "feature-activation-manifest"
        and result.schema_version == "2.2"
        and features.get("profile_v2_schema") is True
        and features.get("profile_v2_composer") is True
        and features.get("monograph_base_v041") is True
        and features.get("final_ready_eligible") is False
    )


def load_artifact_schema(
    artifact_kind: str,
    *,
    version: str = "2.0",
    registry_contract_version: str | None = None,
    authority_contract_version: str | None = None,
    matrix_version: str = "1.0",
) -> dict[str, Any]:
    route_index = {
        (
            item.artifact_kind,
            item.schema_version,
            item.registry_contract_version,
            item.authority_contract_version,
        ): item.schema_file
        for item in _contract_routes(matrix_version)
    }
    matches = [
        (identity, filename)
        for identity, filename in route_index.items()
        if identity[0] == artifact_kind
        and identity[1] == version
        and (registry_contract_version is None or identity[2] == registry_contract_version)
        and (authority_contract_version is None or identity[3] == authority_contract_version)
    ]
    if len(matches) != 1:
        raise ArtifactContractError(
            f"Unsupported or ambiguous V2 artifact contract: {artifact_kind}@{version}"
        )
    filename = matches[0][1]
    return _load_json(SCHEMA_DIR / filename)


def _schema_documents(
    schema_overrides: Mapping[str, dict[str, Any]] | None = None,
    *,
    matrix_version: str = "1.0",
) -> dict[str, dict[str, Any]]:
    documents: dict[str, dict[str, Any]] = {}
    files = {
        item.schema_file for item in _contract_routes(matrix_version)
    } | {
        item.schema_file for item in _schema_resource_contracts(matrix_version)
    }
    for filename in sorted(files):
        path = SCHEMA_DIR / filename
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


def schema_documents(*, matrix_version: str = "1.0") -> dict[str, dict[str, Any]]:
    """Return only the repository's committed offline schema set."""

    return _schema_documents(matrix_version=matrix_version)


def _offline_schema_registry(
    schema_overrides: Mapping[str, dict[str, Any]] | None = None,
    *,
    matrix_version: str = "1.0",
) -> Registry:
    registry = Registry()
    for schema_id, schema in _schema_documents(
        schema_overrides, matrix_version=matrix_version
    ).items():
        registry = registry.with_resource(schema_id, Resource.from_contents(schema))
    return registry


def offline_schema_registry(*, matrix_version: str = "1.0") -> Registry:
    """Build the production resolver from committed local schemas only."""

    return _offline_schema_registry(matrix_version=matrix_version)


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
    validator_override: Draft202012Validator | None = None,
    matrix_version: str = "1.0",
) -> list[str]:
    try:
        schema = schema_override or load_artifact_schema(
            artifact_kind,
            version=str(document.get("schema_version")),
            matrix_version=matrix_version,
        )
    except ArtifactContractError as exc:
        return [str(exc)]
    if validator_override is None:
        Draft202012Validator.check_schema(schema)
        validator = Draft202012Validator(
            schema,
            registry=_offline_schema_registry(
                schema_documents_override, matrix_version=matrix_version
            ),
            format_checker=FormatChecker(),
        )
    else:
        validator = validator_override
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
    production = load_registry(version=registry["schema_version"])
    for collection in TEST_REGISTRY_SHARED_CATALOGS:
        if registry.get(collection) != production.get(collection):
            raise ArtifactContractError(
                f"Test registry must reuse the production {collection} catalog exactly."
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


def _keyed_collection_ids(
    items: list[dict[str, Any]],
    key_field: str,
    context: str,
    errors: list[str],
    *,
    nested_object_field: str | None = None,
) -> set[str]:
    """Collect stable IDs and reject duplicate keys regardless of payload equality."""

    seen: set[str] = set()
    for item in items:
        source = item[nested_object_field] if nested_object_field else item
        value = source[key_field]
        if value in seen:
            errors.append(f"Duplicate {key_field} {value} in {context}.")
        seen.add(value)
    return seen


def artifact_semantic_errors(
    artifact_kind: str,
    document: dict[str, Any],
    registry: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    properties = property_index(registry)
    schema_version = str(document.get("schema_version"))

    _keyed_collection_ids(
        document.get("input_fingerprints", []),
        "input_id",
        "artifact input fingerprints",
        errors,
    )

    if artifact_kind == "capability-snapshot":
        _keyed_collection_ids(
            document.get("capabilities", []),
            "capability_id",
            "capability snapshot",
            errors,
        )
    elif artifact_kind == "legacy-migration-manifest":
        _keyed_collection_ids(
            document.get("mappings", []),
            "source_rule_id",
            "legacy migration mappings",
            errors,
        )
    elif artifact_kind == "execution-evidence-artifact":
        for collection, key_field, context in (
            ("capability_versions", "capability_id", "evidence capability versions"),
            ("measured_results", "metric_id", "evidence measured results"),
            ("history", "entry_id", "evidence history"),
        ):
            _keyed_collection_ids(
                document.get(collection, []), key_field, context, errors
            )

    def validate_key_and_candidate(
        key: dict[str, Any], candidate: dict[str, Any], context: str
    ) -> None:
        if registry.get("schema_version") in {"2.1", "2.2"}:
            try:
                from profile_v2_scope import normalize_scope

                normalize_scope(key["normalized_scope"])
            except ValueError as exc:
                errors.append(f"{context} has invalid normalized scope: {exc}")
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

    def composition_key(
        key: dict[str, Any], context: str
    ) -> tuple[str, str, str] | None:
        try:
            from profile_v2_scope import normalized_property_scope_key

            return normalized_property_scope_key(
                key["semantic_object_kind"],
                key["property_id"],
                key["normalized_scope"],
            )
        except (KeyError, ValueError) as exc:
            errors.append(f"{context} has no valid normalized composition key: {exc}")
            return None

    def inputs_by_role() -> dict[str, list[str]]:
        result: dict[str, list[str]] = {}
        for item in document.get("input_fingerprints", []):
            result.setdefault(item["role"], []).append(item["fingerprint"])
        return result

    def validate_singleton_binding(
        role: str,
        binding_field: str,
        bindings: Mapping[str, Any],
        context: str,
    ) -> None:
        values = inputs_by_role().get(role, [])
        if len(values) != 1:
            errors.append(f"{context} requires exactly one {role} input binding.")
        elif values[0] != bindings.get(binding_field):
            errors.append(f"{context} {role} fingerprint does not match {binding_field}.")

    if artifact_kind == "feature-activation-manifest" and schema_version in {"2.1", "2.2"}:
        features = document.get("features", {})
        if features.get("profile_v2_composer") is True and features.get("profile_v2_schema") is not True:
            errors.append("profile_v2_composer requires profile_v2_schema=true.")
    elif artifact_kind == "qa-approval-artifact" and schema_version in {"2.1", "2.2"}:
        bindings = document.get("bindings", {})
        validate_singleton_binding("source_document", "input_fingerprint", bindings, "QA approval")
        validate_singleton_binding("structure", "structure_fingerprint", bindings, "QA approval")
        validate_singleton_binding("conflict_report", "composition_report_fingerprint", bindings, "QA approval")
        decision_type = document.get("decision_type")
        decision = document.get("decision")
        decisions_by_type = {
            "conflict_resolution": {"adopt_proposed", "select_candidate"},
            "keep_original": {"keep_original"},
            "qa_exclusion": {"exclude_candidate"},
        }
        if decision not in decisions_by_type.get(decision_type, set()):
            errors.append(
                f"QA decision_type {decision_type} is inconsistent with decision {decision}."
            )
        candidate_id = document.get("target", {}).get("candidate_id")
        if decision in {"select_candidate", "exclude_candidate"} and candidate_id is None:
            errors.append(f"QA decision {decision} requires target.candidate_id.")
        if decision in {"adopt_proposed", "keep_original"} and candidate_id is not None:
            errors.append(f"QA decision {decision} cannot bind a candidate_id.")

    if artifact_kind == "layered-rule-asset":
        layer_kind = document["layer_kind"]
        if schema_version == "2.2":
            from profile_v2_canonical import canonical_data_digest

            registry_binding = document.get("property_registry_binding", {})
            registry_fingerprint = registry_binding.get("registry_fingerprint")
            expected_registry_fingerprint = canonical_data_digest(registry)
            if registry_fingerprint != expected_registry_fingerprint:
                errors.append(
                    "Layered intent asset property registry fingerprint does not match "
                    "the routed registry."
                )
            if registry_binding.get("required_completeness") != registry.get(
                "registry_completeness"
            ):
                errors.append(
                    "Layered intent asset registry completeness does not match the "
                    "routed registry."
                )
            registry_inputs = [
                item.get("fingerprint")
                for item in document.get("input_fingerprints", [])
                if item.get("role") == "property_registry"
            ]
            if registry_inputs != [registry_fingerprint]:
                errors.append(
                    "Layered intent asset requires exactly one matching property_registry "
                    "input fingerprint."
                )
        _keyed_collection_ids(
            document.get("rules", []), "rule_id", "layered rules", errors
        )
        for rule in document.get("rules", []):
            rule_id = rule["rule_id"]
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
        if schema_version in {"2.1", "2.2"}:
            try:
                from profile_v2_scope import normalize_rule_scope, validate_module_asset_scope

                normalize_rule_scope(document["asset_scope"])
                for rule in document.get("rules", []):
                    normalize_rule_scope(rule["scope"])
                validate_module_asset_scope(document)
            except ValueError as exc:
                errors.append(str(exc))
    elif artifact_kind == "final-execution-profile" and schema_version in {"2.2", "2.3"}:
        if document.get("legacy_input") is not False:
            errors.append("V2.2 final execution profiles cannot be legacy inputs.")
        if document.get("activation") != "disabled":
            errors.append("V2.2 final execution profiles must remain disabled.")
        if document.get("final_ready_eligible") is not False:
            errors.append("V2.2 final execution profiles cannot be final-ready eligible.")
        if document.get("delivery_allowed") is not False:
            errors.append("V2.2 final execution profiles cannot allow delivery.")
        bindings = document.get("bindings", {})
        for role, field in (
            ("task", "task_fingerprint"),
            ("source_document", "input_fingerprint"),
            ("feature_activation", "feature_activation_fingerprint"),
            ("property_registry", "property_registry_fingerprint"),
            ("authority_contract", "authority_contract_fingerprint"),
            ("structure", "structure_fingerprint"),
            ("conflict_report", "composition_report_fingerprint"),
        ):
            validate_singleton_binding(role, field, bindings, "Final profile")
        for role, field, minimum in (
            ("rule_asset", "rule_asset_fingerprints", 1),
            ("approval", "approval_fingerprints", 0),
        ):
            values = inputs_by_role().get(role, [])
            if len(values) < minimum:
                errors.append(f"Final profile requires at least {minimum} {role} inputs.")
            elif len(values) != len(set(values)):
                errors.append(f"Final profile has duplicate {role} fingerprints.")
            elif sorted(values) != sorted(bindings.get(field, [])):
                errors.append(f"Final profile {role} fingerprints do not match {field}.")
        try:
            expected_authority = authority_contract_fingerprint("1.0")
            if bindings.get("authority_contract_fingerprint") != expected_authority:
                errors.append("Final profile authority fingerprint is not authority contract 1.0.")
        except AuthorityContractError as exc:
            errors.append(str(exc))
        _keyed_collection_ids(
            document.get("resolved_properties", []), "resolution_id", "V2.2 resolved properties", errors
        )
        _keyed_collection_ids(
            document.get("closure_evidence", []), "conflict_id", "V2.2 closure evidence", errors
        )
        seen_keys: set[tuple[str, str, str]] = set()
        for resolved in document.get("resolved_properties", []):
            key = resolved["key"]
            normalized_key = composition_key(key, "V2.2 resolved property")
            if normalized_key is not None and normalized_key in seen_keys:
                errors.append("V2.2 final profile repeats a normalized composition key.")
            if normalized_key is not None:
                seen_keys.add(normalized_key)
            binding = resolved["resolved_binding"]
            if resolved["execution_mode"] != binding["mode"]:
                errors.append("V2.2 execution_mode does not match resolved binding mode.")
            validate_key_and_candidate(
                key,
                {"property_binding": binding, "layer_kind": resolved["final_layer_kind"]},
                "V2.2 resolved property",
            )
            active_ids = _keyed_collection_ids(
                resolved["candidate_chain"], "candidate_id", "V2.2 active candidates", errors
            )
            excluded_ids = _keyed_collection_ids(
                resolved["excluded_candidates"], "candidate_id", "V2.2 excluded candidates", errors,
                nested_object_field="candidate",
            )
            if active_ids & excluded_ids:
                errors.append("V2.2 candidate IDs cannot be both active and excluded.")
            final_source_matches = False
            for candidate in resolved["candidate_chain"]:
                validate_key_and_candidate(key, candidate, "V2.2 resolved property")
                if candidate["scope_status"] != "applicable":
                    errors.append("V2.2 active candidates must be applicable.")
                if (
                    candidate["source"] == resolved["final_source"]
                    and candidate["layer_kind"] == resolved["final_layer_kind"]
                    and candidate["property_binding"] == binding
                ):
                    final_source_matches = True
            for item in resolved["excluded_candidates"]:
                excluded = item["candidate"]
                validate_key_and_candidate(key, excluded, "V2.2 excluded candidate")
                if excluded["scope_status"] == "applicable":
                    errors.append("V2.2 excluded candidates cannot be applicable.")
            if set(resolved["override_chain"]) - active_ids:
                errors.append("V2.2 override_chain references an unknown active candidate.")
            if not final_source_matches:
                errors.append(
                    "V2.2 final source/layer/binding is absent from its candidate chain."
                )
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

        _keyed_collection_ids(
            document.get("resolved_properties", []),
            "resolution_id",
            "final resolved properties",
            errors,
        )
        seen_composition_keys: set[tuple[str, str, str]] = set()
        for resolved in document.get("resolved_properties", []):
            resolution_id = resolved["resolution_id"]
            key = resolved["key"]
            resolved_key = composition_key(key, "Resolved property")
            if resolved_key is not None and resolved_key in seen_composition_keys:
                errors.append("Final profile repeats a semantic object/property/scope composition key.")
            if resolved_key is not None:
                seen_composition_keys.add(resolved_key)
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
            active_ids = _keyed_collection_ids(
                resolved["candidate_chain"],
                "candidate_id",
                f"resolved property {resolution_id} active candidates",
                errors,
            )
            excluded_ids = _keyed_collection_ids(
                resolved["excluded_candidates"],
                "candidate_id",
                f"resolved property {resolution_id} excluded candidates",
                errors,
                nested_object_field="candidate",
            )
            applicable_ids: set[str] = set()
            final_source_matches = False
            for candidate in resolved["candidate_chain"]:
                candidate_id = candidate["candidate_id"]
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
    elif artifact_kind == "conflict-report" and schema_version in {"2.2", "2.3"}:
        bindings = document.get("bindings", {})
        for role, field in (
            ("source_document", "input_fingerprint"),
            ("feature_activation", "feature_activation_fingerprint"),
            ("property_registry", "property_registry_fingerprint"),
            ("authority_contract", "authority_contract_fingerprint"),
            ("structure", "structure_fingerprint"),
        ):
            validate_singleton_binding(role, field, bindings, "Composition report")
        rule_values = inputs_by_role().get("rule_asset", [])
        if not rule_values or len(rule_values) != len(set(rule_values)):
            errors.append("Composition report requires unique rule_asset inputs.")
        elif sorted(rule_values) != sorted(bindings.get("rule_asset_fingerprints", [])):
            errors.append("Composition report rule assets do not match bindings.")
        try:
            expected_authority = authority_contract_fingerprint("1.0")
            if bindings.get("authority_contract_fingerprint") != expected_authority:
                errors.append("Composition report authority fingerprint is not contract 1.0.")
        except AuthorityContractError as exc:
            errors.append(str(exc))
        keyed = (
            ("candidate_groups", "candidate_group_id"),
            ("scope_partitions", "partition_id"),
            ("proposed_resolutions", "proposed_resolution_id"),
            ("fatal_diagnostics", "diagnostic_id"),
            ("unresolvable_blockers", "blocker_id"),
            ("approval_required_conflicts", "conflict_id"),
            ("diagnostics", "diagnostic_id"),
        )
        for collection, key_field in keyed:
            _keyed_collection_ids(document.get(collection, []), key_field, collection, errors)
        composition_maps: dict[str, dict[tuple[str, str, str], dict[str, Any]]] = {}
        for collection in ("candidate_groups", "scope_partitions", "proposed_resolutions"):
            keyed_items: dict[tuple[str, str, str], dict[str, Any]] = {}
            for item in document.get(collection, []):
                normalized = composition_key(item["key"], f"Composition report {collection}")
                if normalized is None:
                    continue
                if normalized in keyed_items:
                    errors.append(
                        f"Composition report {collection} repeats a normalized composition key."
                    )
                keyed_items[normalized] = item
            composition_maps[collection] = keyed_items
        proposal_by_id = {
            item["proposed_resolution_id"]: item
            for item in document.get("proposed_resolutions", [])
        }
        for group in document.get("candidate_groups", []):
            key = group["key"]
            active_ids = _keyed_collection_ids(
                group["candidates"], "candidate_id", "candidate group", errors
            )
            excluded_ids = _keyed_collection_ids(
                group["excluded_candidates"], "candidate_id", "excluded candidate group", errors,
                nested_object_field="candidate",
            )
            if active_ids & excluded_ids:
                errors.append("Candidate group active and excluded IDs overlap.")
            for candidate in group["candidates"]:
                validate_key_and_candidate(key, candidate, "Candidate group")
                if candidate["scope_status"] != "applicable":
                    errors.append("Validated candidate groups contain only applicable candidates.")
            for item in group["excluded_candidates"]:
                excluded = item["candidate"]
                validate_key_and_candidate(key, excluded, "Excluded candidate group")
                if excluded["scope_status"] == "applicable":
                    errors.append("Excluded candidate-group candidates cannot be applicable.")
        for proposal in document.get("proposed_resolutions", []):
            key = proposal["key"]
            normalized = composition_key(key, "Proposed resolution")
            group = (
                composition_maps["candidate_groups"].get(normalized)
                if normalized is not None
                else None
            )
            if group is None:
                errors.append("Proposed resolution has no candidate group for its normalized key.")
                continue
            validate_key_and_candidate(
                key,
                {
                    "property_binding": proposal["proposed_binding"],
                    "layer_kind": proposal["final_layer_kind"],
                },
                "Proposed resolution",
            )
            if proposal["execution_mode"] != proposal["proposed_binding"]["mode"]:
                errors.append("Proposed resolution execution_mode differs from its binding mode.")
            group_candidates = {item["candidate_id"]: item for item in group["candidates"]}
            chain_ids = _keyed_collection_ids(
                proposal["candidate_chain"],
                "candidate_id",
                "proposed resolution candidate chain",
                errors,
            )
            chain_candidates = {
                item["candidate_id"]: item for item in proposal["candidate_chain"]
            }
            for candidate in proposal["candidate_chain"]:
                validate_key_and_candidate(key, candidate, "Proposed resolution")
                if candidate["scope_status"] != "applicable":
                    errors.append("Proposed-resolution candidates must be applicable.")
            if chain_ids != set(group_candidates) or any(
                chain_candidates.get(candidate_id) != candidate
                for candidate_id, candidate in group_candidates.items()
            ):
                errors.append(
                    "Proposed resolution candidate_chain does not exactly reference its candidate group."
                )
            if set(proposal["override_chain"]) - chain_ids:
                errors.append("Proposed resolution override_chain references a non-active candidate.")
            final_source_matches = any(
                candidate["source"] == proposal["final_source"]
                and candidate["layer_kind"] == proposal["final_layer_kind"]
                and candidate["property_binding"] == proposal["proposed_binding"]
                for candidate in proposal["candidate_chain"]
            )
            if not final_source_matches:
                errors.append(
                    "Proposed resolution final source/layer/binding is absent from its candidate chain."
                )
        for conflict in document.get("approval_required_conflicts", []):
            proposal = proposal_by_id.get(conflict["proposed_resolution_id"])
            if proposal is None:
                errors.append("Approval-required conflict references an unknown proposed resolution.")
            else:
                conflict_key = composition_key(conflict["key"], "Approval-required conflict")
                proposal_key = composition_key(proposal["key"], "Conflict proposal")
                if conflict_key is not None and proposal_key is not None and conflict_key != proposal_key:
                    errors.append(
                        "Approval-required conflict and proposed resolution use different normalized keys."
                    )
            candidate_ids = _keyed_collection_ids(
                conflict["candidates"], "candidate_id", "approval-required candidates", errors
            )
            for candidate in conflict["candidates"]:
                validate_key_and_candidate(conflict["key"], candidate, "Approval-required conflict")
            conflict_key = composition_key(conflict["key"], "Approval-required conflict")
            group = (
                composition_maps["candidate_groups"].get(conflict_key)
                if conflict_key is not None
                else None
            )
            if group is None:
                errors.append("Approval-required conflict has no candidate group for its key.")
            else:
                group_candidates = {item["candidate_id"]: item for item in group["candidates"]}
                conflict_candidates = {
                    item["candidate_id"]: item for item in conflict["candidates"]
                }
                if candidate_ids != set(group_candidates) or any(
                    conflict_candidates.get(candidate_id) != candidate
                    for candidate_id, candidate in group_candidates.items()
                ):
                    errors.append(
                        "Approval-required conflict does not carry the complete candidate group."
                    )
            if "select_candidate" in conflict["allowed_decisions"] and not candidate_ids:
                errors.append("select_candidate requires a non-empty candidate set.")
        blocker_keys = {
            normalized
            for item in document.get("unresolvable_blockers", [])
            if (normalized := composition_key(item["key"], "Unresolvable blocker"))
            is not None
        }
        for partition in document.get("scope_partitions", []):
            partition_key = composition_key(partition["key"], "Scope partition")
            if (
                partition["evidence_status"] in {"failed", "not_evaluated"}
                and partition_key not in blocker_keys
            ):
                errors.append(
                    "Failed or not-evaluated scope partitions require an unresolvable "
                    "blocker for the same key."
                )
        if schema_version == "2.3":
            from profile_v2_canonical import canonical_data_digest
            from profile_v2_scope import scope_equal, scope_subset

            coverage = document.get("coverage_evidence", {})
            expected_items = coverage.get("expected_bindings", [])
            consumed_items = coverage.get("consumed_bindings", [])
            expected_ids = _keyed_collection_ids(
                expected_items, "binding_id", "coverage expected bindings", errors
            )
            consumed_ids = _keyed_collection_ids(
                consumed_items, "binding_id", "coverage consumed bindings", errors
            )
            if coverage.get("expected_binding_count") != len(expected_items):
                errors.append("Coverage expected_binding_count does not match its inventory.")
            if coverage.get("consumed_binding_count") != len(consumed_items):
                errors.append("Coverage consumed_binding_count does not match its inventory.")
            if expected_ids != consumed_ids:
                errors.append(
                    "Coverage consumed bindings must close every and only expected binding."
                )
            ordered_expected = sorted(
                expected_items, key=lambda item: item.get("binding_id", "")
            )
            ordered_consumed = sorted(
                consumed_items, key=lambda item: item.get("binding_id", "")
            )
            if coverage.get("expected_inventory_digest") != canonical_data_digest(
                ordered_expected
            ):
                errors.append("Coverage expected inventory digest is inconsistent.")
            if coverage.get("consumed_inventory_digest") != canonical_data_digest(
                ordered_consumed
            ):
                errors.append("Coverage consumed inventory digest is inconsistent.")

            expected_by_id = {
                item["binding_id"]: item
                for item in expected_items
                if isinstance(item, dict) and "binding_id" in item
            }
            artifact_fingerprints: dict[str, str] = {}
            bound_rule_fingerprints = set(
                bindings.get("rule_asset_fingerprints", [])
            )
            for item in expected_items:
                key = item.get("key", {})
                if (
                    key.get("semantic_object_kind") != item.get("semantic_object_kind")
                    or key.get("property_id") != item.get("property_id")
                    or not scope_equal(
                        key.get("normalized_scope", {}),
                        item.get("normalized_scope", {}),
                    )
                ):
                    errors.append(
                        "Coverage expected binding key differs from its object/property/scope."
                    )
                source_id = item.get("source_artifact_id")
                source_fingerprint = item.get("source_artifact_fingerprint")
                previous = artifact_fingerprints.setdefault(source_id, source_fingerprint)
                if previous != source_fingerprint:
                    errors.append(
                        "Coverage maps one source artifact to multiple fingerprints."
                    )
                if source_fingerprint not in bound_rule_fingerprints:
                    errors.append(
                        "Coverage expected binding references an unbound rule asset fingerprint."
                    )

            candidate_locations: dict[str, dict[str, set[str]]] = {}
            proposals_by_key = composition_maps.get("proposed_resolutions", {})
            for group in document.get("candidate_groups", []):
                group_key = composition_key(group["key"], "Coverage candidate group")
                proposal = proposals_by_key.get(group_key)
                if proposal is None:
                    errors.append("Coverage candidate group has no proposed resolution.")
                    continue
                for candidate in group.get("candidates", []):
                    candidate_id = candidate["candidate_id"]
                    expected = expected_by_id.get(candidate_id)
                    if expected is None:
                        errors.append("Coverage candidate has no expected binding source.")
                        continue
                    if (
                        candidate["source"].get("source_artifact_id")
                        != expected.get("source_artifact_id")
                        or candidate["source"].get("source_rule_id")
                        != expected.get("source_rule_id")
                        or candidate.get("layer_kind") != expected.get("layer_kind")
                        or candidate["property_binding"].get("property_id")
                        != expected.get("property_id")
                        or canonical_data_digest(candidate["property_binding"])
                        != expected.get("property_binding_digest")
                    ):
                        errors.append(
                            "Coverage candidate provenance/layer/property/binding is inconsistent."
                        )
                    if (
                        group["key"].get("semantic_object_kind")
                        != expected.get("semantic_object_kind")
                        or group["key"].get("property_id")
                        != expected.get("property_id")
                        or not scope_subset(
                            group["key"].get("normalized_scope", {}),
                            expected.get("normalized_scope", {}),
                        )
                    ):
                        errors.append(
                            "Coverage candidate group key is outside its expected binding scope."
                        )
                    location = candidate_locations.setdefault(
                        candidate_id,
                        {"candidate_group_ids": set(), "proposed_resolution_ids": set()},
                    )
                    location["candidate_group_ids"].add(group["candidate_group_id"])
                    location["proposed_resolution_ids"].add(
                        proposal["proposed_resolution_id"]
                    )

            blocker_locations: dict[str, set[str]] = {}
            for blocker in document.get("unresolvable_blockers", []):
                for candidate_id in blocker.get("candidate_ids", []):
                    if candidate_id not in expected_by_id:
                        errors.append("Coverage blocker references an unexpected binding.")
                    blocker_locations.setdefault(candidate_id, set()).add(
                        blocker["blocker_id"]
                    )
            if set(candidate_locations) & set(blocker_locations):
                errors.append(
                    "Coverage binding cannot be both candidate/proposal and blocker consumed."
                )
            if any(len(values) != 1 for values in blocker_locations.values()):
                errors.append("Coverage binding cannot be consumed by multiple blockers.")

            computed_consumed: dict[str, dict[str, Any]] = {}
            for binding_id in sorted(expected_by_id):
                if binding_id in candidate_locations:
                    location = candidate_locations[binding_id]
                    computed_consumed[binding_id] = {
                        "binding_id": binding_id,
                        "classification": "candidate_proposal",
                        "candidate_group_ids": sorted(location["candidate_group_ids"]),
                        "proposed_resolution_ids": sorted(
                            location["proposed_resolution_ids"]
                        ),
                        "blocker_ids": [],
                    }
                elif binding_id in blocker_locations:
                    computed_consumed[binding_id] = {
                        "binding_id": binding_id,
                        "classification": "unresolvable_blocker",
                        "candidate_group_ids": [],
                        "proposed_resolution_ids": [],
                        "blocker_ids": sorted(blocker_locations[binding_id]),
                    }
            if set(computed_consumed) != set(expected_by_id):
                errors.append("Coverage has expected bindings with no consumption record.")
            supplied_consumed = {
                item["binding_id"]: item
                for item in consumed_items
                if isinstance(item, dict) and "binding_id" in item
            }
            if supplied_consumed != computed_consumed:
                errors.append(
                    "Coverage consumed classifications do not match report candidates, "
                    "proposals, and blockers."
                )
        fatal = bool(document.get("fatal_diagnostics"))
        unresolvable = bool(document.get("unresolvable_blockers"))
        approvals = bool(document.get("approval_required_conflicts"))
        expected_status = (
            "fatal" if fatal else "unresolvable" if unresolvable else "awaiting_approval" if approvals else "resolvable"
        )
        if document.get("proposal_status") != expected_status:
            errors.append("Composition report proposal_status does not match its blocker collections.")
    elif artifact_kind == "conflict-report":
        _keyed_collection_ids(
            document.get("conflicts", []),
            "conflict_id",
            "conflict report",
            errors,
        )
        seen_conflict_keys: set[tuple[str, str, str]] = set()
        for conflict in document.get("conflicts", []):
            key = conflict["key"]
            conflict_id = conflict["conflict_id"]
            conflict_key = composition_key(key, f"Conflict {conflict_id}")
            if conflict_key is not None and conflict_key in seen_conflict_keys:
                errors.append(
                    "Conflict report repeats a semantic object/property/scope composition key."
                )
            if conflict_key is not None:
                seen_conflict_keys.add(conflict_key)
            active_ids = _keyed_collection_ids(
                conflict["candidates"],
                "candidate_id",
                f"conflict {conflict_id} active candidates",
                errors,
            )
            excluded_ids = _keyed_collection_ids(
                conflict["excluded_candidates"],
                "candidate_id",
                f"conflict {conflict_id} excluded candidates",
                errors,
                nested_object_field="candidate",
            )
            if active_ids & excluded_ids:
                errors.append(
                    f"Conflict {conflict_id} candidate IDs cannot be both active and excluded."
                )
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
    resolved_schema: dict[str, Any] | None = None,
    schema_override: dict[str, Any] | None = None,
    schema_documents_override: Mapping[str, dict[str, Any]] | None = None,
    validator_override: Draft202012Validator | None = None,
    matrix_version: str = "1.0",
) -> ProfileReadResult:
    if not profile_v2_schema_enabled(features):
        raise ProfileV2DisabledError(
            "profile_v2_schema is disabled by default; enable it only for the P1 schema path."
        )
    artifact_kind = document.get("artifact_kind")
    if artifact_kind not in ARTIFACT_KINDS:
        raise ArtifactContractError(f"Unknown V2 artifact_kind: {artifact_kind}")
    if resolved_schema is not None and schema_override is not None:
        raise ArtifactContractError(
            "Resolved production schemas and test-only schema overrides are mutually exclusive."
        )
    if resolved_schema is not None:
        effective_schema = resolved_schema
        compatible_minor = False
    elif schema_override is None:
        effective_schema = load_artifact_schema(
            artifact_kind,
            version=str(document.get("schema_version")),
            matrix_version=matrix_version,
        )
        compatible_minor = False
    else:
        effective_schema, compatible_minor = schema_for_requested_minor(
            schema_override, document.get("schema_version")
        )
    errors = _schema_errors(
        artifact_kind,
        document,
        schema_override=effective_schema,
        schema_documents_override=schema_documents_override,
        validator_override=validator_override,
        matrix_version=matrix_version,
    )
    if not errors:
        errors.extend(artifact_semantic_errors(artifact_kind, document, registry))
    if not errors:
        try:
            if registry.get("registry_scope") == "test":
                from profile_v2_canonical import _compute_semantic_fingerprint_for_test

                documents = _schema_documents(schema_documents_override)
                expected = _compute_semantic_fingerprint_for_test(
                    document,
                    schema=effective_schema,
                    documents=documents,
                    registry=registry,
                )
                if document.get("semantic_fingerprint") != expected:
                    errors.append(
                        "semantic_fingerprint does not match canonical test semantics."
                    )
            else:
                if matrix_version == "1.1":
                    from profile_v2_canonical import (
                        verify_intent_semantic_fingerprint_v041,
                    )

                    verify_intent_semantic_fingerprint_v041(document)
                else:
                    from profile_v2_canonical import verify_semantic_fingerprint

                    verify_semantic_fingerprint(document)
        except ValueError as exc:
            errors.append(str(exc))
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

    verify_contract_matrix_alignment()
    _, effective_schema, effective_registry, _ = load_routed_contracts(document)
    return _validate_artifact_contract(
        document,
        features=features,
        registry=effective_registry,
        resolved_schema=effective_schema,
    )


def _session_digest(value: Any) -> str:
    """Content-bind one private validation session without a process-wide cache."""

    payload = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(payload).hexdigest()


class _IntentValidationSessionV041:
    """One compose/apply-local set of immutable validation preparation resources."""

    def __init__(self) -> None:
        self._closed = False
        verify_contract_matrix_alignment(matrix_version="1.1")
        self._matrix = load_artifact_contract_matrix("1.1")
        self._matrix_digest = _session_digest(self._matrix)
        self._routes = _contract_routes_from_matrix(self._matrix)
        self._schemas = _schema_documents(matrix_version="1.1")
        self._schema_documents_digest = _session_digest(self._schemas)
        registry = Registry()
        for schema_id, schema in self._schemas.items():
            registry = registry.with_resource(schema_id, Resource.from_contents(schema))
        self._offline_registry = registry
        self._registries: dict[tuple[str, str], tuple[dict[str, Any], str]] = {}
        self._authorities: dict[str, tuple[dict[str, Any], str]] = {}
        self._validators: dict[
            tuple[str, str, str, str, str, str, str, str], Draft202012Validator
        ] = {}

    @property
    def registry(self) -> dict[str, Any]:
        return self._registry_for("2.2", "declaration_intent")[0]

    def close(self) -> None:
        self._closed = True
        self._validators.clear()
        self._registries.clear()
        self._authorities.clear()

    def _require_open(self) -> None:
        if self._closed:
            raise ArtifactContractError("Intent validation session is closed.")

    def _registry_for(self, version: str, context: str) -> tuple[dict[str, Any], str]:
        self._require_open()
        key = (version, context)
        cached = self._registries.get(key)
        if cached is None:
            try:
                registry = load_registry(version=version, validation_context=context)
                validate_registry_document(registry, validation_context=context)
                verify_committed_catalog(registry, version=version, validation_context=context)
            except RegistryContractError as exc:
                raise ArtifactRouteError(str(exc)) from exc
            cached = (registry, _session_digest(registry))
            self._registries[key] = cached
        return cached

    def _authority_for(self, version: str) -> tuple[dict[str, Any], str]:
        self._require_open()
        cached = self._authorities.get(version)
        if cached is None:
            try:
                authority = load_authority_contract(version)
                verify_authority_projection(version)
            except AuthorityContractError as exc:
                raise ArtifactRouteError(str(exc)) from exc
            cached = (authority, _session_digest(authority))
            self._authorities[version] = cached
        return cached

    def _resources_for(
        self, document: Mapping[str, Any]
    ) -> tuple[ContractRoute, dict[str, Any], dict[str, Any], Draft202012Validator]:
        self._require_open()
        route = _route_artifact_contract_from_routes(document, self._routes)
        try:
            schema = self._schemas[route.schema_id]
        except KeyError as exc:
            raise ArtifactRouteError(
                f"Route {route.route_id} has no committed schema document."
            ) from exc
        registry, registry_digest = self._registry_for(
            route.registry_contract_version, route.registry_validation_context
        )
        if route.authority_contract_version == "none":
            authority_digest = "none"
        else:
            _, authority_digest = self._authority_for(route.authority_contract_version)
        validator_key = (
            route.artifact_kind,
            route.schema_version,
            route.registry_contract_version,
            route.authority_contract_version,
            self._matrix_digest,
            self._schema_documents_digest,
            registry_digest,
            authority_digest,
        )
        validator = self._validators.get(validator_key)
        if validator is None:
            Draft202012Validator.check_schema(schema)
            validator = Draft202012Validator(
                schema,
                registry=self._offline_registry,
                format_checker=FormatChecker(),
            )
            self._validators[validator_key] = validator
        return route, schema, registry, validator

    def validate(self, document: dict[str, Any]) -> ProfileReadResult:
        """Fully validate one independent document with this session's resources."""

        self._require_open()
        _, schema, registry, validator = self._resources_for(document)
        return _validate_artifact_contract(
            document,
            features={"profile_v2_schema": True},
            registry=registry,
            resolved_schema=schema,
            validator_override=validator,
            matrix_version="1.1",
        )


def _new_intent_validation_session_v041() -> _IntentValidationSessionV041:
    """Create a compose/apply-local private session; callers must close it."""

    return _IntentValidationSessionV041()


def validate_intent_artifact_v041(document: dict[str, Any]) -> ProfileReadResult:
    """Validate one append-only P3 declaration/intent artifact contract."""

    session = _new_intent_validation_session_v041()
    try:
        return session.validate(document)
    finally:
        session.close()


def _is_obvious_placeholder_fingerprint(value: Any) -> bool:
    if not isinstance(value, str) or not re.fullmatch(r"sha256:[0-9a-f]{64}", value):
        return False
    digest = value.split(":", 1)[1]
    return len(set(digest)) == 1


def _validate_artifact_dag_with_reader(
    documents: list[dict[str, Any]],
    reader: Callable[[dict[str, Any]], ProfileReadResult],
    *,
    matrix_version: str = "1.0",
) -> DagValidationResult:
    """Validate H bindings and closure evidence without composing any result."""

    if not documents:
        raise ArtifactDagError("Artifact DAG cannot be empty.")
    validated: list[dict[str, Any]] = []
    for document in documents:
        result = reader(document)
        validated.append(result.document)
    by_id: dict[str, dict[str, Any]] = {}
    by_fingerprint: dict[str, dict[str, Any]] = {}
    for document in validated:
        artifact_id = document["artifact_id"]
        fingerprint = document["semantic_fingerprint"]
        if artifact_id in by_id:
            raise ArtifactDagError(f"Artifact DAG repeats artifact_id {artifact_id}.")
        if fingerprint in by_fingerprint:
            raise ArtifactDagError("Artifact DAG repeats a semantic fingerprint.")
        if str(document.get("schema_version")) in {"2.1", "2.2", "2.3"} and _is_obvious_placeholder_fingerprint(fingerprint):
            raise ArtifactDagError("New artifact contracts reject placeholder semantic fingerprints.")
        by_id[artifact_id] = document
        by_fingerprint[fingerprint] = document

    dependencies: dict[str, set[str]] = {artifact_id: set() for artifact_id in by_id}

    def require_internal(
        owner: dict[str, Any],
        fingerprint: str,
        expected_kind: str,
        *,
        expected_schema_version: str | None = None,
        expected_registry_contract_version: str | None = None,
    ) -> dict[str, Any]:
        target = by_fingerprint.get(fingerprint)
        if target is None:
            raise ArtifactDagError(
                f"{owner['artifact_id']} has a missing internal {expected_kind} fingerprint."
            )
        if target["artifact_kind"] != expected_kind:
            raise ArtifactDagError(
                f"{owner['artifact_id']} binds {expected_kind} to {target['artifact_kind']}."
            )
        if (
            expected_schema_version is not None
            and str(target.get("schema_version")) != expected_schema_version
        ):
            raise ArtifactDagError(
                f"{owner['artifact_id']} requires {expected_kind} schema "
                f"{expected_schema_version}."
            )
        if expected_registry_contract_version is not None:
            target_route = route_artifact_contract(
                target, matrix_version=matrix_version
            )
            if target_route.registry_contract_version != expected_registry_contract_version:
                raise ArtifactDagError(
                    f"{owner['artifact_id']} requires {expected_kind} registry contract "
                    f"{expected_registry_contract_version}."
                )
        dependencies[owner["artifact_id"]].add(target["artifact_id"])
        return target

    reports: dict[str, dict[str, Any]] = {}
    approvals: dict[str, dict[str, Any]] = {}
    finals: list[dict[str, Any]] = []
    for document in validated:
        kind = document["artifact_kind"]
        version = str(document.get("schema_version"))
        if kind == "conflict-report" and version in {"2.2", "2.3"}:
            reports[document["semantic_fingerprint"]] = document
            bindings = document["bindings"]
            intent_route = version == "2.3"
            feature_version = "2.2" if intent_route else "2.1"
            registry_version = "2.2" if intent_route else "2.1"
            asset_version = "2.2" if intent_route else "2.1"
            feature = require_internal(
                document,
                bindings["feature_activation_fingerprint"],
                "feature-activation-manifest",
                expected_schema_version=feature_version,
                expected_registry_contract_version=registry_version,
            )
            if (
                feature.get("features", {}).get("profile_v2_schema") is not True
                or feature.get("features", {}).get("profile_v2_composer") is not True
                or (
                    intent_route
                    and feature.get("features", {}).get("monograph_base_v041")
                    is not True
                )
            ):
                raise ArtifactDagError(
                    "Composition report requires profile_v2_schema=true and "
                    "profile_v2_composer=true; intent reports also require "
                    "monograph_base_v041=true."
                )
            for fingerprint in bindings["rule_asset_fingerprints"]:
                require_internal(
                    document,
                    fingerprint,
                    "layered-rule-asset",
                    expected_schema_version=asset_version,
                    expected_registry_contract_version=registry_version,
                )
        elif kind == "qa-approval-artifact" and version in {"2.1", "2.2"}:
            approvals[document["semantic_fingerprint"]] = document
            intent_route = version == "2.2"
            require_internal(
                document,
                document["bindings"]["composition_report_fingerprint"],
                "conflict-report",
                expected_schema_version="2.3" if intent_route else "2.2",
                expected_registry_contract_version="2.2" if intent_route else "2.1",
            )
        elif kind == "final-execution-profile" and version in {"2.2", "2.3"}:
            finals.append(document)
            bindings = document["bindings"]
            intent_route = version == "2.3"
            report_version = "2.3" if intent_route else "2.2"
            feature_version = "2.2" if intent_route else "2.1"
            approval_version = "2.2" if intent_route else "2.1"
            asset_version = "2.2" if intent_route else "2.1"
            registry_version = "2.2" if intent_route else "2.1"
            require_internal(
                document,
                bindings["composition_report_fingerprint"],
                "conflict-report",
                expected_schema_version=report_version,
                expected_registry_contract_version=registry_version,
            )
            feature = require_internal(
                document,
                bindings["feature_activation_fingerprint"],
                "feature-activation-manifest",
                expected_schema_version=feature_version,
                expected_registry_contract_version=registry_version,
            )
            if (
                feature.get("features", {}).get("profile_v2_schema") is not True
                or feature.get("features", {}).get("profile_v2_composer") is not True
                or (
                    intent_route
                    and feature.get("features", {}).get("monograph_base_v041")
                    is not True
                )
            ):
                raise ArtifactDagError(
                    "Final profile requires profile_v2_schema=true and "
                    "profile_v2_composer=true; intent profiles also require "
                    "monograph_base_v041=true."
                )
            for fingerprint in bindings["rule_asset_fingerprints"]:
                require_internal(
                    document,
                    fingerprint,
                    "layered-rule-asset",
                    expected_schema_version=asset_version,
                    expected_registry_contract_version=registry_version,
                )
            for fingerprint in bindings["approval_fingerprints"]:
                require_internal(
                    document,
                    fingerprint,
                    "qa-approval-artifact",
                    expected_schema_version=approval_version,
                    expected_registry_contract_version=registry_version,
                )

    from profile_v2_scope import normalized_property_scope_key, scope_equal

    for approval in approvals.values():
        report_fingerprint = approval["bindings"]["composition_report_fingerprint"]
        report = reports.get(report_fingerprint)
        if report is None:
            raise ArtifactDagError("QA approval binds a report outside the validated DAG.")
        if approval["bindings"]["input_fingerprint"] != report["bindings"]["input_fingerprint"]:
            raise ArtifactDagError("QA approval input binding differs from its report.")
        if approval["bindings"]["structure_fingerprint"] != report["bindings"]["structure_fingerprint"]:
            raise ArtifactDagError("QA approval structure binding differs from its report.")
        target = approval["target"]
        conflicts = {
            item["conflict_id"]: item for item in report["approval_required_conflicts"]
        }
        conflict = conflicts.get(target["conflict_id"])
        if conflict is None:
            fatal_ids = {item["diagnostic_id"] for item in report["fatal_diagnostics"]}
            blocker_ids = {item["blocker_id"] for item in report["unresolvable_blockers"]}
            if target["conflict_id"] in fatal_ids:
                raise ArtifactDagError("Fatal diagnostics cannot accept QA closure.")
            if target["conflict_id"] in blocker_ids:
                raise ArtifactDagError("Unresolvable blockers cannot accept QA closure.")
            raise ArtifactDagError("QA approval references an unknown conflict_id.")
        if target["proposed_resolution_id"] != conflict["proposed_resolution_id"]:
            raise ArtifactDagError("QA approval references an unknown proposed_resolution_id.")
        if not scope_equal(target["normalized_scope"], conflict["key"]["normalized_scope"]):
            raise ArtifactDagError("QA approval scope differs from the report conflict scope.")
        if approval["decision"] not in conflict["allowed_decisions"]:
            raise ArtifactDagError("QA approval decision is not allowed by the report.")
        candidate_id = target.get("candidate_id")
        candidate_ids = {item["candidate_id"] for item in conflict["candidates"]}
        if candidate_id is not None and candidate_id not in candidate_ids:
            raise ArtifactDagError("QA approval candidate_id is not part of the conflict.")

    for final in finals:
        bindings = final["bindings"]
        report = reports.get(bindings["composition_report_fingerprint"])
        if report is None:
            raise ArtifactDagError("Final profile binds a report outside the validated DAG.")
        if report["fatal_diagnostics"]:
            raise ArtifactDagError("A report with fatal diagnostics cannot have a final profile.")
        if report["unresolvable_blockers"]:
            raise ArtifactDagError("A report with unresolvable blockers cannot have a final profile.")
        if any(
            partition["evidence_status"] in {"failed", "not_evaluated"}
            for partition in report["scope_partitions"]
        ):
            raise ArtifactDagError(
                "A report with failed or not-evaluated scope partitions cannot have a final profile."
            )
        for field in (
            "input_fingerprint",
            "feature_activation_fingerprint",
            "property_registry_fingerprint",
            "authority_contract_fingerprint",
            "structure_fingerprint",
        ):
            if bindings[field] != report["bindings"][field]:
                raise ArtifactDagError(
                    f"Final profile {field} differs from its composition report."
                )
        if sorted(bindings["rule_asset_fingerprints"]) != sorted(
            report["bindings"]["rule_asset_fingerprints"]
        ):
            raise ArtifactDagError(
                "Final profile rule assets differ from its composition report."
            )
        proposals = {
            item["proposed_resolution_id"]: item
            for item in report["proposed_resolutions"]
        }
        resolved = {
            item["resolution_id"]: item for item in final["resolved_properties"]
        }
        if set(resolved) != set(proposals):
            raise ArtifactDagError(
                "Final profile must express every and only report proposed resolution."
            )
        for proposed_resolution_id, proposal in proposals.items():
            final_item = resolved[proposed_resolution_id]
            proposal_key = normalized_property_scope_key(
                proposal["key"]["semantic_object_kind"],
                proposal["key"]["property_id"],
                proposal["key"]["normalized_scope"],
            )
            final_key = normalized_property_scope_key(
                final_item["key"]["semantic_object_kind"],
                final_item["key"]["property_id"],
                final_item["key"]["normalized_scope"],
            )
            if final_key != proposal_key:
                raise ArtifactDagError(
                    "Final resolved property key differs from its report proposal."
                )
        required = {item["conflict_id"]: item for item in report["approval_required_conflicts"]}
        bound_approvals = [approvals[item] for item in bindings["approval_fingerprints"]]
        if not required:
            if bound_approvals or final["closure_evidence"]:
                raise ArtifactDagError("Reports without approval-required conflicts require explicit empty approvals and closures.")
            continue
        approval_by_conflict: dict[str, dict[str, Any]] = {}
        for approval in bound_approvals:
            if (
                approval["bindings"]["composition_report_fingerprint"]
                != report["semantic_fingerprint"]
            ):
                raise ArtifactDagError(
                    "Final profile binds an approval from another composition report."
                )
            conflict_id = approval["target"]["conflict_id"]
            if conflict_id in approval_by_conflict:
                raise ArtifactDagError("Final profile binds multiple approvals for one conflict.")
            approval_by_conflict[conflict_id] = approval
        closure_by_conflict = {item["conflict_id"]: item for item in final["closure_evidence"]}
        if set(approval_by_conflict) != set(required):
            raise ArtifactDagError("Final profile approvals do not close every and only required conflict.")
        if set(closure_by_conflict) != set(required):
            raise ArtifactDagError("Final profile closure evidence is not bidirectional with required conflicts.")
        for conflict_id, conflict in required.items():
            approval = approval_by_conflict[conflict_id]
            closure = closure_by_conflict[conflict_id]
            if closure["qa_decision_id"] != approval["approval_id"]:
                raise ArtifactDagError("Closure qa_decision_id does not match the bound approval.")
            if closure["proposed_resolution_id"] != conflict["proposed_resolution_id"]:
                raise ArtifactDagError("Closure proposed_resolution_id does not match the report.")

    order = _topological_artifact_order(dependencies)
    return DagValidationResult(order, len(validated), runtime_eligible=False)


def _topological_artifact_order(
    dependencies: Mapping[str, set[str]],
) -> tuple[str, ...]:
    temporary: set[str] = set()
    permanent: set[str] = set()
    order: list[str] = []

    def visit(artifact_id: str) -> None:
        if artifact_id in permanent:
            return
        if artifact_id in temporary:
            raise ArtifactDagError("Artifact bindings contain a cycle.")
        temporary.add(artifact_id)
        for dependency in sorted(dependencies[artifact_id]):
            visit(dependency)
        temporary.remove(artifact_id)
        permanent.add(artifact_id)
        order.append(artifact_id)

    for artifact_id in sorted(dependencies):
        visit(artifact_id)
    return tuple(order)


def validate_artifact_dag(
    documents: list[dict[str, Any]],
) -> DagValidationResult:
    """Validate production artifact bindings without exposing contract overrides."""

    return _validate_artifact_dag_with_reader(
        documents,
        lambda document: validate_artifact(
            document, features={"profile_v2_schema": True}
        ),
        matrix_version="1.0",
    )


def validate_intent_artifact_dag_v041(
    documents: list[dict[str, Any]],
) -> DagValidationResult:
    """Validate the explicit matrix-1.1 declaration-intent DAG without upgrading legacy routes."""

    return _validate_artifact_dag_with_reader(
        documents,
        validate_intent_artifact_v041,
        matrix_version="1.1",
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


def _validate_artifact_dag_for_test(
    documents: list[dict[str, Any]],
    *,
    registry: dict[str, Any],
) -> DagValidationResult:
    """Validate synthetic H DAGs while preserving test-only ineligibility."""

    _require_test_registry(registry)
    return _validate_artifact_dag_with_reader(
        documents,
        lambda document: _validate_artifact_for_test(
            document,
            registry=registry,
            features={"profile_v2_schema": True},
        ),
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
