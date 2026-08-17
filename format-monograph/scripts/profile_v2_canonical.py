#!/usr/bin/env python3
"""Schema-driven canonical JSON and semantic fingerprints for V2 artifacts."""

from __future__ import annotations

import hashlib
import json
import unicodedata
from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import urldefrag, urljoin

from jsonschema import Draft202012Validator
from referencing import Registry, Resource

from profile_v2_registry import load_registry, validate_registry_document
from profile_v2_scope import normalize_scope
from profile_v2_values import ValueNormalizationError, normalize_property_binding


FINGERPRINT_PATTERN_PREFIX = "sha256:"
ARRAY_SEMANTICS = {"set_by_scalar", "set_by_key", "ordered"}
FIELD_EVIDENCE_MODES = {
    "semantic_direct",
    "semantic_guarded_combination",
    "fingerprint_excluded",
}


class CanonicalizationError(ValueError):
    """Raised when an artifact lacks deterministic canonical semantics."""


def _nfc(value: str) -> str:
    return unicodedata.normalize("NFC", value)


def _pointer(document: dict[str, Any], fragment: str) -> Any:
    if not fragment:
        return document
    if not fragment.startswith("/"):
        raise CanonicalizationError(f"Unsupported schema fragment: #{fragment}")
    current: Any = document
    for raw in fragment[1:].split("/"):
        token = raw.replace("~1", "/").replace("~0", "~")
        try:
            current = current[token]
        except (KeyError, TypeError) as exc:
            raise CanonicalizationError(f"Unresolved schema pointer: #{fragment}") from exc
    return current


def _schema_registry(documents: Mapping[str, dict[str, Any]]) -> Registry:
    registry = Registry()
    for schema_id, schema in documents.items():
        registry = registry.with_resource(schema_id, Resource.from_contents(schema))
    return registry


def _merge_schema(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in overlay.items():
        if key in {"properties", "$defs"} and isinstance(value, dict):
            destination = result.setdefault(key, {})
            for child_key, child_value in value.items():
                if child_key in destination and isinstance(destination[child_key], dict) and isinstance(child_value, dict):
                    destination[child_key] = _merge_schema(destination[child_key], child_value)
                else:
                    destination[child_key] = deepcopy(child_value)
        else:
            result[key] = deepcopy(value)
    return result


def _resolve_schema(
    schema: dict[str, Any],
    base_uri: str,
    documents: Mapping[str, dict[str, Any]],
    instance: Any,
) -> tuple[dict[str, Any], str]:
    current = deepcopy(schema)
    current_base = current.get("$id", base_uri)
    if "$ref" in current:
        reference = urljoin(current_base, current.pop("$ref"))
        target_uri, fragment = urldefrag(reference)
        target_document = documents.get(target_uri)
        if target_document is None:
            raise CanonicalizationError(f"Offline schema reference is unavailable: {target_uri}")
        target = _pointer(target_document, fragment)
        if not isinstance(target, dict):
            raise CanonicalizationError(f"Schema reference is not an object: {reference}")
        resolved, resolved_base = _resolve_schema(target, target_uri, documents, instance)
        current = _merge_schema(resolved, current)
        current_base = resolved_base

    if "oneOf" in current:
        schema_registry = _schema_registry(documents)
        matches: list[dict[str, Any]] = []
        for branch in current["oneOf"]:
            resolved_branch, _ = _resolve_schema(
                branch, current_base, documents, instance
            )
            branch_schema = {
                "$schema": "https://json-schema.org/draft/2020-12/schema",
                "$id": current_base,
                **resolved_branch,
            }
            validator = Draft202012Validator(branch_schema, registry=schema_registry)
            if validator.is_valid(instance):
                matches.append(resolved_branch)
        if len(matches) != 1:
            raise CanonicalizationError(
                f"Canonical schema selection requires exactly one oneOf match; found {len(matches)}."
            )
        remaining = {key: value for key, value in current.items() if key != "oneOf"}
        current = _merge_schema(matches[0], remaining)
    return current, current_base


def _key_value(item: Any, dotted_path: str) -> Any:
    current = item
    for part in dotted_path.split("."):
        if not isinstance(current, dict) or part not in current:
            raise CanonicalizationError(f"set_by_key item lacks stable key {dotted_path}.")
        current = current[part]
    return current


def _raw_canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")


def _project(
    value: Any,
    schema: dict[str, Any],
    base_uri: str,
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
) -> Any:
    effective, effective_base = _resolve_schema(schema, base_uri, documents, value)
    if effective.get("x-property-binding") is True and registry["schema_version"] == "2.1":
        try:
            value = normalize_property_binding(value, registry)
        except ValueNormalizationError as exc:
            raise CanonicalizationError(str(exc)) from exc
    if effective.get("x-normalized-scope") is True:
        value = normalize_scope(value)

    if isinstance(value, dict):
        properties = effective.get("properties", {})
        normalized_keys: dict[str, str] = {}
        for key in value:
            if not isinstance(key, str):
                raise CanonicalizationError("JSON object keys must be strings.")
            normalized = _nfc(key)
            if normalized in normalized_keys and normalized_keys[normalized] != key:
                raise CanonicalizationError("Unicode NFC normalization creates a key collision.")
            normalized_keys[normalized] = key
        result: dict[str, Any] = {}
        for normalized_key in sorted(normalized_keys):
            source_key = normalized_keys[normalized_key]
            child_schema = properties.get(source_key)
            if child_schema is None:
                raise CanonicalizationError(f"Schema does not declare object field {source_key}.")
            if child_schema.get("x-semantic-fingerprint") == "exclude":
                continue
            result[normalized_key] = _project(
                value[source_key], child_schema, effective_base, documents, registry
            )
        return result

    if isinstance(value, list):
        semantics = effective.get("x-semantic-array")
        if semantics not in ARRAY_SEMANTICS:
            raise CanonicalizationError("Array schema lacks a supported x-semantic-array value.")
        item_schema = effective.get("items")
        if not isinstance(item_schema, dict):
            raise CanonicalizationError("Canonical arrays require a single item schema.")
        projected = [
            _project(item, item_schema, effective_base, documents, registry)
            for item in value
        ]
        if semantics == "ordered":
            return projected
        if semantics == "set_by_scalar":
            keys = [_raw_canonical_bytes(item) for item in projected]
        else:
            key_path = effective.get("x-semantic-key")
            if not isinstance(key_path, str) or not key_path:
                raise CanonicalizationError("set_by_key requires x-semantic-key.")
            keys = [_raw_canonical_bytes(_key_value(item, key_path)) for item in projected]
        if len(keys) != len(set(keys)):
            raise CanonicalizationError("Canonical set contains a duplicate stable key.")
        return [item for _, item in sorted(zip(keys, projected), key=lambda pair: pair[0])]

    if isinstance(value, str):
        return _nfc(value)
    if isinstance(value, float):
        raise CanonicalizationError("Binary floating-point values are not canonical JSON inputs.")
    if value is None or isinstance(value, (bool, int)):
        return value
    raise CanonicalizationError(f"Unsupported canonical JSON value type: {type(value).__name__}")


def _schema_for_document(document: dict[str, Any]) -> tuple[dict[str, Any], dict[str, dict[str, Any]], dict[str, Any]]:
    from profile_v2_artifacts import load_artifact_schema, schema_documents

    kind = document.get("artifact_kind")
    version = document.get("schema_version")
    schema = load_artifact_schema(kind, version=version)
    registry = load_registry(version=version)
    return schema, schema_documents(), registry


def _semantic_projection(
    document: dict[str, Any],
    *,
    schema: dict[str, Any],
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
) -> dict[str, Any]:
    validate_registry_document(registry)
    schema_id = schema.get("$id")
    if not isinstance(schema_id, str):
        raise CanonicalizationError("Artifact schema lacks a stable $id.")
    projected = _project(deepcopy(document), schema, schema_id, documents, registry)
    if not isinstance(projected, dict):
        raise CanonicalizationError("Artifact semantic projection must be an object.")
    return projected


def semantic_projection(document: dict[str, Any]) -> dict[str, Any]:
    schema, documents, registry = _schema_for_document(document)
    return _semantic_projection(
        document, schema=schema, documents=documents, registry=registry
    )


def canonical_semantic_bytes(document: dict[str, Any]) -> bytes:
    return _raw_canonical_bytes(semantic_projection(document))


def compute_semantic_fingerprint(document: dict[str, Any]) -> str:
    return FINGERPRINT_PATTERN_PREFIX + hashlib.sha256(
        canonical_semantic_bytes(document)
    ).hexdigest()


def verify_semantic_fingerprint(document: dict[str, Any]) -> None:
    expected = compute_semantic_fingerprint(document)
    if document.get("semantic_fingerprint") != expected:
        raise CanonicalizationError("semantic_fingerprint does not match canonical semantics.")


def stamp_semantic_fingerprint(document: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(document)
    result["semantic_fingerprint"] = compute_semantic_fingerprint(result)
    return result


def _semantic_projection_for_test(
    document: dict[str, Any],
    *,
    schema: dict[str, Any],
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
) -> dict[str, Any]:
    if registry.get("registry_scope") != "test":
        raise CanonicalizationError("Private canonical test path requires registry_scope=test.")
    return _semantic_projection(
        document, schema=schema, documents=documents, registry=registry
    )


def _compute_semantic_fingerprint_for_test(
    document: dict[str, Any],
    *,
    schema: dict[str, Any],
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
) -> str:
    projection = _semantic_projection_for_test(
        document, schema=schema, documents=documents, registry=registry
    )
    return FINGERPRINT_PATTERN_PREFIX + hashlib.sha256(
        _raw_canonical_bytes(projection)
    ).hexdigest()


def _stamp_semantic_fingerprint_for_test(
    document: dict[str, Any],
    *,
    schema: dict[str, Any],
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
) -> dict[str, Any]:
    result = deepcopy(document)
    result["semantic_fingerprint"] = _compute_semantic_fingerprint_for_test(
        result, schema=schema, documents=documents, registry=registry
    )
    return result


def unclassified_array_paths(
    documents: Mapping[str, dict[str, Any]]
) -> list[str]:
    missing: list[str] = []

    def visit(value: Any, path: str) -> None:
        if isinstance(value, dict):
            if value.get("type") == "array" and value.get("x-semantic-array") not in ARRAY_SEMANTICS:
                missing.append(path or "<root>")
            for key, child in value.items():
                visit(child, f"{path}/{key}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, f"{path}/{index}")

    for schema_id, schema in sorted(documents.items()):
        visit(schema, schema_id)
    return missing


def fingerprint_field_inventory(
    documents: Mapping[str, dict[str, Any]],
) -> dict[str, list[str]]:
    """List every schema property node by its machine-tested fingerprint evidence mode."""

    inventory = {mode: [] for mode in sorted(FIELD_EVIDENCE_MODES)}

    def escape_pointer(value: str) -> str:
        return value.replace("~", "~0").replace("/", "~1")

    def visit(value: Any, path: str, schema_id: str) -> None:
        if isinstance(value, dict):
            properties = value.get("properties")
            if isinstance(properties, dict) and all(
                isinstance(child, dict) for child in properties.values()
            ):
                for name, child in properties.items():
                    child_path = f"{path}/properties/{escape_pointer(name)}"
                    if child.get("x-semantic-fingerprint") == "exclude":
                        mode = "fingerprint_excluded"
                    elif "const" in child or (
                        isinstance(child.get("enum"), list)
                        and len(child["enum"]) == 1
                    ):
                        mode = "semantic_guarded_combination"
                    else:
                        mode = "semantic_direct"
                    inventory[mode].append(f"{schema_id}#{child_path}")
            for key, child in value.items():
                visit(child, f"{path}/{escape_pointer(str(key))}", schema_id)
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, f"{path}/{index}", schema_id)

    for schema_id, schema in sorted(documents.items()):
        visit(schema, "", schema_id)
    return {mode: sorted(paths) for mode, paths in sorted(inventory.items())}


def canonical_file_bytes(path: Path) -> bytes:
    document = json.loads(path.read_text(encoding="utf-8"))
    return canonical_semantic_bytes(document)
