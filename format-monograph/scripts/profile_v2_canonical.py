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
    "semantic_projected",
    "semantic_guarded",
    "fingerprint_excluded",
    "validation_guard",
}
SCHEMA_DIR = Path(__file__).resolve().parent.parent / "references" / "schemas" / "v2"
COMPOSITION_POLICY_PATH = SCHEMA_DIR / "canonical-composition-policy.v1.0.json"
COMPOSITION_POLICY_SCHEMA_PATH = SCHEMA_DIR / "canonical-composition-policy.schema.json"
LOCAL_SCHEMA_ORIGIN = "https://schemas.format-monograph.local/"


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
            current = current[int(token)] if isinstance(current, list) else current[token]
        except (KeyError, IndexError, TypeError, ValueError) as exc:
            raise CanonicalizationError(f"Unresolved schema pointer: #{fragment}") from exc
    return current


def _schema_registry(documents: Mapping[str, dict[str, Any]]) -> Registry:
    registry = Registry()
    for schema_id, schema in documents.items():
        registry = registry.with_resource(schema_id, Resource.from_contents(schema))
    return registry


def _resolve_local_reference(
    reference_value: str,
    base_uri: str,
    documents: Mapping[str, dict[str, Any]],
    seen_references: set[str] | frozenset[str],
) -> tuple[dict[str, Any], str, str]:
    """Resolve one local reference with the shared fail-closed ref policy."""

    reference = urljoin(base_uri, reference_value)
    if reference in seen_references:
        raise CanonicalizationError(f"Cyclic schema reference: {reference}")
    target_uri, fragment = urldefrag(reference)
    if not target_uri.startswith(LOCAL_SCHEMA_ORIGIN):
        raise CanonicalizationError(f"Remote schema reference is forbidden: {target_uri}")
    target_document = documents.get(target_uri)
    if target_document is None:
        raise CanonicalizationError(f"Offline schema reference is unavailable: {target_uri}")
    target = _pointer(target_document, fragment)
    if not isinstance(target, dict):
        raise CanonicalizationError(f"Schema reference is not an object: {reference}")
    return target, target_uri, reference


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
    seen_references: frozenset[str] = frozenset(),
) -> tuple[dict[str, Any], str]:
    current = deepcopy(schema)
    current_base = current.get("$id", base_uri)
    fingerprint_excluded = current.get("x-semantic-fingerprint") == "exclude"
    if "$ref" in current:
        target, target_uri, reference = _resolve_local_reference(
            current.pop("$ref"), current_base, documents, seen_references
        )
        resolved, resolved_base = _resolve_schema(
            target,
            target_uri,
            documents,
            instance,
            seen_references | {reference},
        )
        fingerprint_excluded = (
            fingerprint_excluded
            or resolved.get("x-semantic-fingerprint") == "exclude"
        )
        current = _merge_schema(resolved, current)
        current_base = resolved_base

    if "oneOf" in current:
        schema_registry = _schema_registry(documents)
        matches: list[dict[str, Any]] = []
        for branch in current["oneOf"]:
            resolved_branch, _ = _resolve_schema(
                branch, current_base, documents, instance, seen_references
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
    if fingerprint_excluded:
        current["x-semantic-fingerprint"] = "exclude"
    return current, current_base


def effective_fingerprint_exclusion(
    schema: dict[str, Any],
    base_uri: str,
    documents: Mapping[str, dict[str, Any]],
) -> bool:
    """Return the monotonic exclusion value across the complete local ref chain."""

    current = schema
    current_base = base_uri
    excluded = False
    seen: set[str] = set()
    while isinstance(current, dict):
        excluded = excluded or current.get("x-semantic-fingerprint") == "exclude"
        reference_value = current.get("$ref")
        if not isinstance(reference_value, str):
            break
        target, target_uri, reference = _resolve_local_reference(
            reference_value, current.get("$id", current_base), documents, seen
        )
        seen.add(reference)
        current = target
        current_base = target_uri
    return excluded


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


def _node_digest(value: Any) -> str:
    return FINGERPRINT_PATTERN_PREFIX + hashlib.sha256(_raw_canonical_bytes(value)).hexdigest()


def load_canonical_composition_policy() -> dict[str, Any]:
    try:
        schema = json.loads(COMPOSITION_POLICY_SCHEMA_PATH.read_text(encoding="utf-8"))
        policy = json.loads(COMPOSITION_POLICY_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        raise CanonicalizationError(f"Cannot read canonical composition policy: {exc}") from exc
    Draft202012Validator.check_schema(schema)
    errors = sorted(Draft202012Validator(schema).iter_errors(policy), key=lambda item: list(item.path))
    if errors:
        raise CanonicalizationError(
            "Invalid canonical composition policy: "
            + " | ".join(error.message for error in errors)
        )
    return policy


def audit_schema_composition(
    documents: Mapping[str, dict[str, Any]],
) -> None:
    """Fail closed for unaudited composition nodes across the full schema tree."""

    policy = load_canonical_composition_policy()
    approved = {
        (item["schema_id"], item["json_pointer"]): item["node_digest"]
        for item in policy["approved_all_of_nodes"]
    }
    observed: dict[tuple[str, str], str] = {}

    def escape(value: str) -> str:
        return value.replace("~", "~0").replace("/", "~1")

    def visit(value: Any, schema_id: str, pointer: str) -> None:
        if isinstance(value, dict):
            if "anyOf" in value:
                raise CanonicalizationError(
                    "Fingerprint evidence does not support schema composition "
                    f"anyOf: {schema_id}{pointer}/anyOf."
                )
            if "allOf" in value:
                node_pointer = f"{pointer}/allOf"
                identity = (schema_id, node_pointer)
                digest = _node_digest(value["allOf"])
                observed[identity] = digest
                if approved.get(identity) != digest:
                    raise CanonicalizationError(
                        f"Canonical policy rejects unaudited allOf at {schema_id}{node_pointer}."
                    )
            for key, child in value.items():
                visit(child, schema_id, f"{pointer}/{escape(str(key))}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, schema_id, f"{pointer}/{index}")

    for schema_id, schema in sorted(documents.items()):
        visit(schema, schema_id, "#")
    missing = set(approved) - set(observed)
    if missing:
        detail = ", ".join(f"{schema_id}{pointer}" for schema_id, pointer in sorted(missing))
        raise CanonicalizationError(f"Canonical policy references missing allOf nodes: {detail}")


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
            if effective_fingerprint_exclusion(
                child_schema, effective_base, documents
            ):
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
    from profile_v2_artifacts import load_routed_contracts, schema_documents

    _, schema, registry, _ = load_routed_contracts(document)
    return schema, schema_documents(), registry


def _schema_for_intent_document_v041(
    document: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, dict[str, Any]], dict[str, Any]]:
    from profile_v2_artifacts import load_routed_contracts, schema_documents

    _, schema, registry, _ = load_routed_contracts(
        document, matrix_version="1.1"
    )
    return schema, schema_documents(matrix_version="1.1"), registry


def _semantic_projection(
    document: dict[str, Any],
    *,
    schema: dict[str, Any],
    documents: Mapping[str, dict[str, Any]],
    registry: dict[str, Any],
    registry_validation_context: str = "strict_execution",
) -> dict[str, Any]:
    validate_registry_document(
        registry, validation_context=registry_validation_context
    )
    audit_schema_composition(documents)
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


def intent_semantic_projection_v041(document: dict[str, Any]) -> dict[str, Any]:
    """Project a P3 declaration/intent artifact using the append-only 1.1 route."""

    schema, documents, registry = _schema_for_intent_document_v041(document)
    return _semantic_projection(
        document,
        schema=schema,
        documents=documents,
        registry=registry,
        registry_validation_context="declaration_intent",
    )


def canonical_intent_semantic_bytes_v041(document: dict[str, Any]) -> bytes:
    return _raw_canonical_bytes(intent_semantic_projection_v041(document))


def compute_intent_semantic_fingerprint_v041(document: dict[str, Any]) -> str:
    return FINGERPRINT_PATTERN_PREFIX + hashlib.sha256(
        canonical_intent_semantic_bytes_v041(document)
    ).hexdigest()


def verify_intent_semantic_fingerprint_v041(document: dict[str, Any]) -> None:
    expected = compute_intent_semantic_fingerprint_v041(document)
    if document.get("semantic_fingerprint") != expected:
        raise CanonicalizationError(
            "semantic_fingerprint does not match canonical intent semantics."
        )


def stamp_intent_semantic_fingerprint_v041(
    document: dict[str, Any],
) -> dict[str, Any]:
    result = deepcopy(document)
    result["semantic_fingerprint"] = compute_intent_semantic_fingerprint_v041(result)
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


def fingerprint_field_node(
    documents: Mapping[str, dict[str, Any]], field_path: str
) -> dict[str, Any]:
    """Resolve an inventory path to the exact committed schema property node."""

    schema_id, fragment = urldefrag(field_path)
    schema = documents.get(schema_id)
    if schema is None:
        raise CanonicalizationError(f"Inventory field references an unknown schema: {schema_id}")
    node = _pointer(schema, fragment)
    if not isinstance(node, dict):
        raise CanonicalizationError(f"Inventory field is not a schema object: {field_path}")
    return node


def schema_node_fingerprint(node: dict[str, Any]) -> str:
    """Fingerprint a raw schema node so the evidence lock detects schema drift."""

    return FINGERPRINT_PATTERN_PREFIX + hashlib.sha256(
        _raw_canonical_bytes(node)
    ).hexdigest()


def _effective_schema_for_inventory(
    node: dict[str, Any],
    base_uri: str,
    documents: Mapping[str, dict[str, Any]],
) -> dict[str, Any]:
    """Resolve only the reference chain needed to describe a field's real type."""

    current = deepcopy(node)
    current_base = base_uri
    seen_references: set[str] = set()
    while "$ref" in current:
        target, target_uri, reference = _resolve_local_reference(
            current.pop("$ref"), current_base, documents, seen_references
        )
        seen_references.add(reference)
        current = _merge_schema(target, current)
        current_base = target_uri
    return current


def _schema_type_feature(node: dict[str, Any]) -> str:
    value_type = node.get("type")
    declared_types = set(value_type) if isinstance(value_type, list) else {value_type}
    if declared_types == {"array"}:
        return "array"
    if declared_types == {"object"} or "properties" in node:
        return "object"
    return "scalar"


def fingerprint_field_inventory(
    documents: Mapping[str, dict[str, Any]],
    *,
    matrix_version: str = "1.0",
) -> list[dict[str, Any]]:
    """Bind every real schema property node to a classification evidence record.

    This is a schema-node classification lock, not a claim that every field has two
    independently valid artifact values. Const and single-value enum fields use
    guarded evidence; executable projection and end-to-end cases are separate.
    """

    from profile_v2_artifacts import schema_inventory_contract

    inventory: list[dict[str, Any]] = []
    audit_schema_composition(documents)

    def escape_pointer(value: str) -> str:
        return value.replace("~", "~0").replace("/", "~1")

    def context_guards_for_schema(
        node: Any, base_uri: str, seen_refs: set[str] | None = None
    ) -> set[str]:
        seen_refs = set() if seen_refs is None else seen_refs
        guards: set[str] = set()
        if isinstance(node, dict):
            if node.get("x-normalized-scope") is True:
                guards.add("normalized_scope_context")
            if node.get("x-property-binding") is True:
                guards.add("property_binding_context")
            reference_value = node.get("$ref")
            if isinstance(reference_value, str):
                reference = urljoin(base_uri, reference_value)
                if reference not in seen_refs:
                    seen_refs.add(reference)
                    target_uri, fragment = urldefrag(reference)
                    target = documents.get(target_uri)
                    if target is None:
                        raise CanonicalizationError(
                            f"Inventory field has an unresolved $ref: {reference}"
                        )
                    guards.update(
                        context_guards_for_schema(
                            _pointer(target, fragment), target_uri, seen_refs
                        )
                    )
            for key, child in node.items():
                if key != "$ref":
                    guards.update(context_guards_for_schema(child, base_uri, seen_refs))
        elif isinstance(node, list):
            for child in node:
                guards.update(context_guards_for_schema(child, base_uri, seen_refs))
        return guards

    def visit(
        value: Any,
        path: str,
        schema_id: str,
        schema_version: str,
        registry_contract_version: str,
        *,
        validation_guard: bool = False,
        inherited_guards: tuple[str, ...] = (),
    ) -> None:
        if isinstance(value, dict):
            if "anyOf" in value:
                raise CanonicalizationError(
                    "Fingerprint evidence does not support schema composition "
                    f"anyOf: {schema_id}#{path or '/'}"
                )
            context_guards = set(inherited_guards)
            if value.get("x-normalized-scope") is True:
                context_guards.add("normalized_scope_context")
            if value.get("x-property-binding") is True:
                context_guards.add("property_binding_context")
            properties = value.get("properties")
            if isinstance(properties, dict) and all(
                isinstance(child, dict) for child in properties.values()
            ):
                for name, child in properties.items():
                    child_path = f"{path}/properties/{escape_pointer(name)}"
                    effective_child = _effective_schema_for_inventory(
                        child, schema_id, documents
                    )
                    child_guards = set(context_guards)
                    child_guards.update(context_guards_for_schema(child, schema_id))
                    if "const" in child:
                        child_guards.add("const")
                    if isinstance(child.get("enum"), list) and len(child["enum"]) == 1:
                        child_guards.add("single_enum")
                    if "$ref" in child:
                        reference = urljoin(schema_id, child["$ref"])
                        target_uri, fragment = urldefrag(reference)
                        target = documents.get(target_uri)
                        if target is None:
                            raise CanonicalizationError(
                                f"Inventory field has an unresolved $ref: {reference}"
                            )
                        target_node = _pointer(target, fragment)
                        if not isinstance(target_node, dict):
                            raise CanonicalizationError(
                                f"Inventory field $ref is not a schema object: {reference}"
                            )
                        if target_node.get("x-normalized-scope") is True:
                            child_guards.add("normalized_scope_context")
                        if target_node.get("x-property-binding") is True or (
                            "property-catalog" in target_uri
                        ):
                            child_guards.add("property_binding_context")

                    if validation_guard:
                        classification = "validation_guard"
                    elif effective_fingerprint_exclusion(
                        child, schema_id, documents
                    ):
                        classification = "fingerprint_excluded"
                    elif child_guards:
                        classification = "semantic_guarded"
                    else:
                        classification = "semantic_projected"
                    if classification not in FIELD_EVIDENCE_MODES:
                        raise CanonicalizationError(
                            f"Unsupported fingerprint evidence classification: {classification}"
                        )

                    features: list[str] = []
                    if "$ref" in child:
                        features.append("ref")
                    if "oneOf" in child or "oneOf" in effective_child:
                        features.append("one_of")
                    if "allOf" in child or "allOf" in effective_child:
                        features.append("all_of_guard")
                    features.append(_schema_type_feature(effective_child))
                    if "generated.schema.json" in schema_id:
                        features.append("generated_schema")
                    features.extend(child_guards)
                    if validation_guard:
                        features.append("validation_guard")

                    field_path = f"{schema_id}#{child_path}"
                    inventory.append(
                        {
                            "field_path": field_path,
                            "container_path": f"{schema_id}#{path}",
                            "field_name": name,
                            "schema_version": schema_version,
                            "registry_contract_version": registry_contract_version,
                            "classification": classification,
                            "guard_reasons": sorted(child_guards),
                            "schema_features": sorted(features),
                            "schema_node_fingerprint": schema_node_fingerprint(child),
                        }
                    )
            for key, child in value.items():
                child_is_guard = validation_guard or key in {
                    "allOf",
                    "if",
                    "then",
                    "else",
                    "not",
                }
                visit(
                    child,
                    f"{path}/{escape_pointer(str(key))}",
                    schema_id,
                    schema_version,
                    registry_contract_version,
                    validation_guard=child_is_guard,
                    inherited_guards=tuple(sorted(context_guards)),
                )
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(
                    child,
                    f"{path}/{index}",
                    schema_id,
                    schema_version,
                    registry_contract_version,
                    validation_guard=validation_guard,
                    inherited_guards=inherited_guards,
                )

    for schema_id, schema in sorted(documents.items()):
        schema_version, registry_contract_version, inventory_enabled = (
            schema_inventory_contract(schema_id, matrix_version=matrix_version)
        )
        if not inventory_enabled:
            continue
        visit(
            schema,
            "",
            schema_id,
            schema_version,
            registry_contract_version,
        )
    return sorted(inventory, key=lambda record: record["field_path"])


def canonical_file_bytes(path: Path) -> bytes:
    document = json.loads(path.read_text(encoding="utf-8"))
    return canonical_semantic_bytes(document)
