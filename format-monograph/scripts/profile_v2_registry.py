#!/usr/bin/env python3
"""Load and mechanically verify the V0.4.1 property registry contract."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker


SCHEMA_DIR = Path(__file__).resolve().parent.parent / "references" / "schemas" / "v2"
REGISTRY_SCHEMA_PATH = SCHEMA_DIR / "property-registry.schema.json"
CORE_REGISTRY_PATH = SCHEMA_DIR / "property-registry.core.json"
GENERATED_CATALOG_PATH = SCHEMA_DIR / "property-catalog.generated.schema.json"
GENERATED_TYPED_VALUE_PATH = SCHEMA_DIR / "typed-value.generated.schema.json"

CATALOG_COLLECTIONS = {
    "data_types": "data_type_id",
    "units": "unit_id",
    "normalizers": "normalizer_id",
    "comparators": "comparator_id",
    "executor_capabilities": "capability_id",
    "auditor_capabilities": "capability_id",
    "constraints": "constraint_id",
    "properties": "property_id",
}

FORBIDDEN_CALLABLE_TOKENS = {
    "__",
    "eval",
    "exec",
    "getattr",
    "globals",
    "import",
    "lambda",
    "locals",
    "module",
    "popen",
    "setattr",
    "system",
}


class RegistryContractError(ValueError):
    """Raised when registry data violates the P1 declaration contract."""


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise RegistryContractError(f"Registry file does not exist: {path}") from exc
    except json.JSONDecodeError as exc:
        raise RegistryContractError(f"Registry file is not valid JSON: {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise RegistryContractError(f"Registry root must be an object: {path}")
    return value


def _format_errors(validator: Draft202012Validator, value: Any) -> list[str]:
    return [
        f"{'/'.join(str(item) for item in error.absolute_path) or '<root>'}: {error.message}"
        for error in sorted(validator.iter_errors(value), key=lambda item: list(item.path))
    ]


def _catalog_ids(registry: dict[str, Any], collection: str) -> list[str]:
    id_key = CATALOG_COLLECTIONS[collection]
    return [str(item[id_key]) for item in registry.get(collection, [])]


def _ensure_unique_catalog_ids(registry: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for collection in CATALOG_COLLECTIONS:
        ids = _catalog_ids(registry, collection)
        if len(ids) != len(set(ids)):
            errors.append(f"{collection} contains duplicate stable IDs.")
    return errors


def _contains_forbidden_callable_token(value: str) -> bool:
    normalized = value.lower().replace("-", ".").replace("_", ".")
    parts = {part for part in normalized.split(".") if part}
    return bool(parts & FORBIDDEN_CALLABLE_TOKENS)


def registry_semantic_errors(registry: dict[str, Any]) -> list[str]:
    """Check cross-catalog references without duplicating property knowledge."""

    errors = _ensure_unique_catalog_ids(registry)
    known_types = set(_catalog_ids(registry, "data_types"))
    known_units = set(_catalog_ids(registry, "units"))
    known_normalizers = set(_catalog_ids(registry, "normalizers"))
    known_comparators = set(_catalog_ids(registry, "comparators"))
    known_executors = set(_catalog_ids(registry, "executor_capabilities"))
    known_auditors = set(_catalog_ids(registry, "auditor_capabilities"))
    known_constraints = set(_catalog_ids(registry, "constraints"))
    known_properties = set(_catalog_ids(registry, "properties"))
    data_type_index = {
        item["data_type_id"]: item for item in registry.get("data_types", [])
    }

    for collection, values, expected_prefix in (
        ("normalizers", known_normalizers, "normalizer."),
        ("comparators", known_comparators, "comparator."),
        ("executor_capabilities", known_executors, "executor."),
        ("auditor_capabilities", known_auditors, "auditor."),
    ):
        for value in values:
            if not value.startswith(expected_prefix) or _contains_forbidden_callable_token(value):
                errors.append(f"{collection} contains a forbidden callable-like ID: {value}.")

    for constraint in registry.get("constraints", []):
        constraint_id = constraint["constraint_id"]
        if _contains_forbidden_callable_token(constraint_id):
            errors.append(f"constraints contains a forbidden callable-like ID: {constraint_id}.")
        missing_properties = sorted(set(constraint["property_ids"]) - known_properties)
        if missing_properties:
            errors.append(
                f"{constraint_id} references unregistered properties: "
                f"{', '.join(missing_properties)}."
            )

    scope = registry.get("registry_scope")
    for item in registry.get("properties", []):
        property_id = item["property_id"]
        if item["data_type_id"] not in known_types:
            errors.append(f"{property_id} references an unregistered data type.")
        if item["normalizer_id"] not in known_normalizers:
            errors.append(f"{property_id} references an unregistered normalizer.")
        if item["comparator_id"] not in known_comparators:
            errors.append(f"{property_id} references an unregistered comparator.")
        if item["executor_capability_id"] not in known_executors:
            errors.append(f"{property_id} references an unregistered executor capability.")
        if item["auditor_capability_id"] not in known_auditors:
            errors.append(f"{property_id} references an unregistered auditor capability.")
        missing_constraints = sorted(set(item.get("constraint_ids", [])) - known_constraints)
        if missing_constraints:
            errors.append(
                f"{property_id} references unregistered constraints: "
                f"{', '.join(missing_constraints)}."
            )

        data_type = data_type_index.get(item["data_type_id"])
        if data_type is not None:
            scalar_schema: dict[str, Any] = {"type": data_type["json_type"]}
            if "pattern" in data_type:
                scalar_schema["pattern"] = data_type["pattern"]
            if "max_length" in data_type:
                scalar_schema["maxLength"] = data_type["max_length"]
            scalar_validator = Draft202012Validator(scalar_schema)
            enum_values = item.get("value_constraints", {}).get("enum_values", [])
            if any(scalar_validator.is_valid(value) is False for value in enum_values):
                errors.append(f"{property_id} has enum values outside its registered data type.")
            numeric_range = item.get("value_constraints", {}).get("numeric_range")
            if numeric_range is not None and item["data_type_id"] not in {"integer", "decimal"}:
                errors.append(f"{property_id} declares a numeric range for a nonnumeric data type.")

        canonical_unit = item.get("canonical_unit_id")
        allowed_units = set(item.get("allowed_unit_ids", []))
        if canonical_unit is not None and canonical_unit not in known_units:
            errors.append(f"{property_id} references an unregistered canonical unit.")
        missing_units = sorted(allowed_units - known_units)
        if missing_units:
            errors.append(
                f"{property_id} references unregistered allowed units: {', '.join(missing_units)}."
            )
        if canonical_unit is not None and canonical_unit not in allowed_units:
            errors.append(f"{property_id} must include its canonical unit in allowed_unit_ids.")
        if canonical_unit is None and allowed_units:
            errors.append(f"{property_id} cannot allow units without a canonical unit.")

        if item.get("safety_invariant"):
            if item.get("overridable") is not False:
                errors.append(f"Safety invariant {property_id} must not be overridable.")
            if item.get("allowed_layers") != ["safety"]:
                errors.append(f"Safety invariant {property_id} must belong only to the safety layer.")
        elif "safety" in item.get("allowed_layers", []):
            errors.append(f"Ordinary property {property_id} cannot enter the safety layer.")

        if scope == "production" and (item.get("test_only") or property_id.startswith("test.")):
            errors.append(f"Production registry cannot contain test property {property_id}.")
        if scope == "test" and (
            item.get("test_only") is not True or not property_id.startswith("test.")
        ):
            errors.append(f"Test registry property must be test-only and use test.*: {property_id}.")
    return errors


def validate_registry_document(registry: dict[str, Any]) -> None:
    schema = _load_json(REGISTRY_SCHEMA_PATH)
    Draft202012Validator.check_schema(schema)
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = _format_errors(validator, registry)
    if not errors:
        errors.extend(registry_semantic_errors(registry))
    if errors:
        raise RegistryContractError("Invalid property registry: " + " | ".join(errors))


def load_registry(path: Path | None = None, *, allow_test: bool = False) -> dict[str, Any]:
    registry = _load_json(path or CORE_REGISTRY_PATH)
    validate_registry_document(registry)
    if registry["registry_scope"] == "test" and not allow_test:
        raise RegistryContractError("Production loader refuses test-only property registries.")
    return registry


def _value_schema(data_type: dict[str, Any]) -> dict[str, Any]:
    value_contract: dict[str, Any] = {"type": data_type["json_type"]}
    if "pattern" in data_type:
        value_contract["pattern"] = data_type["pattern"]
    if "max_length" in data_type:
        value_contract["maxLength"] = data_type["max_length"]
    return {
        "type": "object",
        "additionalProperties": False,
        "required": ["type", "value"],
        "properties": {
            "type": {"const": data_type["data_type_id"]},
            "value": value_contract,
        },
    }


def build_typed_value_schema(registry: dict[str, Any]) -> dict[str, Any]:
    """Generate every typed scalar contract from the registry's data type catalog."""

    validate_registry_document(registry)
    definitions = {
        item["data_type_id"]: _value_schema(item)
        for item in sorted(registry["data_types"], key=lambda value: value["data_type_id"])
    }
    schema = {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "$id": "https://schemas.format-monograph.local/v2/typed-value.generated.schema.json",
        "title": "Generated Format Monograph V2 Typed Values",
        "description": (
            "Generated mechanically from property-registry.core.json; "
            "do not maintain data type contracts here by hand."
        ),
        "$defs": definitions,
        "oneOf": [
            {"$ref": f"#/$defs/{data_type_id}"}
            for data_type_id in sorted(definitions)
        ],
    }
    Draft202012Validator.check_schema(schema)
    return schema


def build_property_catalog_schema(registry: dict[str, Any]) -> dict[str, Any]:
    """Generate the property-binding schema only from registry declarations."""

    validate_registry_document(registry)
    variants: list[dict[str, Any]] = []
    for item in sorted(registry["properties"], key=lambda value: value["property_id"]):
        unit_schema: dict[str, Any]
        if item["canonical_unit_id"] is None:
            unit_schema = {"type": "null"}
        else:
            unit_schema = {
                "type": "string",
                "enum": sorted(item["allowed_unit_ids"]),
            }
        variants.append(
            {
                "type": "object",
                "additionalProperties": False,
                "required": ["property_id", "value", "unit_id", "mode"],
                "properties": {
                    "property_id": {"const": item["property_id"]},
                    "value": {
                        "$ref": (
                            "typed-value.generated.schema.json#/$defs/"
                            f"{item['data_type_id']}"
                        )
                    },
                    "unit_id": unit_schema,
                    "mode": {"type": "string", "enum": sorted(item["modes"])},
                },
            }
        )
    catalog = {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "$id": "https://schemas.format-monograph.local/v2/property-catalog.generated.schema.json",
        "title": "Generated Format Monograph V2 Property Bindings",
        "description": (
            "Generated mechanically from property-registry.core.json; "
            "do not maintain property IDs here by hand."
        ),
        "oneOf": variants,
    }
    Draft202012Validator.check_schema(catalog)
    return catalog


def catalog_property_ids(catalog_schema: dict[str, Any]) -> set[str]:
    result: set[str] = set()
    for variant in catalog_schema.get("oneOf", []):
        property_id = variant.get("properties", {}).get("property_id", {}).get("const")
        if isinstance(property_id, str):
            result.add(property_id)
    return result


def typed_value_type_ids(typed_value_schema: dict[str, Any]) -> set[str]:
    definitions = typed_value_schema.get("$defs", {})
    return set(definitions) if isinstance(definitions, dict) else set()


def catalog_differences(
    registry: dict[str, Any], catalog_schema: dict[str, Any]
) -> dict[str, list[str]]:
    registry_ids = set(_catalog_ids(registry, "properties"))
    schema_ids = catalog_property_ids(catalog_schema)
    return {
        "registry_only": sorted(registry_ids - schema_ids),
        "schema_only": sorted(schema_ids - registry_ids),
    }


def typed_value_differences(
    registry: dict[str, Any], typed_value_schema: dict[str, Any]
) -> dict[str, list[str]]:
    registry_ids = set(_catalog_ids(registry, "data_types"))
    schema_ids = typed_value_type_ids(typed_value_schema)
    return {
        "registry_only": sorted(registry_ids - schema_ids),
        "schema_only": sorted(schema_ids - registry_ids),
    }


def verify_committed_catalog(registry: dict[str, Any] | None = None) -> None:
    registry = registry or load_registry()
    committed = _load_json(GENERATED_CATALOG_PATH)
    generated = build_property_catalog_schema(registry)
    committed_typed_values = _load_json(GENERATED_TYPED_VALUE_PATH)
    generated_typed_values = build_typed_value_schema(registry)
    differences = catalog_differences(registry, committed)
    if differences["registry_only"] or differences["schema_only"]:
        raise RegistryContractError(f"Registry/schema property difference is not zero: {differences}")
    if committed != generated:
        raise RegistryContractError(
            "Generated property catalog differs from the registry-derived contract."
        )
    type_differences = typed_value_differences(registry, committed_typed_values)
    if type_differences["registry_only"] or type_differences["schema_only"]:
        raise RegistryContractError(
            f"Registry/schema data type difference is not zero: {type_differences}"
        )
    if committed_typed_values != generated_typed_values:
        raise RegistryContractError(
            "Generated typed-value schema differs from the registry-derived contract."
        )


def property_index(registry: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["property_id"]: item for item in registry["properties"]}


def validate_binding_for_layer(
    binding: dict[str, Any], layer_kind: str, registry: dict[str, Any]
) -> None:
    entry = property_index(registry).get(binding.get("property_id"))
    if entry is None:
        raise RegistryContractError(f"Unregistered property: {binding.get('property_id')}")
    if layer_kind not in entry["allowed_layers"]:
        raise RegistryContractError(
            f"Property {entry['property_id']} is not allowed in layer {layer_kind}."
        )
    if entry["safety_invariant"] and layer_kind != "safety":
        raise RegistryContractError(
            f"Ordinary rules cannot declare safety invariant {entry['property_id']}."
        )
