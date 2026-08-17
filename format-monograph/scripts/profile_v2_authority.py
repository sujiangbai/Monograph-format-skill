#!/usr/bin/env python3
"""Read-only authority precedence contract for V0.4.1 P2b-H."""

from __future__ import annotations

import hashlib
import json
from copy import deepcopy
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator


SCHEMA_DIR = Path(__file__).resolve().parent.parent / "references" / "schemas" / "v2"
AUTHORITY_SCHEMA_PATH = SCHEMA_DIR / "authority-contract.schema.json"
AUTHORITY_CONTRACT_PATHS = {
    "1.0": SCHEMA_DIR / "authority-contract.v1.0.json",
}
AUTHORITY_LAYER_SCHEMA_PATH = SCHEMA_DIR / "authority-layer.generated.schema.json"


class AuthorityContractError(ValueError):
    """Raised when the read-only authority contract is invalid or inconsistent."""


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        raise AuthorityContractError(f"Cannot read authority contract file {path.name}: {exc}") from exc
    if not isinstance(value, dict):
        raise AuthorityContractError(f"Authority contract file must be an object: {path.name}")
    return value


def load_authority_contract(version: str = "1.0") -> dict[str, Any]:
    """Load and validate an immutable authority contract version."""

    path = AUTHORITY_CONTRACT_PATHS.get(version)
    if path is None:
        raise AuthorityContractError(f"Unsupported authority contract version: {version}")
    schema = _load_json(AUTHORITY_SCHEMA_PATH)
    Draft202012Validator.check_schema(schema)
    contract = _load_json(path)
    validate_authority_contract_document(contract, schema=schema)
    return deepcopy(contract)


def validate_authority_contract_document(
    contract: dict[str, Any],
    *,
    schema: dict[str, Any] | None = None,
) -> None:
    """Validate a contract value without creating an alternate authority source."""

    schema = _load_json(AUTHORITY_SCHEMA_PATH) if schema is None else schema
    Draft202012Validator.check_schema(schema)
    errors = sorted(Draft202012Validator(schema).iter_errors(contract), key=lambda item: list(item.path))
    if errors:
        detail = " | ".join(error.message for error in errors)
        raise AuthorityContractError(f"Invalid authority contract: {detail}")
    layer_ids = [item["layer_id"] for item in contract["layers"]]
    ranks = [item["rank"] for item in contract["layers"]]
    if len(layer_ids) != len(set(layer_ids)):
        raise AuthorityContractError("Authority contract repeats a layer_id.")
    if ranks != list(range(len(ranks))):
        raise AuthorityContractError("Authority ranks must be contiguous and ordered from zero.")
    if "safety" in layer_ids:
        raise AuthorityContractError("Safety invariants cannot enter the ordinary authority layers.")


def authority_layer_ids(version: str = "1.0") -> tuple[str, ...]:
    return tuple(item["layer_id"] for item in load_authority_contract(version)["layers"])


def authority_rank(layer_id: str, version: str = "1.0") -> int:
    for item in load_authority_contract(version)["layers"]:
        if item["layer_id"] == layer_id:
            return int(item["rank"])
    raise AuthorityContractError(f"Unknown authority layer: {layer_id}")


def build_authority_layer_schema(version: str = "1.0") -> dict[str, Any]:
    """Generate the layer enum projection from the sole editable authority source."""

    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "$id": f"https://schemas.format-monograph.local/contracts/authority/{version}/layer.schema.json",
        "title": f"Format Monograph Authority Layer {version}",
        "type": "string",
        "enum": list(authority_layer_ids(version)),
    }


def verify_authority_projection(version: str = "1.0") -> None:
    generated = build_authority_layer_schema(version)
    committed = _load_json(AUTHORITY_LAYER_SCHEMA_PATH)
    if generated != committed:
        raise AuthorityContractError(
            "Committed authority layer schema differs from the authority contract projection."
        )


def verify_legacy_layer_compatibility(schema: dict[str, Any], version: str = "1.0") -> None:
    """Check only the legacy enum set; legacy order never becomes authoritative."""

    try:
        values = schema["$defs"]["layer_kind"]["enum"]
    except (KeyError, TypeError) as exc:
        raise AuthorityContractError("Legacy common schema has no layer_kind enum.") from exc
    if set(values) != set(authority_layer_ids(version)):
        raise AuthorityContractError("Legacy layer enum is incompatible with authority contract 1.0.")


def authority_contract_fingerprint(version: str = "1.0") -> str:
    payload = json.dumps(
        load_authority_contract(version),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(payload).hexdigest()
