#!/usr/bin/env python3
"""Deterministic V2.1 scope normalization and conservative set relations."""

from __future__ import annotations

import hashlib
import json
import unicodedata
from copy import deepcopy
from typing import Any, Literal


SELECTOR_KINDS = {
    "document",
    "section",
    "chapter",
    "semantic_role",
    "object",
    "property",
    "rule",
    "conflict",
}
OVERLAP_STATES = {"overlap", "disjoint", "unknown"}


class ScopeContractError(ValueError):
    """Raised when a scope is ambiguous, contradictory, or incorrectly identified."""


def _nfc(value: str) -> str:
    return unicodedata.normalize("NFC", value)


def _normalize_selector(selector: dict[str, Any]) -> dict[str, Any]:
    kind = selector.get("selector_kind")
    if kind not in SELECTOR_KINDS:
        raise ScopeContractError(f"Unsupported selector kind: {kind}")
    values = selector.get("selector_ids")
    if not isinstance(values, list) or not values:
        raise ScopeContractError(f"Selector {kind} requires at least one stable ID.")
    normalized = [_nfc(value) for value in values if isinstance(value, str)]
    if len(normalized) != len(values) or len(set(normalized)) != len(normalized):
        raise ScopeContractError(f"Selector {kind} contains invalid or duplicate IDs.")
    return {"selector_kind": kind, "selector_ids": sorted(normalized)}


def _normalize_selector_set(
    selectors: list[dict[str, Any]], context: str
) -> list[dict[str, Any]]:
    by_kind: dict[str, dict[str, Any]] = {}
    for selector in selectors:
        normalized = _normalize_selector(selector)
        kind = normalized["selector_kind"]
        if kind in by_kind:
            raise ScopeContractError(f"{context} repeats selector kind {kind}.")
        by_kind[kind] = normalized
    return [by_kind[kind] for kind in sorted(by_kind)]


def _scope_payload(scope: dict[str, Any]) -> dict[str, Any]:
    selectors = _normalize_selector_set(scope.get("selectors", []), "selectors")
    if not selectors:
        raise ScopeContractError("A normalized scope requires a positive selector.")
    exclusions = _normalize_selector_set(scope.get("exclusions", []), "exclusions")
    positive = {item["selector_kind"]: set(item["selector_ids"]) for item in selectors}
    negative = {item["selector_kind"]: set(item["selector_ids"]) for item in exclusions}
    for kind in positive.keys() & negative.keys():
        if positive[kind] <= negative[kind]:
            raise ScopeContractError(
                f"Scope excludes every positive target for selector kind {kind}."
            )

    conditions: dict[str, dict[str, Any]] = {}
    for condition in scope.get("mutually_exclusive_conditions", []):
        condition_id = condition.get("condition_id")
        condition_kind = condition.get("condition_kind")
        if not isinstance(condition_id, str) or not condition_id:
            raise ScopeContractError("Scope condition requires a stable condition_id.")
        condition_id = _nfc(condition_id)
        if condition_id in conditions:
            raise ScopeContractError(f"Duplicate scope condition: {condition_id}")
        if condition_kind not in {
            "requires_selector",
            "excludes_selector",
            "mutually_exclusive",
        }:
            raise ScopeContractError(f"Unsupported scope condition: {condition_kind}")
        conditions[condition_id] = {
            "condition_id": condition_id,
            "condition_kind": condition_kind,
            "target": _normalize_selector(condition.get("target", {})),
        }
    return {
        "selectors": selectors,
        "exclusions": exclusions,
        "mutually_exclusive_conditions": [conditions[key] for key in sorted(conditions)],
    }


def _payload_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")


def derive_scope_id(scope: dict[str, Any]) -> str:
    payload = _scope_payload(scope)
    return "scope:" + hashlib.sha256(_payload_bytes(payload)).hexdigest()


def normalize_scope(scope: dict[str, Any]) -> dict[str, Any]:
    payload = _scope_payload(scope)
    derived = "scope:" + hashlib.sha256(_payload_bytes(payload)).hexdigest()
    supplied = scope.get("scope_id")
    if supplied is not None and supplied != derived:
        raise ScopeContractError("Provided scope_id does not match normalized scope semantics.")
    return {"scope_id": derived, **payload}


def _selector_map(scope: dict[str, Any], field: str) -> dict[str, set[str]]:
    normalized = normalize_scope(scope)
    return {
        item["selector_kind"]: set(item["selector_ids"])
        for item in normalized[field]
    }


def scope_equal(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return normalize_scope(left) == normalize_scope(right)


def scope_overlap_state(
    left: dict[str, Any], right: dict[str, Any]
) -> Literal["overlap", "disjoint", "unknown"]:
    left_normalized = normalize_scope(left)
    right_normalized = normalize_scope(right)
    left_selectors = _selector_map(left_normalized, "selectors")
    right_selectors = _selector_map(right_normalized, "selectors")
    left_exclusions = _selector_map(left_normalized, "exclusions")
    right_exclusions = _selector_map(right_normalized, "exclusions")

    for kind in left_selectors.keys() & right_selectors.keys():
        if left_selectors[kind].isdisjoint(right_selectors[kind]):
            return "disjoint"
    for kind, values in left_selectors.items():
        if values <= right_exclusions.get(kind, set()):
            return "disjoint"
    for kind, values in right_selectors.items():
        if values <= left_exclusions.get(kind, set()):
            return "disjoint"
    if (
        left_normalized["mutually_exclusive_conditions"]
        or right_normalized["mutually_exclusive_conditions"]
    ):
        return "unknown"
    return "overlap"


def scope_disjoint(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return scope_overlap_state(left, right) == "disjoint"


def scope_subset(child: dict[str, Any], parent: dict[str, Any]) -> bool:
    child_normalized = normalize_scope(child)
    parent_normalized = normalize_scope(parent)
    if (
        child_normalized["mutually_exclusive_conditions"]
        or parent_normalized["mutually_exclusive_conditions"]
    ):
        return False
    child_selectors = _selector_map(child_normalized, "selectors")
    parent_selectors = _selector_map(parent_normalized, "selectors")
    child_exclusions = _selector_map(child_normalized, "exclusions")
    parent_exclusions = _selector_map(parent_normalized, "exclusions")
    for kind, parent_values in parent_selectors.items():
        child_values = child_selectors.get(kind)
        if child_values is None or not child_values <= parent_values:
            return False
    for kind, parent_values in parent_exclusions.items():
        if not parent_values <= child_exclusions.get(kind, set()):
            return False
    return True


def normalize_rule_scope(scope: dict[str, Any]) -> dict[str, Any]:
    """Normalize a rule/asset scope expression and derive its stable ID."""

    return normalize_scope({**deepcopy(scope), "scope_id": derive_scope_id(scope)})


def validate_module_asset_scope(asset: dict[str, Any]) -> None:
    if asset.get("layer_kind") != "module":
        return
    boundary = normalize_rule_scope(asset["asset_scope"])
    allowed_kinds = set(asset["allowed_semantic_object_kinds"])
    for rule in asset.get("rules", []):
        if rule["semantic_object_kind"] not in allowed_kinds:
            raise ScopeContractError(
                f"Module rule {rule['rule_id']} exceeds the semantic object boundary."
            )
        rule_scope = normalize_rule_scope(rule["scope"])
        if not scope_subset(rule_scope, boundary):
            raise ScopeContractError(
                f"Module rule {rule['rule_id']} scope is not provably within asset_scope."
            )
