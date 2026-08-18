#!/usr/bin/env python3
"""Deterministic V2.1 scope normalization and conservative set relations."""

from __future__ import annotations

import hashlib
import json
import unicodedata
from copy import deepcopy
from math import prod
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
ALGEBRA_ERROR_CODES = {
    "crossing_overlap",
    "unknown_overlap",
    "unprovable_condition",
    "unexpressible_difference",
    "partition_conservation_failure",
}


class ScopeContractError(ValueError):
    """Raised when a scope is ambiguous, contradictory, or incorrectly identified."""


class ScopeAlgebraError(ScopeContractError):
    """Raised when valid scope inputs cannot be combined with a proven result."""

    def __init__(self, code: str, details: dict[str, Any] | None = None) -> None:
        if code not in ALGEBRA_ERROR_CODES:
            raise ValueError(f"Unsupported scope algebra error code: {code}")
        self.code = code
        self.details = deepcopy(details or {})
        super().__init__(code)


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


def normalized_property_scope_key(
    semantic_object_kind: str,
    property_id: str,
    scope: dict[str, Any],
) -> tuple[str, str, str]:
    """Return the sole semantic identity key for a resolved property or conflict."""

    if not isinstance(semantic_object_kind, str) or not semantic_object_kind:
        raise ScopeContractError("A composition key requires semantic_object_kind.")
    if not isinstance(property_id, str) or not property_id:
        raise ScopeContractError("A composition key requires property_id.")
    if "selectors" in scope:
        scope_id = normalize_scope(scope)["scope_id"]
    else:
        scope_id = scope.get("scope_id")
        if not isinstance(scope_id, str) or not scope_id:
            raise ScopeContractError("A legacy composition key requires scope_id.")
        scope_id = _nfc(scope_id)
    return (_nfc(semantic_object_kind), _nfc(property_id), scope_id)


def _selector_map(scope: dict[str, Any], field: str) -> dict[str, set[str]]:
    normalized = normalize_scope(scope)
    return {
        item["selector_kind"]: set(item["selector_ids"])
        for item in normalized[field]
    }


def _normalized_selector_map(
    normalized: dict[str, Any], field: str
) -> dict[str, frozenset[str]]:
    return {
        item["selector_kind"]: frozenset(item["selector_ids"])
        for item in normalized[field]
    }


def _axis_expression(
    normalized: dict[str, Any], kind: str
) -> tuple[frozenset[str] | None, frozenset[str]]:
    selectors = _normalized_selector_map(normalized, "selectors")
    exclusions = _normalized_selector_map(normalized, "exclusions")
    positive = selectors.get(kind)
    negative = exclusions.get(kind, frozenset())
    if positive is not None:
        positive = positive - negative
        negative = frozenset()
    return positive, negative


def _partial_positive_exclusion_kinds(normalized: dict[str, Any]) -> set[str]:
    selectors = _normalized_selector_map(normalized, "selectors")
    exclusions = _normalized_selector_map(normalized, "exclusions")
    return {
        kind
        for kind in selectors.keys() & exclusions.keys()
        if selectors[kind] & exclusions[kind]
    }


def _axis_subset(
    child: tuple[frozenset[str] | None, frozenset[str]],
    parent: tuple[frozenset[str] | None, frozenset[str]],
) -> bool:
    child_positive, child_exclusions = child
    parent_positive, parent_exclusions = parent
    if child_positive is not None:
        if parent_positive is not None:
            return child_positive <= parent_positive
        return child_positive.isdisjoint(parent_exclusions)
    if parent_positive is not None:
        return False
    return parent_exclusions <= child_exclusions


def _axis_disjoint(
    left: tuple[frozenset[str] | None, frozenset[str]],
    right: tuple[frozenset[str] | None, frozenset[str]],
) -> bool:
    left_positive, left_exclusions = left
    right_positive, right_exclusions = right
    if left_positive is not None and right_positive is not None:
        return left_positive.isdisjoint(right_positive)
    if left_positive is not None:
        return left_positive <= right_exclusions
    if right_positive is not None:
        return right_positive <= left_exclusions
    return False


def _active_kinds(*scopes: dict[str, Any]) -> list[str]:
    kinds: set[str] = set()
    for scope in scopes:
        kinds.update(item["selector_kind"] for item in scope["selectors"])
        kinds.update(item["selector_kind"] for item in scope["exclusions"])
    return sorted(kinds)


def _safe_disjoint_axis(
    left: dict[str, Any], right: dict[str, Any], unsafe_kinds: set[str]
) -> str | None:
    for kind in _active_kinds(left, right):
        if kind in unsafe_kinds:
            continue
        if _axis_disjoint(
            _axis_expression(left, kind), _axis_expression(right, kind)
        ):
            return kind
    return None


def _scope_subset_normalized(
    child: dict[str, Any], parent: dict[str, Any]
) -> bool:
    return all(
        _axis_subset(
            _axis_expression(child, kind), _axis_expression(parent, kind)
        )
        for kind in _active_kinds(child, parent)
    )


def _scope_relation_normalized(
    source: dict[str, Any], cut: dict[str, Any]
) -> Literal[
    "equal",
    "disjoint",
    "strict_subset",
    "strict_superset",
    "crossing",
]:
    if source == cut:
        return "equal"

    unsafe_kinds = _partial_positive_exclusion_kinds(
        source
    ) | _partial_positive_exclusion_kinds(cut)
    if _safe_disjoint_axis(source, cut, unsafe_kinds) is not None:
        return "disjoint"
    if unsafe_kinds:
        raise ScopeAlgebraError(
            "unexpressible_difference",
            {"relation": "unknown", "selector_kinds": sorted(unsafe_kinds)},
        )
    if (
        source["mutually_exclusive_conditions"]
        or cut["mutually_exclusive_conditions"]
    ):
        raise ScopeAlgebraError(
            "unprovable_condition",
            {"relation": "unknown"},
        )

    cut_within_source = _scope_subset_normalized(cut, source)
    source_within_cut = _scope_subset_normalized(source, cut)
    if cut_within_source and not source_within_cut:
        return "strict_subset"
    if source_within_cut and not cut_within_source:
        return "strict_superset"
    if not cut_within_source and not source_within_cut:
        return "crossing"
    raise ScopeAlgebraError("unknown_overlap", {"relation": "unknown"})


def scope_equal(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return normalize_scope(left) == normalize_scope(right)


def scope_overlap_state(
    left: dict[str, Any], right: dict[str, Any]
) -> Literal["overlap", "disjoint", "unknown"]:
    left_normalized = normalize_scope(left)
    right_normalized = normalize_scope(right)
    try:
        relation = _scope_relation_normalized(left_normalized, right_normalized)
    except ScopeAlgebraError:
        return "unknown"
    return "disjoint" if relation == "disjoint" else "overlap"


def scope_disjoint(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return scope_overlap_state(left, right) == "disjoint"


def scope_subset(child: dict[str, Any], parent: dict[str, Any]) -> bool:
    child_normalized = normalize_scope(child)
    parent_normalized = normalize_scope(parent)
    if child_normalized == parent_normalized:
        return True
    if (
        child_normalized["mutually_exclusive_conditions"]
        or parent_normalized["mutually_exclusive_conditions"]
    ):
        return False
    if _partial_positive_exclusion_kinds(
        child_normalized
    ) | _partial_positive_exclusion_kinds(parent_normalized):
        return False
    if _safe_disjoint_axis(child_normalized, parent_normalized, set()) is not None:
        return False
    return _scope_subset_normalized(child_normalized, parent_normalized)


def scope_intersection(
    source: dict[str, Any], cut: dict[str, Any]
) -> dict[str, Any] | None:
    source_normalized = normalize_scope(source)
    cut_normalized = normalize_scope(cut)
    relation = _scope_relation_normalized(source_normalized, cut_normalized)
    if relation == "equal":
        return source_normalized
    if relation == "disjoint":
        return None
    if relation == "strict_subset":
        return cut_normalized
    if relation == "strict_superset":
        return source_normalized
    raise ScopeAlgebraError("crossing_overlap", {"relation": "crossing"})


def _axis_difference(
    source: tuple[frozenset[str] | None, frozenset[str]],
    cut: tuple[frozenset[str] | None, frozenset[str]],
) -> tuple[frozenset[str] | None, frozenset[str]]:
    source_positive, source_exclusions = source
    cut_positive, cut_exclusions = cut
    if source_positive is not None and cut_positive is not None:
        return source_positive - cut_positive, frozenset()
    if source_positive is None and cut_positive is not None:
        return None, source_exclusions | cut_positive
    if source_positive is None and cut_positive is None:
        return cut_exclusions - source_exclusions, frozenset()
    raise ScopeAlgebraError(
        "unexpressible_difference", {"relation": "strict_subset"}
    )


def _scope_from_axes(
    axes: dict[str, tuple[frozenset[str] | None, frozenset[str]]]
) -> dict[str, Any]:
    selectors = []
    exclusions = []
    for kind in sorted(axes):
        positive, negative = axes[kind]
        if positive is not None:
            if not positive:
                raise ScopeAlgebraError(
                    "partition_conservation_failure",
                    {"relation": "strict_subset"},
                )
            selectors.append(
                {"selector_kind": kind, "selector_ids": sorted(positive)}
            )
        if negative:
            exclusions.append(
                {"selector_kind": kind, "selector_ids": sorted(negative)}
            )
    if not selectors:
        raise ScopeAlgebraError(
            "unexpressible_difference", {"relation": "strict_subset"}
        )
    return normalize_scope(
        {
            "selectors": selectors,
            "exclusions": exclusions,
            "mutually_exclusive_conditions": [],
        }
    )


def _construct_strict_subset_residuals(
    source: dict[str, Any], cut: dict[str, Any]
) -> list[dict[str, Any]]:
    kinds = _active_kinds(source, cut)
    source_axes = {kind: _axis_expression(source, kind) for kind in kinds}
    cut_axes = {kind: _axis_expression(cut, kind) for kind in kinds}
    current = dict(source_axes)
    residuals: list[dict[str, Any]] = []
    for kind in kinds:
        if current[kind] == cut_axes[kind]:
            continue
        difference = _axis_difference(current[kind], cut_axes[kind])
        residual_axes = dict(current)
        residual_axes[kind] = difference
        residuals.append(_scope_from_axes(residual_axes))
        current[kind] = cut_axes[kind]
    if current != cut_axes or not residuals:
        raise ScopeAlgebraError(
            "partition_conservation_failure", {"relation": "strict_subset"}
        )
    residuals.sort(key=lambda item: item["scope_id"])
    return residuals


def _proof_axis_masks(
    scope: dict[str, Any],
    kinds: list[str],
    bit_for_id: dict[str, dict[str, int]],
) -> tuple[int, ...]:
    selectors = {
        item["selector_kind"]: set(item["selector_ids"])
        for item in scope["selectors"]
    }
    exclusions = {
        item["selector_kind"]: set(item["selector_ids"])
        for item in scope["exclusions"]
    }
    masks: list[int] = []
    for kind in kinds:
        id_bits = bit_for_id[kind]
        full_mask = (1 << (len(id_bits) + 1)) - 1
        if kind in selectors:
            mask = sum(1 << id_bits[value] for value in selectors[kind])
        else:
            mask = full_mask
        for value in exclusions.get(kind, set()):
            mask &= ~(1 << id_bits[value])
        masks.append(mask)
    return tuple(masks)


def _proof_subset(child: tuple[int, ...], parent: tuple[int, ...]) -> bool:
    return all(child_mask & ~parent_mask == 0 for child_mask, parent_mask in zip(child, parent))


def _proof_disjoint(left: tuple[int, ...], right: tuple[int, ...]) -> bool:
    return any(left_mask & right_mask == 0 for left_mask, right_mask in zip(left, right))


def _proof_size(masks: tuple[int, ...]) -> int:
    return prod(mask.bit_count() for mask in masks)


def _prove_partition_conservation(
    source: dict[str, Any],
    intersection: dict[str, Any],
    residuals: list[dict[str, Any]],
) -> bool:
    scopes = [source, intersection, *residuals]
    kinds = _active_kinds(*scopes)
    ids_by_kind: dict[str, set[str]] = {kind: set() for kind in kinds}
    for scope in scopes:
        for field in ("selectors", "exclusions"):
            for selector in scope[field]:
                ids_by_kind[selector["selector_kind"]].update(selector["selector_ids"])
    bit_for_id = {
        kind: {value: index for index, value in enumerate(sorted(ids_by_kind[kind]))}
        for kind in kinds
    }
    source_masks = _proof_axis_masks(source, kinds, bit_for_id)
    intersection_masks = _proof_axis_masks(intersection, kinds, bit_for_id)
    residual_masks = [
        _proof_axis_masks(residual, kinds, bit_for_id) for residual in residuals
    ]
    if not residual_masks or any(mask == 0 for masks in residual_masks for mask in masks):
        return False
    if len({item["scope_id"] for item in residuals}) != len(residuals):
        return False
    components = [intersection_masks, *residual_masks]
    if any(not _proof_subset(component, source_masks) for component in components):
        return False
    for index, component in enumerate(components):
        for other in components[index + 1 :]:
            if not _proof_disjoint(component, other):
                return False
    return sum(_proof_size(component) for component in components) == _proof_size(
        source_masks
    )


def _strict_subset_difference(
    source: dict[str, Any], cut: dict[str, Any]
) -> list[dict[str, Any]]:
    residuals = _construct_strict_subset_residuals(source, cut)
    if not _prove_partition_conservation(source, cut, residuals):
        raise ScopeAlgebraError(
            "partition_conservation_failure", {"relation": "strict_subset"}
        )
    return residuals


def _raise_for_non_difference_relation(relation: str) -> None:
    if relation == "strict_superset":
        raise ScopeAlgebraError(
            "unexpressible_difference", {"relation": "strict_superset"}
        )
    if relation == "crossing":
        raise ScopeAlgebraError("crossing_overlap", {"relation": "crossing"})
    raise ScopeAlgebraError("unknown_overlap", {"relation": "unknown"})


def scope_difference(
    source: dict[str, Any], cut: dict[str, Any]
) -> list[dict[str, Any]]:
    source_normalized = normalize_scope(source)
    cut_normalized = normalize_scope(cut)
    relation = _scope_relation_normalized(source_normalized, cut_normalized)
    if relation == "equal":
        return []
    if relation == "disjoint":
        return [source_normalized]
    if relation == "strict_subset":
        return _strict_subset_difference(source_normalized, cut_normalized)
    _raise_for_non_difference_relation(relation)
    raise AssertionError("unreachable")


def _partition_evidence(
    relation: str,
    conservation: Literal["proven", "not_proven"],
    source_scope_id: str,
    cut_scope_id: str,
) -> dict[str, Any]:
    return {
        "relation": relation,
        "conservation": conservation,
        "scope_ids": {"source": source_scope_id, "cut": cut_scope_id},
    }


def scope_partition(source: dict[str, Any], cut: dict[str, Any]) -> dict[str, Any]:
    source_normalized = normalize_scope(source)
    cut_normalized = normalize_scope(cut)
    source_scope_id = source_normalized["scope_id"]
    cut_scope_id = cut_normalized["scope_id"]
    try:
        relation = _scope_relation_normalized(source_normalized, cut_normalized)
        if relation == "equal":
            return {
                "status": "equal",
                "code": None,
                "source_scope_id": source_scope_id,
                "cut_scope_id": cut_scope_id,
                "intersection": source_normalized,
                "residual_scopes": [],
                "evidence": _partition_evidence(
                    relation, "proven", source_scope_id, cut_scope_id
                ),
            }
        if relation == "disjoint":
            return {
                "status": "disjoint",
                "code": None,
                "source_scope_id": source_scope_id,
                "cut_scope_id": cut_scope_id,
                "intersection": None,
                "residual_scopes": [source_normalized],
                "evidence": _partition_evidence(
                    relation, "proven", source_scope_id, cut_scope_id
                ),
            }
        if relation == "strict_subset":
            residuals = _strict_subset_difference(source_normalized, cut_normalized)
            return {
                "status": "partitioned",
                "code": None,
                "source_scope_id": source_scope_id,
                "cut_scope_id": cut_scope_id,
                "intersection": cut_normalized,
                "residual_scopes": residuals,
                "evidence": _partition_evidence(
                    relation, "proven", source_scope_id, cut_scope_id
                ),
            }
        _raise_for_non_difference_relation(relation)
    except ScopeAlgebraError as exc:
        relation = exc.details.get("relation", "unknown")
        return {
            "status": "blocked",
            "code": exc.code,
            "source_scope_id": source_scope_id,
            "cut_scope_id": cut_scope_id,
            "intersection": None,
            "residual_scopes": [],
            "evidence": _partition_evidence(
                relation, "not_proven", source_scope_id, cut_scope_id
            ),
        }
    raise AssertionError("unreachable")


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
