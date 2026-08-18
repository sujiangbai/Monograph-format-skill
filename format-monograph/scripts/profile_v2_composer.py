#!/usr/bin/env python3
"""Deterministic, disabled-only Profile V2 composition for V0.4.1 P2b-C.

This module is deliberately not imported by any runtime entry point.  It consumes
the immutable H contracts and the pure S scope algebra, and it emits only a
composition report or a disabled final profile.
"""

from __future__ import annotations

import hashlib
import json
import unicodedata
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Iterable, Mapping, Sequence

from profile_v2_artifacts import (
    ArtifactContractError,
    ProfileV2DisabledError,
    _schema_documents,
    _test_schema_overrides,
    _validate_artifact_for_test,
    load_artifact_schema,
    profile_v2_composer_contract_enabled,
    validate_artifact,
)
from profile_v2_authority import (
    authority_contract_fingerprint,
    authority_rank,
)
from profile_v2_canonical import (
    _stamp_semantic_fingerprint_for_test,
    stamp_semantic_fingerprint,
)
from profile_v2_registry import (
    RegistryContractError,
    load_registry,
    property_index,
    validate_binding_for_layer,
    validate_registry_document,
)
from profile_v2_scope import (
    ScopeContractError,
    normalize_rule_scope,
    normalize_scope,
    normalized_property_scope_key,
    scope_disjoint,
    scope_partition,
    scope_subset,
)
from profile_v2_values import (
    ValueNormalizationError,
    compare_property_bindings,
    normalize_property_binding,
)


REPORT_STATUS_PRIORITY = ("fatal", "unresolvable", "awaiting_approval", "resolvable")
APPLICATION_STATUSES = {"fatal", "unresolvable", "awaiting_approval", "profile_generated"}
CONFIDENCE_ORDER = {"high": 0, "medium": 1, "low": 2}


class ComposerContractError(ValueError):
    """Raised when caller-supplied composition metadata is malformed."""


class ComposerDisabledError(ProfileV2DisabledError):
    """Raised before composition when the H feature contract is not enabled."""


@dataclass(frozen=True)
class _ContractAdapter:
    registry: dict[str, Any]
    validate: Callable[[dict[str, Any]], None]
    stamp: Callable[[dict[str, Any]], dict[str, Any]]
    feature_enabled: Callable[[dict[str, Any] | None], bool]


@dataclass(frozen=True)
class _Candidate:
    record: dict[str, Any]
    semantic_object_kind: str
    property_id: str
    scope: dict[str, Any]


@dataclass(frozen=True)
class _Region:
    scope: dict[str, Any]
    candidates: tuple[_Candidate, ...]


def _canonical_value(value: Any) -> Any:
    if isinstance(value, dict):
        normalized: dict[str, Any] = {}
        for key, item in value.items():
            normalized_key = unicodedata.normalize("NFC", key)
            if normalized_key in normalized:
                raise ComposerContractError("NFC normalization creates a duplicate object key.")
            normalized[normalized_key] = _canonical_value(item)
        return {key: normalized[key] for key in sorted(normalized)}
    if isinstance(value, (list, tuple)):
        return [_canonical_value(item) for item in value]
    if isinstance(value, str):
        return unicodedata.normalize("NFC", value)
    if isinstance(value, float):
        raise ComposerContractError("Binary floating-point values are not deterministic inputs.")
    if value is None or isinstance(value, (bool, int)):
        return value
    raise ComposerContractError(f"Unsupported deterministic input type: {type(value).__name__}")


def _canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        _canonical_value(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")


def _digest(value: Any) -> str:
    return "sha256:" + hashlib.sha256(_canonical_bytes(value)).hexdigest()


def _stable_id(prefix: str, value: Any) -> str:
    return f"{prefix}:{hashlib.sha256(_canonical_bytes(value)).hexdigest()}"


def _input(input_id: str, role: str, fingerprint: str) -> dict[str, str]:
    return {"input_id": input_id, "role": role, "fingerprint": fingerprint}


def _production_adapter() -> _ContractAdapter:
    registry = load_registry(version="2.1")

    def validate(document: dict[str, Any]) -> None:
        validate_artifact(document, features={"profile_v2_schema": True})

    return _ContractAdapter(
        registry=registry,
        validate=validate,
        stamp=stamp_semantic_fingerprint,
        feature_enabled=profile_v2_composer_contract_enabled,
    )


def _test_adapter(registry: dict[str, Any]) -> _ContractAdapter:
    validate_registry_document(registry)
    if registry.get("registry_scope") != "test":
        raise ComposerContractError("Private composer test path requires registry_scope=test.")
    documents = _schema_documents(_test_schema_overrides(registry))

    def validate(document: dict[str, Any]) -> None:
        _validate_artifact_for_test(
            document,
            registry=registry,
            features={"profile_v2_schema": True},
        )

    def stamp(document: dict[str, Any]) -> dict[str, Any]:
        schema = load_artifact_schema(
            document["artifact_kind"], version=document["schema_version"]
        )
        return _stamp_semantic_fingerprint_for_test(
            document,
            schema=schema,
            documents=documents,
            registry=registry,
        )

    def feature_enabled(manifest: dict[str, Any] | None) -> bool:
        if manifest is None:
            return False
        try:
            validate(manifest)
        except (ArtifactContractError, ValueError):
            return False
        features = manifest.get("features", {})
        return bool(
            manifest.get("artifact_kind") == "feature-activation-manifest"
            and manifest.get("schema_version") == "2.1"
            and features.get("profile_v2_schema") is True
            and features.get("profile_v2_composer") is True
        )

    return _ContractAdapter(registry, validate, stamp, feature_enabled)


def _validate_fingerprint(value: str, name: str) -> None:
    if not isinstance(value, str) or len(value) != 71 or not value.startswith("sha256:"):
        raise ComposerContractError(f"{name} must be a sha256 semantic fingerprint.")
    try:
        int(value[7:], 16)
    except ValueError as exc:
        raise ComposerContractError(f"{name} must be a sha256 semantic fingerprint.") from exc


def _report_status(report: Mapping[str, Any]) -> str:
    if report.get("fatal_diagnostics"):
        return "fatal"
    if report.get("unresolvable_blockers"):
        return "unresolvable"
    if report.get("approval_required_conflicts"):
        return "awaiting_approval"
    return "resolvable"


def _reason_code(prefix: str, value: Any) -> str:
    return f"{prefix}-{hashlib.sha256(_canonical_bytes(value)).hexdigest()[:16].upper()}"


def _fatal(category: str, value: Any, artifact_id: str | None = None) -> dict[str, Any]:
    result: dict[str, Any] = {
        "diagnostic_id": _stable_id("fatal", [category, value]),
        "category": category,
        "reason_code": _reason_code("COMPOSER-FATAL", [category, value]),
    }
    if isinstance(artifact_id, str) and ":" in artifact_id:
        result["related_artifact_id"] = artifact_id
    return result


def _diagnostic(category: str, value: Any) -> dict[str, Any]:
    return {
        "diagnostic_id": _stable_id("diagnostic", [category, value]),
        "category": category,
        "reason_code": _reason_code("COMPOSER-DIAGNOSTIC", [category, value]),
    }


def _application_diagnostic(category: str, value: Any, conflict_id: str | None = None) -> dict[str, Any]:
    result: dict[str, Any] = {
        "diagnostic_id": _stable_id("application-diagnostic", [category, value, conflict_id]),
        "category": category,
        "reason_code": _reason_code("APPLICATION", [category, value, conflict_id]),
    }
    if conflict_id is not None:
        result["conflict_id"] = conflict_id
    return result


def _safe_artifact_fingerprint(asset: Mapping[str, Any]) -> str:
    value = asset.get("semantic_fingerprint")
    try:
        _validate_fingerprint(value, "rule asset fingerprint")
    except ComposerContractError:
        return _digest({"invalid_rule_asset": asset})
    return str(value)


def _candidate_identity(
    asset: Mapping[str, Any],
    rule: Mapping[str, Any],
    binding: Mapping[str, Any],
    scope: Mapping[str, Any],
) -> dict[str, Any]:
    return {
        "source_artifact_id": asset["artifact_id"],
        "source_rule_id": rule["rule_id"],
        "semantic_object_kind": rule["semantic_object_kind"],
        "layer_kind": asset["layer_kind"],
        "property_binding": binding,
        "scope_id": scope["scope_id"],
    }


def _candidate_sort_key(candidate: _Candidate) -> tuple[int, str]:
    return (authority_rank(candidate.record["layer_kind"]), candidate.record["candidate_id"])


def _region_sort_key(region: _Region) -> str:
    return region.scope["scope_id"]


def _insert_candidate(
    regions: Sequence[_Region], scope: dict[str, Any], candidate: _Candidate
) -> tuple[list[_Region], dict[str, Any] | None]:
    ordered = sorted(regions, key=_region_sort_key)
    for index, region in enumerate(ordered):
        partition = scope_partition(region.scope, scope)
        if partition["status"] == "disjoint":
            continue
        if partition["status"] == "equal":
            replacement = _Region(
                region.scope,
                tuple(sorted((*region.candidates, candidate), key=_candidate_sort_key)),
            )
            return sorted([*ordered[:index], replacement, *ordered[index + 1 :]], key=_region_sort_key), None
        if partition["status"] == "partitioned":
            residuals = [_Region(item, region.candidates) for item in partition["residual_scopes"]]
            intersection = _Region(
                partition["intersection"],
                tuple(sorted((*region.candidates, candidate), key=_candidate_sort_key)),
            )
            result = [*ordered[:index], *residuals, intersection, *ordered[index + 1 :]]
            return sorted(result, key=_region_sort_key), None
        if partition["evidence"]["relation"] == "strict_superset":
            reverse = scope_partition(scope, region.scope)
            if reverse["status"] != "partitioned":
                return ordered, partition
            result = [*ordered[:index], *ordered[index + 1 :]]
            updated = _Region(
                region.scope,
                tuple(sorted((*region.candidates, candidate), key=_candidate_sort_key)),
            )
            result.append(updated)
            for residual in reverse["residual_scopes"]:
                result, blocked = _insert_candidate(result, residual, candidate)
                if blocked is not None:
                    return ordered, blocked
            return sorted(result, key=_region_sort_key), None
        return ordered, partition
    return sorted([*ordered, _Region(scope, (candidate,))], key=_region_sort_key), None


def _partition_candidates(
    candidates: Sequence[_Candidate],
) -> tuple[list[_Region], dict[str, Any] | None]:
    regions: list[_Region] = []
    for candidate in sorted(candidates, key=_candidate_sort_key):
        regions, blocked = _insert_candidate(regions, candidate.scope, candidate)
        if blocked is not None:
            return [], blocked
    for left_index, left in enumerate(regions):
        for right in regions[left_index + 1 :]:
            if not scope_disjoint(left.scope, right.scope):
                return [], {
                    "status": "blocked",
                    "code": "partition_conservation_failure",
                    "evidence": {"relation": "unknown"},
                }
    return regions, None


def _scope_components(candidates: Sequence[_Candidate]) -> list[list[_Candidate]]:
    """Group candidates only by proven scope connectivity, not by input order."""

    ordered = sorted(candidates, key=_candidate_sort_key)
    remaining = set(range(len(ordered)))
    components: list[list[_Candidate]] = []
    while remaining:
        seed = min(remaining)
        pending = [seed]
        component: set[int] = set()
        while pending:
            current = pending.pop()
            if current in component:
                continue
            component.add(current)
            remaining.discard(current)
            for other in sorted(remaining):
                relation = scope_partition(
                    ordered[current].scope, ordered[other].scope
                )
                if relation["status"] != "disjoint":
                    pending.append(other)
        components.append([ordered[index] for index in sorted(component)])
    return sorted(
        components,
        key=lambda items: min(item.scope["scope_id"] for item in items),
    )


def _binding_groups(
    candidates: Sequence[_Candidate], registry: dict[str, Any]
) -> list[list[_Candidate]]:
    groups: list[list[_Candidate]] = []
    for candidate in sorted(candidates, key=_candidate_sort_key):
        for group in groups:
            if compare_property_bindings(
                candidate.record["property_binding"],
                group[0].record["property_binding"],
                registry,
            ):
                group.append(candidate)
                break
        else:
            groups.append([candidate])
    return sorted(groups, key=lambda group: group[0].record["candidate_id"])


def _confidence(candidates: Iterable[_Candidate]) -> str:
    return max((item.record["confidence"] for item in candidates), key=CONFIDENCE_ORDER.__getitem__)


def _override_chain(candidates: Sequence[_Candidate]) -> list[str]:
    by_layer: dict[str, list[_Candidate]] = {}
    for candidate in candidates:
        by_layer.setdefault(candidate.record["layer_kind"], []).append(candidate)
    result: list[str] = []
    for layer in sorted(by_layer, key=authority_rank):
        result.append(min(item.record["candidate_id"] for item in by_layer[layer]))
    return result


def _key(candidate: _Candidate, scope: dict[str, Any]) -> dict[str, Any]:
    return {
        "semantic_object_kind": candidate.semantic_object_kind,
        "property_id": candidate.property_id,
        "normalized_scope": deepcopy(scope),
    }


def _make_candidate_group(key: dict[str, Any], candidates: Sequence[_Candidate]) -> dict[str, Any]:
    identity = normalized_property_scope_key(
        key["semantic_object_kind"], key["property_id"], key["normalized_scope"]
    )
    return {
        "candidate_group_id": _stable_id("candidate-group", identity),
        "key": deepcopy(key),
        "candidates": [deepcopy(item.record) for item in sorted(candidates, key=_candidate_sort_key)],
        "excluded_candidates": [],
    }


def _make_proposal(
    key: dict[str, Any], candidates: Sequence[_Candidate], registry: dict[str, Any]
) -> tuple[dict[str, Any], list[list[_Candidate]], list[_Candidate]]:
    winning_rank = max(authority_rank(item.record["layer_kind"]) for item in candidates)
    winning = [
        item for item in candidates if authority_rank(item.record["layer_kind"]) == winning_rank
    ]
    value_groups = _binding_groups(winning, registry)
    chosen_group = min(value_groups, key=lambda group: group[0].record["candidate_id"])
    final_candidate = min(chosen_group, key=lambda item: item.record["candidate_id"])
    identity = normalized_property_scope_key(
        key["semantic_object_kind"], key["property_id"], key["normalized_scope"]
    )
    proposal = {
        "proposed_resolution_id": _stable_id("proposed-resolution", identity),
        "key": deepcopy(key),
        "proposed_binding": deepcopy(final_candidate.record["property_binding"]),
        "final_layer_kind": final_candidate.record["layer_kind"],
        "final_source": deepcopy(final_candidate.record["source"]),
        "candidate_chain": [deepcopy(item.record) for item in sorted(candidates, key=_candidate_sort_key)],
        "override_chain": _override_chain(candidates),
        "confidence": _confidence(chosen_group),
        "execution_mode": final_candidate.record["property_binding"]["mode"],
    }
    return proposal, value_groups, chosen_group


def _report_documents(
    rule_assets: Sequence[dict[str, Any]],
    feature_manifest: dict[str, Any],
    *,
    adapter: _ContractAdapter,
    input_fingerprint: str,
    structure_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
    generated_at: str,
) -> dict[str, Any]:
    if not rule_assets:
        raise ComposerContractError("Composition requires at least one rule asset.")
    _validate_fingerprint(input_fingerprint, "input_fingerprint")
    _validate_fingerprint(structure_fingerprint, "structure_fingerprint")
    if not adapter.feature_enabled(feature_manifest):
        raise ComposerDisabledError(
            "profile_v2_schema and profile_v2_composer must both be explicitly true."
        )

    registry = adapter.registry
    validate_registry_document(registry)
    registry_fingerprint = _digest(registry)
    authority_fingerprint = authority_contract_fingerprint("1.0")
    fatal_diagnostics: list[dict[str, Any]] = []
    diagnostics: list[dict[str, Any]] = []
    candidates_by_property: dict[tuple[str, str], list[_Candidate]] = {}
    seen_artifact_ids: set[str] = set()
    seen_fingerprints: set[str] = set()
    bound_fingerprints: list[str] = []

    for asset in sorted((deepcopy(item) for item in rule_assets), key=lambda item: _canonical_bytes(item)):
        asset_id = asset.get("artifact_id")
        fingerprint = _safe_artifact_fingerprint(asset)
        if asset_id in seen_artifact_ids or fingerprint in seen_fingerprints:
            fatal_diagnostics.append(_fatal("invalid_artifact", ["duplicate", asset_id, fingerprint], asset_id))
            continue
        if isinstance(asset_id, str):
            seen_artifact_ids.add(asset_id)
        seen_fingerprints.add(fingerprint)
        bound_fingerprints.append(fingerprint)
        try:
            adapter.validate(asset)
        except (ArtifactContractError, RegistryContractError, ScopeContractError, ValueError) as exc:
            text = str(exc)
            category = (
                "module_asset_out_of_bounds"
                if asset.get("layer_kind") == "module" and ("Module rule" in text or "asset_scope" in text)
                else "safety_hard_failure"
                if "safety invariant" in text.lower() or asset.get("can_override_safety_invariants") is not False
                else "invalid_fingerprint"
                if "fingerprint" in text.lower()
                else "invalid_binding"
                if "binding" in text.lower() or "property" in text.lower()
                else "invalid_artifact"
            )
            fatal_diagnostics.append(_fatal(category, [asset_id, text], asset_id))
            continue
        if asset.get("activation") != "approved":
            diagnostics.append(_diagnostic("preflight", [asset_id, "not-approved"]))
            continue
        for rule in sorted(asset.get("rules", []), key=lambda item: item["rule_id"]):
            if rule.get("status") != "approved":
                diagnostics.append(_diagnostic("preflight", [asset_id, rule.get("rule_id"), "not-approved"]))
                continue
            try:
                scope = normalize_rule_scope(rule["scope"])
            except ScopeContractError as exc:
                fatal_diagnostics.append(_fatal("invalid_binding", [asset_id, rule["rule_id"], str(exc)], asset_id))
                continue
            for raw_binding in sorted(rule["properties"], key=lambda item: item["property_id"]):
                try:
                    validate_binding_for_layer(raw_binding, asset["layer_kind"], registry)
                    binding = normalize_property_binding(raw_binding, registry)
                except (RegistryContractError, ValueNormalizationError) as exc:
                    fatal_diagnostics.append(_fatal("invalid_binding", [asset_id, rule["rule_id"], str(exc)], asset_id))
                    continue
                entry = property_index(registry)[binding["property_id"]]
                if entry.get("safety_invariant") or asset.get("can_override_safety_invariants") is not False:
                    fatal_diagnostics.append(_fatal("safety_hard_failure", [asset_id, rule["rule_id"], binding["property_id"]], asset_id))
                    continue
                identity = _candidate_identity(asset, rule, binding, scope)
                record = {
                    "candidate_id": _stable_id("candidate", identity),
                    "property_binding": binding,
                    "source": {
                        "source_artifact_id": asset["artifact_id"],
                        "source_rule_id": rule["rule_id"],
                    },
                    "layer_kind": asset["layer_kind"],
                    "confidence": rule["confidence"],
                    "scope_status": "applicable",
                }
                candidate = _Candidate(
                    record,
                    rule["semantic_object_kind"],
                    binding["property_id"],
                    scope,
                )
                candidates_by_property.setdefault(
                    (candidate.semantic_object_kind, candidate.property_id), []
                ).append(candidate)

    candidate_groups: list[dict[str, Any]] = []
    scope_partitions: list[dict[str, Any]] = []
    proposed_resolutions: list[dict[str, Any]] = []
    blockers: list[dict[str, Any]] = []
    approval_conflicts: list[dict[str, Any]] = []

    for group_identity in sorted(candidates_by_property):
        candidates = candidates_by_property[group_identity]
        candidate_ids = [item.record["candidate_id"] for item in candidates]
        if len(candidate_ids) != len(set(candidate_ids)):
            fatal_diagnostics.append(_fatal("invalid_artifact", ["candidate-collision", group_identity]))
            continue
        regions: list[_Region] = []
        successful_candidates: list[_Candidate] = []
        for component in _scope_components(candidates):
            component_regions, blocked = _partition_candidates(component)
            if blocked is None:
                regions.extend(component_regions)
                successful_candidates.extend(component)
                continue
            first = min(component, key=_candidate_sort_key)
            key = _key(first, first.scope)
            code = blocked.get("code") or "unknown_overlap"
            blocker = {
                "blocker_id": _stable_id("blocker", [group_identity, key, code]),
                "category": code,
                "key": key,
                "reason_code": _reason_code("SCOPE-BLOCKED", [group_identity, code]),
            }
            blockers.append(blocker)
            scope_partitions.append(
                {
                    "partition_id": _stable_id("partition", [group_identity, key, "failed"]),
                    "key": deepcopy(key),
                    "source_scope": deepcopy(first.scope),
                    "partition_scopes": [deepcopy(first.scope)],
                    "evidence_status": "failed",
                }
            )

        source_scopes: dict[str, dict[str, Any]] = {
            item.scope["scope_id"]: item.scope for item in successful_candidates
        }
        for source_scope in sorted(source_scopes.values(), key=lambda item: item["scope_id"]):
            first = min(candidates, key=_candidate_sort_key)
            partition_scopes = [
                deepcopy(region.scope)
                for region in regions
                if scope_subset(region.scope, source_scope)
            ]
            key = _key(first, source_scope)
            scope_partitions.append(
                {
                    "partition_id": _stable_id("partition", [group_identity, source_scope["scope_id"]]),
                    "key": key,
                    "source_scope": deepcopy(source_scope),
                    "partition_scopes": sorted(partition_scopes, key=lambda item: item["scope_id"]),
                    "evidence_status": "verified",
                }
            )

        for region in regions:
            first = min(region.candidates, key=_candidate_sort_key)
            key = _key(first, region.scope)
            candidate_groups.append(_make_candidate_group(key, region.candidates))
            proposal, winning_value_groups, chosen_group = _make_proposal(
                key, region.candidates, registry
            )
            proposed_resolutions.append(proposal)
            needs_approval = len(winning_value_groups) > 1 or proposal["confidence"] in {"medium", "low"}
            if needs_approval:
                identity = normalized_property_scope_key(
                    key["semantic_object_kind"], key["property_id"], key["normalized_scope"]
                )
                approval_conflicts.append(
                    {
                        "conflict_id": _stable_id("conflict", identity),
                        "proposed_resolution_id": proposal["proposed_resolution_id"],
                        "key": deepcopy(key),
                        "candidates": [
                            deepcopy(item.record)
                            for item in sorted(region.candidates, key=_candidate_sort_key)
                        ],
                        "allowed_decisions": [
                            "adopt_proposed",
                            "exclude_candidate",
                            "keep_original",
                            "select_candidate",
                        ],
                    }
                )

    bound_fingerprints = sorted(set(bound_fingerprints))
    if not bound_fingerprints:
        bound_fingerprints = [_digest({"empty_rule_asset_set": rule_assets})]
        fatal_diagnostics.append(_fatal("invalid_artifact", "no-bindable-rule-assets"))
    bindings = {
        "input_fingerprint": input_fingerprint,
        "feature_activation_fingerprint": feature_manifest["semantic_fingerprint"],
        "property_registry_fingerprint": registry_fingerprint,
        "authority_contract_fingerprint": authority_fingerprint,
        "rule_asset_fingerprints": bound_fingerprints,
        "structure_fingerprint": structure_fingerprint,
    }
    inputs = [
        _input("input:source", "source_document", input_fingerprint),
        _input("input:feature", "feature_activation", feature_manifest["semantic_fingerprint"]),
        _input("input:registry", "property_registry", registry_fingerprint),
        _input("input:authority", "authority_contract", authority_fingerprint),
        _input("input:structure", "structure", structure_fingerprint),
    ]
    inputs.extend(
        _input(_stable_id("input-rule", fingerprint), "rule_asset", fingerprint)
        for fingerprint in bound_fingerprints
    )
    report = {
        "artifact_kind": "conflict-report",
        "schema_version": "2.2",
        "registry_contract_version": "2.1",
        "authority_contract_version": "1.0",
        "artifact_id": artifact_id,
        "created_by_tool": deepcopy(created_by_tool),
        "input_fingerprints": sorted(inputs, key=lambda item: item["input_id"]),
        "semantic_fingerprint": _digest("unstamped-composition-report"),
        "generated_at": generated_at,
        "bindings": bindings,
        "candidate_groups": sorted(candidate_groups, key=lambda item: item["candidate_group_id"]),
        "scope_partitions": sorted(scope_partitions, key=lambda item: item["partition_id"]),
        "proposed_resolutions": sorted(proposed_resolutions, key=lambda item: item["proposed_resolution_id"]),
        "fatal_diagnostics": sorted(fatal_diagnostics, key=lambda item: item["diagnostic_id"]),
        "unresolvable_blockers": sorted(blockers, key=lambda item: item["blocker_id"]),
        "approval_required_conflicts": sorted(approval_conflicts, key=lambda item: item["conflict_id"]),
        "diagnostics": sorted(diagnostics, key=lambda item: item["diagnostic_id"]),
        "proposal_status": "resolvable",
    }
    report["proposal_status"] = _report_status(report)
    stamped = adapter.stamp(report)
    adapter.validate(stamped)
    return stamped


def compose_profile(
    rule_assets: Sequence[dict[str, Any]],
    feature_manifest: dict[str, Any],
    *,
    input_fingerprint: str,
    structure_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
    generated_at: str,
) -> dict[str, Any]:
    """Compose a production-contract report without enabling any runtime path."""

    return _report_documents(
        rule_assets,
        feature_manifest,
        adapter=_production_adapter(),
        input_fingerprint=input_fingerprint,
        structure_fingerprint=structure_fingerprint,
        artifact_id=artifact_id,
        created_by_tool=created_by_tool,
        generated_at=generated_at,
    )


def _compose_profile_for_test(
    rule_assets: Sequence[dict[str, Any]],
    feature_manifest: dict[str, Any],
    *,
    registry: dict[str, Any],
    input_fingerprint: str,
    structure_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
    generated_at: str,
) -> dict[str, Any]:
    return _report_documents(
        rule_assets,
        feature_manifest,
        adapter=_test_adapter(registry),
        input_fingerprint=input_fingerprint,
        structure_fingerprint=structure_fingerprint,
        artifact_id=artifact_id,
        created_by_tool=created_by_tool,
        generated_at=generated_at,
    )


def _proposal_to_resolved(
    proposal: Mapping[str, Any], registry: dict[str, Any]
) -> dict[str, Any]:
    safety_ids = sorted(
        item["property_id"] for item in registry["properties"] if item.get("safety_invariant")
    )
    if not safety_ids:
        raise ComposerContractError("The registry has no declared safety invariant.")
    return {
        "resolution_id": proposal["proposed_resolution_id"],
        "key": deepcopy(proposal["key"]),
        "resolved_binding": deepcopy(proposal["proposed_binding"]),
        "final_layer_kind": proposal["final_layer_kind"],
        "final_source": deepcopy(proposal["final_source"]),
        "candidate_chain": deepcopy(proposal["candidate_chain"]),
        "override_chain": list(proposal["override_chain"]),
        "excluded_candidates": [],
        "confidence": proposal["confidence"],
        "safety_check": {"status": "pass", "checked_invariant_ids": safety_ids},
        "execution_mode": proposal["execution_mode"],
    }


def _candidate_matches_binding(
    candidate: Mapping[str, Any], binding: Mapping[str, Any], registry: dict[str, Any]
) -> bool:
    try:
        return compare_property_bindings(
            candidate["property_binding"], dict(binding), registry
        ) and candidate["property_binding"]["mode"] == binding["mode"]
    except (KeyError, ValueNormalizationError):
        return False


def _resolved_from_candidate(
    proposal: Mapping[str, Any],
    candidate: Mapping[str, Any],
    registry: dict[str, Any],
    *,
    chain: Sequence[dict[str, Any]] | None = None,
    excluded: Sequence[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    active = list(deepcopy(chain if chain is not None else proposal["candidate_chain"]))
    same_group = [
        item
        for item in active
        if item["layer_kind"] == candidate["layer_kind"]
        and _candidate_matches_binding(item, candidate["property_binding"], registry)
    ]
    confidence = max(
        (item["confidence"] for item in same_group),
        key=CONFIDENCE_ORDER.__getitem__,
    )
    wrappers = [
        _Candidate(item, proposal["key"]["semantic_object_kind"], proposal["key"]["property_id"], proposal["key"]["normalized_scope"])
        for item in active
    ]
    safety_ids = sorted(
        item["property_id"] for item in registry["properties"] if item.get("safety_invariant")
    )
    return {
        "resolution_id": proposal["proposed_resolution_id"],
        "key": deepcopy(proposal["key"]),
        "resolved_binding": deepcopy(candidate["property_binding"]),
        "final_layer_kind": candidate["layer_kind"],
        "final_source": deepcopy(candidate["source"]),
        "candidate_chain": active,
        "override_chain": _override_chain(wrappers),
        "excluded_candidates": deepcopy(list(excluded or [])),
        "confidence": confidence,
        "safety_check": {"status": "pass", "checked_invariant_ids": safety_ids},
        "execution_mode": candidate["property_binding"]["mode"],
    }


def _approval_target_key(approval: Mapping[str, Any]) -> tuple[str, str, str, str]:
    target = approval["target"]
    normalized = normalized_property_scope_key(
        "conflict", target["conflict_id"], target["normalized_scope"]
    )
    return (*normalized, target["proposed_resolution_id"])


def _active_approvals(
    approvals: Sequence[dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]]]:
    diagnostics: list[dict[str, Any]] = []
    by_id: dict[str, dict[str, Any]] = {}
    for approval in approvals:
        approval_id = approval.get("approval_id")
        if not isinstance(approval_id, str) or approval_id in by_id:
            diagnostics.append(_application_diagnostic("approval_chain", ["duplicate", approval_id]))
            continue
        by_id[approval_id] = approval
    by_target: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    for approval in by_id.values():
        try:
            by_target.setdefault(_approval_target_key(approval), []).append(approval)
        except (KeyError, ScopeContractError) as exc:
            diagnostics.append(_application_diagnostic("approval_binding", str(exc)))
    active: dict[str, dict[str, Any]] = {}
    for items in by_target.values():
        item_ids = {item["approval_id"] for item in items}
        successors: dict[str, list[str]] = {item_id: [] for item_id in item_ids}
        invalid = False
        for item in items:
            previous = item.get("previous_approval_id")
            if previous is not None:
                if previous not in item_ids:
                    diagnostics.append(
                        _application_diagnostic("approval_chain", ["unknown-predecessor", previous], item["target"].get("conflict_id"))
                    )
                    invalid = True
                else:
                    successors[previous].append(item["approval_id"])
        if any(len(values) > 1 for values in successors.values()):
            diagnostics.append(_application_diagnostic("approval_chain", "branch", items[0]["target"].get("conflict_id")))
            invalid = True
        terminals = [item_id for item_id, values in successors.items() if not values]
        if len(terminals) != 1:
            diagnostics.append(_application_diagnostic("approval_chain", ["terminal-count", len(terminals)], items[0]["target"].get("conflict_id")))
            invalid = True
        if not invalid:
            terminal = terminals[0]
            visited: set[str] = set()
            cursor: str | None = terminal
            while cursor is not None:
                if cursor in visited:
                    diagnostics.append(_application_diagnostic("approval_chain", "cycle", items[0]["target"].get("conflict_id")))
                    invalid = True
                    break
                visited.add(cursor)
                cursor = by_id[cursor].get("previous_approval_id")
            if visited != item_ids:
                diagnostics.append(_application_diagnostic("approval_chain", "disconnected", items[0]["target"].get("conflict_id")))
                invalid = True
        if not invalid:
            terminal_document = by_id[terminals[0]]
            conflict_id = terminal_document["target"]["conflict_id"]
            if conflict_id in active:
                diagnostics.append(_application_diagnostic("approval_chain", "multiple-target-chains", conflict_id))
            else:
                active[conflict_id] = terminal_document
    return active, diagnostics


def _apply_exclusion(
    proposal: Mapping[str, Any],
    conflict: Mapping[str, Any],
    approval: Mapping[str, Any],
    registry: dict[str, Any],
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    candidate_id = approval["target"].get("candidate_id")
    chain = [deepcopy(item) for item in proposal["candidate_chain"]]
    target = next((item for item in chain if item["candidate_id"] == candidate_id), None)
    if target is None:
        return None, _application_diagnostic("approval_binding", "unknown-exclusion-candidate", conflict["conflict_id"])
    remaining = [item for item in chain if item["candidate_id"] != candidate_id]
    if not remaining:
        return None, _application_diagnostic("qa_exclusion", "no-candidates-remain", conflict["conflict_id"])
    winning_rank = max(authority_rank(item["layer_kind"]) for item in remaining)
    winning = [item for item in remaining if authority_rank(item["layer_kind"]) == winning_rank]
    wrapped = [
        _Candidate(item, proposal["key"]["semantic_object_kind"], proposal["key"]["property_id"], proposal["key"]["normalized_scope"])
        for item in winning
    ]
    groups = _binding_groups(wrapped, registry)
    if len(groups) != 1:
        return None, _application_diagnostic("qa_exclusion", "winner-not-unique", conflict["conflict_id"])
    selected = min(groups[0], key=_candidate_sort_key).record
    if _confidence(groups[0]) != "high":
        return None, _application_diagnostic("qa_exclusion", "winner-confidence-not-high", conflict["conflict_id"])
    excluded_candidate = deepcopy(target)
    excluded_candidate["scope_status"] = "excluded"
    resolved = _resolved_from_candidate(
        proposal,
        selected,
        registry,
        chain=remaining,
        excluded=[
            {
                "candidate": excluded_candidate,
                "exclusion_reason": "qa_exclusion",
                "reason_code": _reason_code("QA-EXCLUSION", candidate_id),
            }
        ],
    )
    return resolved, None


def _apply_decision(
    proposal: Mapping[str, Any],
    conflict: Mapping[str, Any],
    approval: Mapping[str, Any],
    registry: dict[str, Any],
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    decision = approval["decision"]
    if decision == "adopt_proposed":
        return _proposal_to_resolved(proposal, registry), None
    if decision == "select_candidate":
        candidate_id = approval["target"].get("candidate_id")
        candidate = next(
            (item for item in conflict["candidates"] if item["candidate_id"] == candidate_id),
            None,
        )
        if candidate is None:
            return None, _application_diagnostic("approval_binding", "unknown-selected-candidate", conflict["conflict_id"])
        return _resolved_from_candidate(proposal, candidate, registry), None
    if decision == "keep_original":
        preserve = [
            item
            for item in proposal["candidate_chain"]
            if item["property_binding"]["mode"] == "preserve"
        ]
        if not preserve:
            return None, _application_diagnostic("keep_original", "preserve-candidate-unavailable", conflict["conflict_id"])
        selected = max(
            preserve,
            key=lambda item: (authority_rank(item["layer_kind"]), item["candidate_id"]),
        )
        return _resolved_from_candidate(proposal, selected, registry), None
    if decision == "exclude_candidate":
        return _apply_exclusion(proposal, conflict, approval, registry)
    return None, _application_diagnostic("approval_binding", "unsupported-decision", conflict["conflict_id"])


def _application_result(
    status: str,
    report: dict[str, Any],
    final_profile: dict[str, Any] | None,
    diagnostics: Sequence[dict[str, Any]],
) -> dict[str, Any]:
    if status not in APPLICATION_STATUSES:
        raise ComposerContractError(f"Unsupported application status: {status}")
    return {
        "status": status,
        "report": deepcopy(report),
        "final_profile": deepcopy(final_profile),
        "application_diagnostics": sorted(
            (deepcopy(item) for item in diagnostics), key=lambda item: item["diagnostic_id"]
        ),
    }


def _apply_report(
    report: dict[str, Any],
    approvals: Sequence[dict[str, Any]],
    *,
    adapter: _ContractAdapter,
    task_id: str,
    task_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
) -> dict[str, Any]:
    original_report = deepcopy(report)
    original_approvals = deepcopy(list(approvals))
    try:
        adapter.validate(report)
    except (ArtifactContractError, ValueError) as exc:
        return _application_result(
            "fatal",
            original_report,
            None,
            [_application_diagnostic("invalid_report", str(exc))],
        )
    if report["fatal_diagnostics"]:
        return _application_result("fatal", original_report, None, [])
    if report["unresolvable_blockers"] or any(
        item["evidence_status"] != "verified" for item in report["scope_partitions"]
    ):
        return _application_result("unresolvable", original_report, None, [])
    if report["proposal_status"] not in {"awaiting_approval", "resolvable"}:
        return _application_result(
            "fatal",
            original_report,
            None,
            [_application_diagnostic("invalid_report", "proposal-status")],
        )

    diagnostics: list[dict[str, Any]] = []
    valid_approvals: list[dict[str, Any]] = []
    for approval in original_approvals:
        try:
            adapter.validate(approval)
        except (ArtifactContractError, ValueError) as exc:
            diagnostics.append(_application_diagnostic("invalid_approval", str(exc)))
            continue
        if approval["bindings"]["composition_report_fingerprint"] != report["semantic_fingerprint"]:
            diagnostics.append(_application_diagnostic("approval_binding", "stale-report"))
            continue
        if approval["bindings"]["input_fingerprint"] != report["bindings"]["input_fingerprint"]:
            diagnostics.append(_application_diagnostic("approval_binding", "input-mismatch"))
            continue
        if approval["bindings"]["structure_fingerprint"] != report["bindings"]["structure_fingerprint"]:
            diagnostics.append(_application_diagnostic("approval_binding", "structure-mismatch"))
            continue
        valid_approvals.append(approval)

    active, chain_diagnostics = _active_approvals(valid_approvals)
    diagnostics.extend(chain_diagnostics)
    conflicts = {item["conflict_id"]: item for item in report["approval_required_conflicts"]}
    proposals = {
        item["proposed_resolution_id"]: item for item in report["proposed_resolutions"]
    }
    if set(active) - set(conflicts):
        diagnostics.append(_application_diagnostic("approval_binding", "unknown-conflict"))

    resolved: dict[str, dict[str, Any]] = {
        proposal_id: _proposal_to_resolved(proposal, adapter.registry)
        for proposal_id, proposal in proposals.items()
    }
    closure: list[dict[str, Any]] = []
    terminal_approvals: list[dict[str, Any]] = []
    for conflict_id, conflict in sorted(conflicts.items()):
        approval = active.get(conflict_id)
        if approval is None:
            diagnostics.append(_application_diagnostic("approval_required", "missing", conflict_id))
            continue
        target = approval["target"]
        proposal_id = conflict["proposed_resolution_id"]
        proposal = proposals[proposal_id]
        if target["proposed_resolution_id"] != proposal_id:
            diagnostics.append(_application_diagnostic("approval_binding", "resolution-mismatch", conflict_id))
            continue
        try:
            target_key = normalized_property_scope_key("conflict", conflict_id, target["normalized_scope"])
            conflict_key = normalized_property_scope_key("conflict", conflict_id, conflict["key"]["normalized_scope"])
        except ScopeContractError as exc:
            diagnostics.append(_application_diagnostic("approval_binding", str(exc), conflict_id))
            continue
        if target_key != conflict_key or approval["decision"] not in conflict["allowed_decisions"]:
            diagnostics.append(_application_diagnostic("approval_binding", "scope-or-decision-mismatch", conflict_id))
            continue
        candidate_id = target.get("candidate_id")
        candidate_ids = {item["candidate_id"] for item in conflict["candidates"]}
        if candidate_id is not None and candidate_id not in candidate_ids:
            diagnostics.append(_application_diagnostic("approval_binding", "candidate-mismatch", conflict_id))
            continue
        applied, error = _apply_decision(proposal, conflict, approval, adapter.registry)
        if error is not None:
            diagnostics.append(error)
            continue
        resolved[proposal_id] = applied
        terminal_approvals.append(approval)
        closure.append(
            {
                "conflict_id": conflict_id,
                "proposed_resolution_id": proposal_id,
                "qa_decision_id": approval["approval_id"],
            }
        )

    if diagnostics or len(closure) != len(conflicts):
        return _application_result("awaiting_approval", original_report, None, diagnostics)

    final_items = sorted(resolved.values(), key=lambda item: item["resolution_id"])
    for left_index, left in enumerate(final_items):
        for right in final_items[left_index + 1 :]:
            if (
                left["key"]["semantic_object_kind"] == right["key"]["semantic_object_kind"]
                and left["key"]["property_id"] == right["key"]["property_id"]
                and not scope_disjoint(
                    left["key"]["normalized_scope"], right["key"]["normalized_scope"]
                )
            ):
                return _application_result(
                    "fatal",
                    original_report,
                    None,
                    [_application_diagnostic("non_disjoint_final", [left["resolution_id"], right["resolution_id"]])],
                )
    _validate_fingerprint(task_fingerprint, "task_fingerprint")
    approval_fingerprints = sorted(
        item["semantic_fingerprint"] for item in terminal_approvals
    )
    bindings = {
        "task_fingerprint": task_fingerprint,
        "input_fingerprint": report["bindings"]["input_fingerprint"],
        "feature_activation_fingerprint": report["bindings"]["feature_activation_fingerprint"],
        "property_registry_fingerprint": report["bindings"]["property_registry_fingerprint"],
        "authority_contract_fingerprint": report["bindings"]["authority_contract_fingerprint"],
        "rule_asset_fingerprints": list(report["bindings"]["rule_asset_fingerprints"]),
        "structure_fingerprint": report["bindings"]["structure_fingerprint"],
        "approval_fingerprints": approval_fingerprints,
        "composition_report_fingerprint": report["semantic_fingerprint"],
    }
    inputs = [
        _input("input:task", "task", task_fingerprint),
        _input("input:source", "source_document", bindings["input_fingerprint"]),
        _input("input:feature", "feature_activation", bindings["feature_activation_fingerprint"]),
        _input("input:registry", "property_registry", bindings["property_registry_fingerprint"]),
        _input("input:authority", "authority_contract", bindings["authority_contract_fingerprint"]),
        _input("input:structure", "structure", bindings["structure_fingerprint"]),
        _input("input:report", "conflict_report", bindings["composition_report_fingerprint"]),
    ]
    inputs.extend(
        _input(_stable_id("input-rule", value), "rule_asset", value)
        for value in bindings["rule_asset_fingerprints"]
    )
    inputs.extend(
        _input(_stable_id("input-approval", value), "approval", value)
        for value in approval_fingerprints
    )
    final = {
        "artifact_kind": "final-execution-profile",
        "schema_version": "2.2",
        "registry_contract_version": "2.1",
        "authority_contract_version": "1.0",
        "artifact_id": artifact_id,
        "created_by_tool": deepcopy(created_by_tool),
        "input_fingerprints": sorted(inputs, key=lambda item: item["input_id"]),
        "semantic_fingerprint": _digest("unstamped-final-profile"),
        "task_id": task_id,
        "legacy_input": False,
        "activation": "disabled",
        "final_ready_eligible": False,
        "delivery_allowed": False,
        "safety_invariants": {
            "overridable": False,
            "author_content_mutation_allowed": False,
            "legacy_auto_activation_allowed": False,
            "delivery_evidence_allowed": False,
            "final_ready_allowed": False,
        },
        "bindings": bindings,
        "resolved_properties": final_items,
        "closure_evidence": sorted(closure, key=lambda item: item["conflict_id"]),
    }
    stamped = adapter.stamp(final)
    adapter.validate(stamped)
    return _application_result("profile_generated", original_report, stamped, [])


def apply_resolutions(
    report: dict[str, Any],
    approvals: Sequence[dict[str, Any]],
    *,
    task_id: str,
    task_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
) -> dict[str, Any]:
    """Apply already-bound approvals without rereading rule assets or repartitioning."""

    return _apply_report(
        report,
        approvals,
        adapter=_production_adapter(),
        task_id=task_id,
        task_fingerprint=task_fingerprint,
        artifact_id=artifact_id,
        created_by_tool=created_by_tool,
    )


def _apply_resolutions_for_test(
    report: dict[str, Any],
    approvals: Sequence[dict[str, Any]],
    *,
    registry: dict[str, Any],
    task_id: str,
    task_fingerprint: str,
    artifact_id: str,
    created_by_tool: dict[str, Any],
) -> dict[str, Any]:
    return _apply_report(
        report,
        approvals,
        adapter=_test_adapter(registry),
        task_id=task_id,
        task_fingerprint=task_fingerprint,
        artifact_id=artifact_id,
        created_by_tool=created_by_tool,
    )
