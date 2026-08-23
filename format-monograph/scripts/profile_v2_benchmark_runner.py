#!/usr/bin/env python3
"""Internal, serial C2B benchmark runner. It never authorizes a benchmark GO."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import queue
import subprocess
import sys
import threading
import time
from statistics import median
from copy import deepcopy
from decimal import Decimal, ROUND_HALF_EVEN
from pathlib import Path
from typing import Any, Iterable

from profile_v2_benchmark import (
    BenchmarkContractError,
    FROZEN_RSS_PROTOCOL,
    FROZEN_SCENARIO_SEMANTICS,
    FROZEN_THRESHOLDS,
    canonical_json_bytes,
    derive_stop_reasons,
    recompute_result_digest,
    recompute_subject_digest,
    validate_benchmark_campaign_context,
    validate_benchmark_result_context,
    validate_benchmark_result_semantics,
    validate_complete_benchmark_suite,
)
from profile_v2_canonical import canonical_data_digest, stamp_intent_semantic_fingerprint_v041
from profile_v2_composer import apply_intent_resolutions_v041, compose_intent_profile_v041
from profile_v2_artifacts import validate_intent_artifact_v041
from profile_v2_registry import load_registry
from profile_v2_scope import normalize_scope, normalized_property_scope_key


class BenchmarkRunnerError(ValueError):
    pass


class _CampaignBudgetStopped(Exception):
    """Internal signal: a child must not start after the campaign deadline."""


SCALES = {"0.5x": Decimal("0.5"), "1.0x": Decimal("1"), "1.5x": Decimal("1.5"), "2.0x": Decimal("2")}
SCENARIOS = ("disjoint", "subset-chain", "dense-crossing", "mixed-conflict-approval")
MICRO_SCENARIO_EXPECTATIONS = {
    "disjoint": {"source_keys": 2, "terminal": "profile_generated", "conflicts": 0, "proposals": 2, "blockers": 0},
    "subset-chain": {"source_keys": 2, "terminal": "profile_generated", "conflicts": 0, "proposals": 2, "blockers": 0},
    "dense-crossing": {"source_keys": 2, "terminal": "unresolvable", "conflicts": 0, "proposals": 0, "blockers": 1},
    "mixed-conflict-approval": {"source_keys": 2, "terminal": "profile_generated", "conflicts": 2, "proposals": 2, "blockers": 0},
}
SUBJECT_PATHS = (
    "format-monograph/scripts/profile_v2_benchmark_runner.py",
    "format-monograph/scripts/profile_v2_benchmark.py",
    "format-monograph/scripts/profile_v2_composer.py",
    "format-monograph/scripts/profile_v2_artifacts.py",
    "format-monograph/scripts/profile_v2_authority.py",
    "format-monograph/scripts/profile_v2_canonical.py",
    "format-monograph/scripts/profile_v2_registry.py",
    "format-monograph/scripts/profile_v2_scope.py",
    "format-monograph/scripts/profile_v2_values.py",
    "format-monograph/references/schemas/v2/artifact-contract-matrix.v1.1.json",
    "format-monograph/references/schemas/v2/artifact-contract-matrix.v1.1.schema.json",
    "format-monograph/references/schemas/v2/authority-contract.v1.0.json",
    "format-monograph/references/schemas/v2/common.v2.3.schema.json",
    "format-monograph/references/schemas/v2/conflict-report.v2.3.schema.json",
    "format-monograph/references/schemas/v2/feature-activation-manifest.v2.2.schema.json",
    "format-monograph/references/schemas/v2/final-execution-profile.v2.3.schema.json",
    "format-monograph/references/schemas/v2/layered-rule-asset.v2.2.schema.json",
    "format-monograph/references/schemas/v2/property-catalog.v2.2.generated.schema.json",
    "format-monograph/references/schemas/v2/property-registry.v2.2.core.json",
    "format-monograph/references/schemas/v2/property-registry.v2.2.schema.json",
    "format-monograph/references/schemas/v2/qa-approval-artifact.v2.2.schema.json",
    "format-monograph/references/schemas/v2/typed-value.v2.2.generated.schema.json",
    "format-monograph/references/benchmarks/v0412/p3a-c2/projected-envelope.schema.json",
    "format-monograph/references/benchmarks/v0412/p3a-c2/benchmark-config.schema.json",
    "format-monograph/references/benchmarks/v0412/p3a-c2/benchmark-result.schema.json",
)


def _sha(data: bytes) -> str:
    return "sha256:" + hashlib.sha256(data).hexdigest()


def _round_count(value: int, scale_id: str) -> int:
    if scale_id not in SCALES:
        raise BenchmarkRunnerError("unknown scale")
    return int((Decimal(value) * SCALES[scale_id]).quantize(Decimal("1"), rounding=ROUND_HALF_EVEN))


def scaled_workload(config: dict[str, Any], scale_id: str) -> dict[str, int]:
    counts = config["projection_binding"]["aggregate_counts"]
    result = {name: _round_count(counts[name], scale_id) for name in ("rule_fragment", "binding", "key", "candidate")}
    if any(value <= 0 for value in result.values()) or result["binding"] != result["candidate"]:
        raise BenchmarkRunnerError("scaled workload is not composable")
    return result


def _scope(object_ids: str | list[str], *, document: bool = False) -> dict[str, Any]:
    ids = [object_ids] if isinstance(object_ids, str) else object_ids
    selectors = [{"selector_kind": "object", "selector_ids": ids}]
    if document:
        selectors.insert(0, {"selector_kind": "document", "selector_ids": ["document:c2b"]})
    return {"selectors": selectors, "exclusions": [], "mutually_exclusive_conditions": []}


def _tool() -> dict[str, str]:
    return {"tool_id": "format-monograph", "version": "0.4.1"}


def feature_manifest() -> dict[str, Any]:
    document = {"artifact_kind": "feature-activation-manifest", "schema_version": "2.2", "registry_contract_version": "2.2", "authority_contract_version": "1.0", "artifact_id": "feature-activation-manifest:c2b", "created_by_tool": _tool(), "input_fingerprints": [{"input_id": "input:c2b", "role": "source_document", "fingerprint": _sha(b"c2b-feature")}], "semantic_fingerprint": _sha(b"unstamped"), "features": {"profile_v2_schema": True, "profile_v2_composer": True, "monograph_base_v041": True, "final_ready_eligible": False}}
    return stamp_intent_semantic_fingerprint_v041(document)


def _scope_for_source_key(scenario_id: str, key_index: int) -> dict[str, Any]:
    """Create one normalized source scope per key index.

    The C1 ``expected_key_count`` contract is the number of normalized source
    composition keys, before any scope partitioning.  In particular, selector
    ordering is not a semantic distinction.
    """
    if scenario_id == "dense-crossing":
        # Adjacent two-object unions overlap without either being a subset.
        return _scope([
            "object:%04d" % key_index,
            "object:%04d" % (key_index + 1),
        ])
    if scenario_id == "subset-chain":
        # Each adjacent pair shares an object target; the odd member also has
        # a document selector and is therefore the narrower source scope.
        object_id = "object:%04d" % (key_index // 2)
        return _scope(object_id, document=key_index % 2 == 1)
    return _scope("object:%04d" % key_index)


def _value_for_rule(scenario_id: str, key_index: int, occurrence: int) -> str:
    if scenario_id == "mixed-conflict-approval" and occurrence % 2:
        return "alternative:%04d" % key_index
    return "value:%04d" % key_index


def normalized_source_key_count(assets: Iterable[dict[str, Any]]) -> int:
    """Count C1 source composition keys before scope partitioning."""
    keys: set[tuple[str, str, str]] = set()
    for asset in assets:
        for rule in asset.get("rules", []):
            scope = normalize_scope(rule["scope"])
            for binding in rule.get("properties", []):
                keys.add(normalized_property_scope_key(
                    rule["semantic_object_kind"], binding["property_id"], scope
                ))
    return len(keys)


def _permutation_order(items: list[dict[str, Any]], seed: int | None) -> list[dict[str, Any]]:
    """Apply a stable, seed-specific permutation without changing semantics."""
    if seed is None:
        return items
    def stable_id(item: dict[str, Any]) -> str:
        value = item.get("artifact_id", item.get("rule_id"))
        if not isinstance(value, str) or not value:
            raise BenchmarkRunnerError("permutation input lacks stable identity")
        return value
    return sorted(items, key=lambda item: hashlib.sha256(
        (str(seed) + "\0" + stable_id(item)).encode("utf-8")
    ).hexdigest())


def generate_assets(config: dict[str, Any], scale_id: str, scenario_id: str, seed: int, *, permutation_seed: int | None = None) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Generate only C1's contract.intent-probe assets; never a base rule asset."""
    if scenario_id not in SCENARIOS:
        raise BenchmarkRunnerError("unknown scenario")
    work = scaled_workload(config, scale_id)
    if scenario_id == "mixed-conflict-approval" and work["binding"] < 2 * work["key"]:
        raise BenchmarkRunnerError("mixed-conflict-approval requires two bindings per source key")
    if scenario_id == "dense-crossing" and work["key"] < 2:
        raise BenchmarkRunnerError("dense-crossing requires at least two source keys")
    registry = load_registry(version="2.2", validation_context="declaration_intent")
    registry_fingerprint = canonical_data_digest(registry)
    rules_per_asset = [work["binding"] // work["rule_fragment"]] * work["rule_fragment"]
    for index in range(work["binding"] % work["rule_fragment"]):
        rules_per_asset[index] += 1
    assets = []
    rule_index = 0
    occurrences = [0] * work["key"]
    for asset_index, count in enumerate(rules_per_asset):
        rules = []
        for _ in range(count):
            key_index = rule_index % work["key"]
            generated_key = key_index + seed * (work["key"] + 1)
            scope = _scope_for_source_key(scenario_id, generated_key)
            value = _value_for_rule(scenario_id, generated_key, occurrences[key_index])
            occurrences[key_index] += 1
            rules.append({"rule_id": "RULE-C2B-%04d" % rule_index, "semantic_object_kind": "paragraph", "scope": scope, "confidence": "high", "status": "approved", "properties": [{"property_id": "contract.intent-probe", "value": {"type": "string", "value": value}, "unit_id": None, "mode": "automatic"}]})
            rule_index += 1
        asset = {"artifact_kind": "layered-rule-asset", "schema_version": "2.2", "registry_contract_version": "2.2", "authority_contract_version": "1.0", "artifact_id": "layered-rule-asset:c2b-%04d" % asset_index, "created_by_tool": _tool(), "input_fingerprints": [{"input_id": "input:c2b-asset-%04d" % asset_index, "role": "rule_asset", "fingerprint": _sha(("c2b-%s-%s-%d" % (scale_id, scenario_id, asset_index)).encode())}, {"input_id": "input:c2b-registry", "role": "property_registry", "fingerprint": registry_fingerprint}], "semantic_fingerprint": _sha(b"unstamped"), "property_registry_binding": {"registry_fingerprint": registry_fingerprint, "required_completeness": "contract_core"}, "layer_kind": "monograph_base", "can_override_safety_invariants": False, "activation": "approved", "asset_scope": {"selectors": [{"selector_kind": "document", "selector_ids": ["document:c2b"]}], "exclusions": [], "mutually_exclusive_conditions": []}, "allowed_semantic_object_kinds": ["paragraph"], "rules": rules}
        assets.append(stamp_intent_semantic_fingerprint_v041(asset))
    if normalized_source_key_count(assets) != work["key"]:
        raise BenchmarkRunnerError("generated source keys do not conserve scaled(key)")
    for index, asset in enumerate(assets):
        asset["rules"] = _permutation_order(asset["rules"], permutation_seed)
        assets[index] = stamp_intent_semantic_fingerprint_v041(asset)
    return _permutation_order(assets, permutation_seed), work


def _approval(report: dict[str, Any], ordinal: int) -> dict[str, Any]:
    conflict = report["approval_required_conflicts"][ordinal]
    document = {"artifact_kind": "qa-approval-artifact", "schema_version": "2.2", "registry_contract_version": "2.2", "authority_contract_version": "1.0", "artifact_id": "qa-approval-artifact:c2b-%04d" % ordinal, "created_by_tool": _tool(), "input_fingerprints": [{"input_id": "input:c2b-source", "role": "source_document", "fingerprint": report["bindings"]["input_fingerprint"]}, {"input_id": "input:c2b-structure", "role": "structure", "fingerprint": report["bindings"]["structure_fingerprint"]}, {"input_id": "input:c2b-report", "role": "conflict_report", "fingerprint": report["semantic_fingerprint"]}], "semantic_fingerprint": _sha(b"unstamped"), "approval_id": "approval:c2b-%04d" % ordinal, "approver": {"actor_id": "actor:c2b-user", "actor_role": "user"}, "decision_type": "conflict_resolution", "decision": "adopt_proposed", "reason": "Deterministic synthetic C2B approval.", "created_at": "2026-08-23T00:00:00Z", "bindings": {"input_fingerprint": report["bindings"]["input_fingerprint"], "structure_fingerprint": report["bindings"]["structure_fingerprint"], "composition_report_fingerprint": report["semantic_fingerprint"]}, "target": {"conflict_id": conflict["conflict_id"], "proposed_resolution_id": conflict["proposed_resolution_id"], "normalized_scope": deepcopy(conflict["key"]["normalized_scope"])}, "previous_approval_id": None}
    return stamp_intent_semantic_fingerprint_v041(document)


def compose_cell(config: dict[str, Any], *, scale_id: str, scenario_id: str, generation_seed: int, permutation_seed: int | None = None) -> dict[str, Any]:
    """Pure worker core: compose one C2B logical cell without a supervisor or GO."""
    generated_at = time.perf_counter()
    assets, expected = generate_assets(config, scale_id, scenario_id, generation_seed, permutation_seed=permutation_seed)
    generation_seconds = time.perf_counter() - generated_at
    manifest = feature_manifest()
    validation_at = time.perf_counter()
    for asset in assets:
        validate_intent_artifact_v041(asset)
    validate_intent_artifact_v041(manifest)
    validation_seconds = time.perf_counter() - validation_at
    composed_at = time.perf_counter()
    composed = compose_intent_profile_v041(assets, manifest, input_fingerprint=_sha(b"c2b-source"), structure_fingerprint=_sha(b"c2b-structure"), artifact_id="conflict-report:c2b-%s-%s" % (scale_id, scenario_id), created_by_tool=_tool(), generated_at="2026-08-23T00:00:00Z", include_metrics=True)
    compose_seconds = time.perf_counter() - composed_at
    metrics = composed.metrics
    metric_expectations = {"input_asset_count": expected["rule_fragment"], "input_rule_count": expected["binding"], "input_binding_count": expected["binding"], "expected_key_count": expected["key"], "candidate_count": expected["candidate"]}
    if metrics is None or any(getattr(metrics, field) != value for field, value in metric_expectations.items()):
        raise BenchmarkRunnerError("C1 metrics do not conserve generated workload")
    final = None
    pre_status = composed.status
    status = pre_status
    if scenario_id != "dense-crossing":
        approvals_at = time.perf_counter()
        approvals = [_approval(composed.report, index) for index in range(len(composed.report["approval_required_conflicts"]))]
        approval_seconds = time.perf_counter() - approvals_at
        apply_at = time.perf_counter()
        applied = apply_intent_resolutions_v041(composed.report, approvals, manifest, task_id="task:c2b", task_fingerprint=_sha(b"c2b-task"), artifact_id="final-execution-profile:c2b", created_by_tool=_tool(), metrics=metrics)
        apply_seconds = time.perf_counter() - apply_at
        status, final = applied.status, applied.final_profile
    else:
        approval_seconds, apply_seconds = 0.0, 0.0
    if scenario_id in {"disjoint", "subset-chain"} and (status != "profile_generated" or final is None):
        raise BenchmarkRunnerError("unexpected final-state boundary")
    if scenario_id == "dense-crossing" and (status != "unresolvable" or final is not None):
        raise BenchmarkRunnerError("dense crossing boundary")
    if scenario_id == "mixed-conflict-approval" and (status != "profile_generated" or final is None):
        raise BenchmarkRunnerError("mixed approval boundary")
    if scenario_id == "mixed-conflict-approval" and (
        not composed.report["approval_required_conflicts"]
        or not composed.report["proposed_resolutions"]
    ):
        raise BenchmarkRunnerError("mixed scenario did not produce approval-required evidence")
    return {"assets": assets, "report": composed.report, "final_profile": final, "metrics": metrics, "status": status, "pre_status": pre_status, "expected": expected, "stage_seconds": {"synthetic_generation": generation_seconds, "schema_registry_validation": validation_seconds, "compose": compose_seconds, "approval_generation": approval_seconds, "apply": apply_seconds}}


def build_subject_manifest(benchmark_subject_commit: str, *, repository: Path | None = None) -> list[dict[str, str]]:
    if len(benchmark_subject_commit) != 40 or any(char not in "0123456789abcdef" for char in benchmark_subject_commit):
        raise BenchmarkRunnerError("subject commit must be lowercase SHA-1")
    root = repository or Path(__file__).resolve().parents[2]
    manifest = []
    for path in SUBJECT_PATHS:
        completed = subprocess.run(["git", "-C", str(root), "show", "%s:%s" % (benchmark_subject_commit, path)], capture_output=True)
        if completed.returncode != 0:
            raise BenchmarkRunnerError("subject path absent from commit: " + path)
        manifest.append({"path": path, "sha256": _sha(completed.stdout)})
    if tuple(item["path"] for item in manifest) != tuple(sorted(SUBJECT_PATHS)):
        manifest.sort(key=lambda item: item["path"])
    if len({item["path"] for item in manifest}) != len(SUBJECT_PATHS):
        raise BenchmarkRunnerError("subject path duplication")
    if recompute_subject_digest(manifest) != recompute_subject_digest(sorted(manifest, key=lambda item: item["path"])):
        raise BenchmarkRunnerError("subject manifest mismatch")
    return manifest


def logical_key(parameters: dict[str, Any]) -> str:
    return "%s-%s-%s-%s" % (parameters["measurement_kind"], parameters["scale_id"], parameters["scenario_id"], parameters["permutation_seed"] if parameters["permutation_seed"] is not None else "none")


def _measured(seconds: float) -> dict[str, Any]:
    return {"status": "measured", "wall_seconds": float(seconds)}


def _not_applicable() -> dict[str, str]:
    return {"status": "not_applicable"}


def _not_reached() -> dict[str, str]:
    return {"status": "not_reached"}


def _run_kind(measurement_kind: str, ordinal: int) -> str:
    if measurement_kind == "performance":
        return "performance_warmup" if ordinal == 1 else "performance_measured"
    return measurement_kind


def _run_record(
    *,
    ordinal: int,
    measurement_kind: str,
    scenario_id: str,
    supervised: dict[str, Any],
    worker: dict[str, Any] | None,
    elapsed_seconds: float,
    input_json_bytes: int | None,
    output_json_bytes: int | None,
    determinism: tuple[str, str] = ("not_applicable", "not_applicable"),
) -> dict[str, Any]:
    """Turn one supervisor observation into the C2A run shape.

    All result arithmetic is reconstructed from these records; callers never
    supply summary, ratio, or determinism values as placeholders.
    """
    # A worker can finish validly while external RSS evidence is unavailable.
    # Preserve its actual composition evidence; the enclosing result remains
    # stopped solely because C2A derives ``rss_unavailable``.
    completed = supervised["status"] in {"completed", "rss_unavailable"} and worker is not None
    final_present = bool(worker and worker["final_present"])
    terminal = FROZEN_SCENARIO_SEMANTICS[scenario_id]
    timed = _measured(elapsed_seconds)
    if completed:
        stages = worker["stage_seconds"]
        approval = _measured(stages["approval_generation"]) if scenario_id == "mixed-conflict-approval" else _not_applicable()
        apply = _measured(stages["apply"]) if final_present else _not_applicable()
        state = terminal["terminal_state"]
        trace = {"pre_approval": terminal["pre_approval_terminal"], "post_approval": terminal["post_approval_terminal"]}
        metrics = worker["metrics"]
        conservation, stable, contract = "passed", "stable", "valid"
    else:
        approval = _not_reached() if scenario_id == "mixed-conflict-approval" else _not_applicable()
        apply = _not_reached() if scenario_id != "dense-crossing" else _not_applicable()
        state, trace, final_present = "not_reached", {"pre_approval": "not_reached", "post_approval": "not_reached"}, False
        metrics, conservation, stable, contract = None, "not_reached", "not_reached", "error"
        determinism = ("not_reached", "not_reached") if measurement_kind == "determinism" else determinism
    return {
        "run_index": ordinal, "run_kind": _run_kind(measurement_kind, ordinal),
        "run_status": "completed" if completed else ("timeout" if supervised["status"] == "timeout" else "process_crash"),
        "timings": {
            "synthetic_generation": _measured(stages["synthetic_generation"]) if completed else _not_reached(),
            "schema_registry_validation": _measured(stages["schema_registry_validation"]) if completed else _not_reached(),
            "compose": _measured(stages["compose"]) if completed else _not_reached(),
            "approval_generation": approval, "apply": apply,
            "canonical_serialization": _measured(stages["canonical_serialization"]) if completed else _not_reached(),
            "end_to_end": timed,
        },
        "rss": supervised.get("rss", {"status": "unavailable"}),
        "input_json_bytes": input_json_bytes if completed else None,
        "output_json_bytes": output_json_bytes if completed else None,
        "metrics": metrics, "terminal_state": state, "terminal_trace": trace,
        "coverage_conservation": conservation, "stable_id_status": stable,
        "canonical_determinism": determinism[0], "fingerprint_determinism": determinism[1],
        "contract_status": contract, "final_profile_present": final_present,
        "final_profile_fingerprint": worker["final_fingerprint"] if completed and final_present else None,
    }


def build_result(
    *,
    cell: dict[str, Any],
    supervised: dict[str, Any],
    config: dict[str, Any],
    subject_commit: str,
    subject_manifest: list[dict[str, str]],
    parameters: dict[str, Any],
    elapsed_seconds: float,
    environment: dict[str, Any] | None = None,
    determinism: tuple[str, str] = ("not_applicable", "not_applicable"),
) -> dict[str, Any]:
    """Construct one non-artifact C2A result; it never decides suite GO."""
    worker = supervised.get("worker")
    scenario = parameters["scenario_id"]
    input_bytes = len(canonical_json_bytes(cell["assets"])) if cell else (worker or {}).get("input_json_bytes")
    output_bytes = len(canonical_json_bytes({"report": cell["report"], "final_profile": cell["final_profile"]})) if cell else (worker or {}).get("output_json_bytes")
    # Multi-run performance evidence is assembled only from the supervisor's
    # actual observations in _result_from_observations.
    run_count = 1
    runs = [_run_record(ordinal=index, measurement_kind=parameters["measurement_kind"], scenario_id=scenario, supervised=supervised, worker=worker, elapsed_seconds=elapsed_seconds, input_json_bytes=input_bytes, output_json_bytes=output_bytes, determinism=determinism) for index in range(1, run_count + 1)]
    result = {
        "document_kind": "p3a_c2_benchmark_result", "contract_version": "1.0",
        "evidence_kind": "non_artifact_benchmark", "benchmark_subject_commit": subject_commit,
        "benchmark_subject_digest": recompute_subject_digest(subject_manifest),
        "benchmark_subject_digest_basis": "canonical_subject_manifest_v1",
        "subject_manifest": subject_manifest,
        "subject_digest_status": {
            "state": "current", "observed_subject_digest": recompute_subject_digest(subject_manifest),
            "revalidation_required": False,
        },
        "benchmark_config_digest": config["config_digest"], "result_digest": _sha(b"unstamped"),
        "result_digest_basis": "canonical_json_excluding_result_digest",
        "environment": deepcopy(environment if environment is not None else _runtime_environment("unspecified")),
        "command_template": "internal-benchmark --worker", "parameters": parameters,
        "execution_status": "completed" if supervised["status"] == "completed" else "stopped", "composer_terminal_state": runs[-1]["terminal_state"],
        "overall_gate": "go", "stop_reasons": [], "rss_protocol": FROZEN_RSS_PROTOCOL,
        "thresholds": FROZEN_THRESHOLDS,
        "output_json_bytes_basis": config["output_json_bytes_basis"],
        "input_json_bytes_basis": config["input_json_bytes_basis"], "runs": runs,
        "summary": None, "ratio_evidence": {name: {"status": "not_applicable"} for name in ("wall", "rss", "output_json")},
        "reference_budget": {"elapsed_hours": 0.0, "limit_hours": config["total_reference_budget_hours"]},
    }
    result["stop_reasons"] = derive_stop_reasons(result)
    result["overall_gate"] = "stop" if result["stop_reasons"] else "go"
    result["result_digest"] = recompute_result_digest(result)
    return result


def _summary_from_runs(runs: list[dict[str, Any]]) -> dict[str, Any]:
    measured = [run for run in runs if run["run_kind"] == "performance_measured"]
    if not measured:
        raise BenchmarkRunnerError("performance result lacks measured runs")
    def values(path: tuple[str, ...]) -> list[Any]:
        out = []
        for run in measured:
            item: Any = run
            for part in path:
                item = item[part]
            out.append(item)
        return out
    return {
        "median_wall_seconds": median(values(("timings", "end_to_end", "wall_seconds"))),
        "max_wall_seconds": max(values(("timings", "end_to_end", "wall_seconds"))),
        "median_peak_rss_mib": median(values(("rss", "peak_rss_mib"))),
        "max_peak_rss_mib": max(values(("rss", "peak_rss_mib"))),
        "median_output_json_bytes": int(median(values(("output_json_bytes",)))),
        "max_output_json_bytes": max(values(("output_json_bytes",))),
    }


def atomic_write_result(directory: Path, result: dict[str, Any]) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    key = logical_key(result["parameters"])
    path = directory / (key + ".json")
    if path.exists():
        raise BenchmarkRunnerError("refusing to overwrite evidence")
    temporary = directory / (key + ".tmp")
    if temporary.exists():
        raise BenchmarkRunnerError("refusing to replace interrupted temporary evidence")
    temporary.write_bytes(canonical_json_bytes(result))
    os.replace(temporary, path)
    return path


def load_cached_result(path: Path, config: dict[str, Any], envelope: dict[str, Any], *, subject_digest: str | None = None) -> dict[str, Any]:
    if path.suffix != ".json" or path.with_suffix(".tmp").exists():
        raise BenchmarkRunnerError("invalid cache path")
    try:
        result = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BenchmarkRunnerError("corrupt evidence") from exc
    validate_benchmark_result_context(result, config, envelope)
    if subject_digest is not None and result.get("benchmark_subject_digest") != subject_digest:
        raise BenchmarkRunnerError("foreign evidence subject")
    if path.stem != logical_key(result["parameters"]):
        raise BenchmarkRunnerError("cache key mismatch")
    return result


def scan_cache(directory: Path, config: dict[str, Any], envelope: dict[str, Any], *, subject_digest: str) -> dict[str, dict[str, Any]]:
    """Load only final JSON evidence; duplicate logical keys fail closed."""
    cached: dict[str, dict[str, Any]] = {}
    if any(directory.glob("*.tmp")):
        raise BenchmarkRunnerError("interrupted temporary evidence present")
    for path in sorted(directory.glob("*.json")):
        result = load_cached_result(path, config, envelope, subject_digest=subject_digest)
        key = logical_key(result["parameters"])
        if key in cached:
            raise BenchmarkRunnerError("duplicate logical cache key")
        cached[key] = result
    expected = [logical_key(item) for item in _campaign_requests(config)]
    if set(cached) != set(expected[:len(cached)]):
        raise BenchmarkRunnerError("cache is not a canonical campaign prefix")
    elapsed = [Decimal(str(cached[key]["reference_budget"]["elapsed_hours"])) for key in expected[:len(cached)]]
    if any(later <= earlier for earlier, later in zip(elapsed, elapsed[1:])):
        raise BenchmarkRunnerError("cache budget elapsed hours are not strictly increasing")
    return cached


def close_campaign(results: Iterable[dict[str, Any]], config: dict[str, Any], envelope: dict[str, Any]) -> dict[str, Any]:
    """The sole C2B delegation point for a suite GO or STOP decision."""
    return validate_complete_benchmark_suite(results, config, envelope)


def _parameters(kind: str, scale_id: str, scenario_id: str, seed: int, permutation_seed: int | None = None) -> dict[str, Any]:
    return {"measurement_kind": kind, "scale_id": scale_id, "scenario_id": scenario_id, "generation_seed": seed, "permutation_seed": permutation_seed}


def _ratio_evidence(one: dict[str, Any], two: dict[str, Any]) -> None:
    """Populate only the 2x side from actual measured performance summaries."""
    fields = {"wall": "median_wall_seconds", "rss": "median_peak_rss_mib", "output_json": "median_output_json_bytes"}
    for name, field in fields.items():
        baseline = Decimal(str(one["summary"][field]))
        observed = Decimal(str(two["summary"][field]))
        if baseline <= 0:
            raise BenchmarkRunnerError("zero ratio baseline")
        ratio = (observed / baseline).quantize(Decimal("0.000001"), rounding=ROUND_HALF_EVEN)
        two["ratio_evidence"][name] = {"status": "measured", "baseline_1x": float(baseline), "observed_2x": float(observed), "ratio": float(ratio)}
    two["stop_reasons"] = derive_stop_reasons(two)
    two["overall_gate"] = "stop" if two["stop_reasons"] else "go"
    two["result_digest"] = recompute_result_digest(two)


def _actual_observations(
    request: dict[str, Any], count: int, *, timeout_seconds: float, supervisor: Any, clock: Any,
    before_call: Any | None = None,
) -> list[tuple[dict[str, Any], float]]:
    observations = []
    for _ in range(count):
        if before_call is not None:
            before_call()
        started = clock()
        observation = supervisor(request, timeout_seconds=timeout_seconds)
        observations.append((observation, max(0.0, clock() - started)))
        if observation.get("status") != "completed":
            break
    return observations


def _runtime_environment(os_build_class: str) -> dict[str, Any]:
    """Collect only the bounded, non-identifying C2A environment envelope."""
    if os_build_class not in {"public_ci", "frozen_reference", "unspecified"}:
        raise BenchmarkRunnerError("unknown OS build class")
    family_by_platform = {"win32": "windows", "linux": "linux", "darwin": "macos"}
    family = family_by_platform.get(sys.platform)
    if family is None:
        raise BenchmarkRunnerError("unsupported OS family")
    machine = platform.machine().lower().replace("-", "_")
    if not machine:
        raise BenchmarkRunnerError("CPU architecture unavailable")
    architecture = "x86_64" if machine in {"amd64", "x86_64", "x64"} else "arm64" if machine in {"arm64", "aarch64"} else "other"
    cpus = os.cpu_count()
    if not isinstance(cpus, int) or cpus < 1:
        raise BenchmarkRunnerError("logical CPU count unavailable")
    try:
        if family == "windows":
            import ctypes
            class MemoryStatus(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]
            status = MemoryStatus(); status.dwLength = ctypes.sizeof(status)
            if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                raise OSError("GlobalMemoryStatusEx")
            total_memory = int(status.ullTotalPhys)
        else:
            total_memory = int(os.sysconf("SC_PAGE_SIZE")) * int(os.sysconf("SC_PHYS_PAGES"))
    except (AttributeError, OSError, ValueError):
        raise BenchmarkRunnerError("RAM capacity unavailable") from None
    if total_memory <= 0:
        raise BenchmarkRunnerError("RAM capacity unavailable")
    gib = Decimal(total_memory) / Decimal(1024 ** 3)
    tier = "le_8gib" if gib <= 8 else "8_to_16gib" if gib <= 16 else "16_to_32gib" if gib <= 32 else "32_to_64gib" if gib <= 64 else "gt_64gib"
    return {"os_family": family, "os_build_class": os_build_class,
            "python_version": "%d.%d.%d" % sys.version_info[:3],
            "cpu_architecture": architecture, "logical_cpu_count": cpus, "ram_tier": tier}


def _campaign_requests(config: dict[str, Any]) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = []
    for item in config["matrices"]["coverage_cells"]:
        requests.append(_parameters("coverage", item["scale_id"], item["scenario_id"], config["generation"]["generation_seed"]))
    for item in config["matrices"]["performance_cells"]:
        requests.append(_parameters("performance", item["scale_id"], item["scenario_id"], config["generation"]["generation_seed"]))
    for item in config["matrices"]["determinism_cells"]:
        requests.append(_parameters("determinism", item["scale_id"], item["scenario_id"], config["generation"]["generation_seed"], item["permutation_seed"]))
    kind_order = {"coverage": 0, "performance": 1, "determinism": 2}
    return sorted(requests, key=lambda item: (kind_order[item["measurement_kind"]], SCALES[item["scale_id"]], item["scenario_id"], item["permutation_seed"] if item["permutation_seed"] is not None else -1))


def _result_from_observations(
    *, config: dict[str, Any], subject_commit: str, subject_manifest: list[dict[str, str]], parameters: dict[str, Any],
    observations: list[tuple[dict[str, Any], float]], environment: dict[str, Any] | None = None, determinism: tuple[str, str] = ("not_applicable", "not_applicable"),
) -> dict[str, Any]:
    first, elapsed = observations[0]
    result = build_result(cell=None, supervised=first, config=config, subject_commit=subject_commit, subject_manifest=subject_manifest, parameters=parameters, elapsed_seconds=elapsed, environment=environment, determinism=determinism)
    if parameters["measurement_kind"] == "performance" and first["status"] in {"completed", "rss_unavailable"}:
        if all(item["status"] == "completed" for item, _ in observations) and len(observations) != 4:
            raise BenchmarkRunnerError("performance requires warmup plus three measured observations")
        scenario = parameters["scenario_id"]
        result["runs"] = [
            _run_record(ordinal=index, measurement_kind="performance", scenario_id=scenario, supervised=supervised,
                        worker=supervised.get("worker"), elapsed_seconds=duration,
                        input_json_bytes=supervised.get("worker", {}).get("input_json_bytes"),
                        output_json_bytes=supervised.get("worker", {}).get("output_json_bytes"))
            for index, (supervised, duration) in enumerate(observations, 1)
        ]
        interrupted = next((index for index, item in enumerate(observations) if item[0]["status"] in {"timeout", "process_crash"}), None)
        if interrupted is not None:
            result["execution_status"] = "stopped"
            result["runs"] = result["runs"][:interrupted + 1]
            result["summary"] = None
        elif any(run["rss"]["status"] != "available" for run in result["runs"]):
            result["execution_status"] = "stopped"
            result["summary"] = None
        else:
            result["summary"] = _summary_from_runs(result["runs"])
    result["composer_terminal_state"] = result["runs"][-1]["terminal_state"]
    result["stop_reasons"] = derive_stop_reasons(result)
    result["overall_gate"] = "stop" if result["stop_reasons"] else "go"
    result["result_digest"] = recompute_result_digest(result)
    return result


def run_benchmark_campaign(
    config: dict[str, Any], envelope: dict[str, Any], *, benchmark_subject_commit: str, cache_directory: Path,
    timeout_seconds: float, os_build_class: str, supervisor: Any | None = None, clock: Any = time.monotonic,
) -> dict[str, Any]:
    """Internal serial orchestration; only complete-suite validation can close it.

    This function is intentionally not wired to a runtime entry point.  C2C is
    responsible for a formal/reference invocation and evidence publication.
    """
    validate_campaign_inputs(config, envelope)
    environment = _runtime_environment(os_build_class)
    supervisor = supervisor or supervise_worker
    manifest = build_subject_manifest(benchmark_subject_commit)
    subject_digest = recompute_subject_digest(manifest)
    cached = scan_cache(cache_directory, config, envelope, subject_digest=subject_digest) if cache_directory.exists() else {}
    requests = _campaign_requests(config)
    elapsed_offset = Decimal("0") if not cached else Decimal(str(cached[logical_key(requests[len(cached) - 1])]["reference_budget"]["elapsed_hours"])) * Decimal("3600")
    started = clock() - float(elapsed_offset)
    if any(item.get("environment") != environment for item in cached.values()):
        raise BenchmarkRunnerError("cached evidence environment differs from campaign")
    results: dict[str, dict[str, Any]] = dict(cached)
    baselines: dict[tuple[str, str], dict[str, Any]] = {}
    def require_budget_before_child() -> None:
        if clock() - started >= config["total_reference_budget_hours"] * 3600:
            raise _CampaignBudgetStopped()

    for parameters in requests:
        key = logical_key(parameters)
        if key in results:
            continue
        if clock() - started >= config["total_reference_budget_hours"] * 3600:
            # C2A has no serializable "not started due to campaign budget"
            # result.  Keep already validated evidence and fail closed without
            # manufacturing timeout/crash records or a complete suite.
            return {"status": "incomplete_budget_exceeded", "overall_gate": "stop", "results": [results[name] for name in sorted(results)]}
        request = {"config": config, "scale_id": parameters["scale_id"], "scenario_id": parameters["scenario_id"], "generation_seed": parameters["generation_seed"], "permutation_seed": parameters["permutation_seed"], "measurement_kind": parameters["measurement_kind"]}
        count = 4 if parameters["measurement_kind"] == "performance" else 1
        try:
            observations = _actual_observations(request, count, timeout_seconds=timeout_seconds, supervisor=supervisor, clock=clock, before_call=require_budget_before_child)
        except _CampaignBudgetStopped:
            return {"status": "incomplete_budget_exceeded", "overall_gate": "stop", "results": [results[name] for name in sorted(results)]}
        # C2A has no legal run shape for a missing worker/contract-error RSS
        # failure.  Preserve prior checkpoints and stop before fabricating one.
        for observed, _ in observations:
            status, worker = observed.get("status"), observed.get("worker")
            if status == "contract_error" or (status == "rss_unavailable" and worker is None):
                return {"status": "incomplete_" + status, "overall_gate": "stop", "results": [results[name] for name in sorted(results)]}
        deterministic = ("not_applicable", "not_applicable")
        if parameters["measurement_kind"] == "determinism" and observations[0][0].get("status") in {"completed", "rss_unavailable"} and observations[0][0].get("worker") is not None:
            baseline_request = dict(request); baseline_request["permutation_seed"] = None
            baseline_key = (parameters["scale_id"], parameters["scenario_id"])
            baseline = baselines.get(baseline_key)
            if baseline is None:
                try:
                    baseline = _actual_observations(baseline_request, 1, timeout_seconds=timeout_seconds, supervisor=supervisor, clock=clock, before_call=require_budget_before_child)[0][0]
                except _CampaignBudgetStopped:
                    return {"status": "incomplete_budget_exceeded", "overall_gate": "stop", "results": [results[name] for name in sorted(results)]}
                baselines[baseline_key] = baseline
            if baseline.get("status") != "completed":
                return {"status": "incomplete_baseline_failed", "overall_gate": "stop", "results": [results[name] for name in sorted(results)]}
            if baseline.get("status") == "completed":
                matched = baseline["worker"]["canonical_output_digest"] == observations[0][0]["worker"]["canonical_output_digest"]
                fingerprint_matched = (baseline["worker"]["report_fingerprint"], baseline["worker"]["final_fingerprint"]) == (observations[0][0]["worker"]["report_fingerprint"], observations[0][0]["worker"]["final_fingerprint"])
                deterministic = ("matched" if matched else "mismatched", "matched" if fingerprint_matched else "mismatched")
        result = _result_from_observations(config=config, subject_commit=benchmark_subject_commit, subject_manifest=manifest, parameters=parameters, observations=observations, environment=environment, determinism=deterministic)
        result["reference_budget"]["elapsed_hours"] = float((Decimal(str(clock() - started)) / Decimal("3600")).quantize(Decimal("0.000001"), rounding=ROUND_HALF_EVEN))
        result["stop_reasons"] = derive_stop_reasons(result)
        result["overall_gate"] = "stop" if result["stop_reasons"] else "go"
        result["result_digest"] = recompute_result_digest(result)
        if parameters["measurement_kind"] == "performance" and parameters["scale_id"] == "2.0x":
            one = results.get(logical_key(_parameters("performance", "1.0x", parameters["scenario_id"], config["generation"]["generation_seed"])))
            if one is not None and one["execution_status"] == result["execution_status"] == "completed":
                _ratio_evidence(one, result)
        validate_benchmark_result_context(result, config, envelope)
        atomic_write_result(cache_directory, result)
        results[key] = result
    ordered = [results[logical_key(parameters)] for parameters in requests]
    return close_campaign(ordered, config, envelope)


def worker_payload(request: dict[str, Any]) -> dict[str, Any]:
    """One logical cell only; parent-side supervision owns timeout/RSS/evidence."""
    try:
        composed = compose_cell(request["config"], scale_id=request["scale_id"], scenario_id=request["scenario_id"], generation_seed=request["generation_seed"], permutation_seed=request.get("permutation_seed"))
        metrics = composed["metrics"]
        report = composed["report"]
        payload = {"report": report, "final_profile": composed["final_profile"]}
        serialization_at = time.perf_counter()
        canonical_payload = canonical_json_bytes(payload)
        stage_seconds = dict(composed["stage_seconds"])
        stage_seconds["canonical_serialization"] = time.perf_counter() - serialization_at
        return {"status": "ok", "proposal_status": composed["status"], "final_present": composed["final_profile"] is not None, "metrics": metrics.__dict__, "report_fingerprint": report["semantic_fingerprint"], "report_counts": {"candidate_groups": len(report["candidate_groups"]), "scope_partitions": len(report["scope_partitions"]), "conflicts": len(report["approval_required_conflicts"]), "proposals": len(report["proposed_resolutions"]), "blockers": len(report["unresolvable_blockers"])}, "final_fingerprint": None if composed["final_profile"] is None else composed["final_profile"]["semantic_fingerprint"], "input_json_bytes": len(canonical_json_bytes(composed["assets"])), "output_json_bytes": len(canonical_payload), "canonical_output_digest": _sha(canonical_payload), "stage_seconds": stage_seconds}
    except Exception as exc:  # Worker errors become parent-side contract evidence.
        return {"status": "error", "error": type(exc).__name__}


def _round_mib(value: int | float) -> float:
    return float((Decimal(str(value)) / Decimal("1048576")).quantize(
        Decimal("0.001"), rounding=ROUND_HALF_EVEN
    ))


def _reported_rss_delta(baseline_bytes: int | float, peak_bytes: int | float) -> float:
    """C2A subtracts reported values, rather than raw bytes, before rounding."""
    return float((Decimal(str(_round_mib(peak_bytes))) - Decimal(str(_round_mib(baseline_bytes)))).quantize(
        Decimal("0.001"), rounding=ROUND_HALF_EVEN
    ))


def _windows_peak_working_set(pid: int) -> int:
    """Return the child-only PeakWorkingSetSize without a third-party package."""
    if os.name != "nt":
        raise OSError("Windows RSS sampler unavailable")
    import ctypes
    from ctypes import wintypes

    class Counters(ctypes.Structure):
        _fields_ = [
            ("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD),
            ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t), ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    process = kernel32.OpenProcess(0x0400 | 0x0010, False, pid)
    if not process:
        raise OSError(ctypes.get_last_error(), "OpenProcess")
    try:
        counters = Counters()
        counters.cb = ctypes.sizeof(Counters)
        if not psapi.GetProcessMemoryInfo(process, ctypes.byref(counters), counters.cb):
            raise OSError(ctypes.get_last_error(), "GetProcessMemoryInfo")
        return int(counters.PeakWorkingSetSize)
    finally:
        kernel32.CloseHandle(process)


def _reader(stream: Any, output: queue.Queue[bytes]) -> None:
    try:
        for line in iter(stream.readline, b""):
            output.put(line)
    finally:
        output.put(b"")


def supervise_worker(
    request: dict[str, Any],
    *,
    timeout_seconds: float,
    python_executable: str | None = None,
    process_factory: Any = subprocess.Popen,
    sampler: Any | None = None,
    clock: Any = time.monotonic,
    sleeper: Any = time.sleep,
) -> dict[str, Any]:
    """Run one child with timeout precedence and fail-closed RSS evidence."""
    command = [python_executable or sys.executable, str(Path(__file__).resolve()), "--worker"]
    # stderr is deliberately not evidence and must not back-pressure a child.
    child = process_factory(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    output: queue.Queue[bytes] = queue.Queue()
    thread = threading.Thread(target=_reader, args=(child.stdout, output), daemon=True)
    thread.start()
    started = clock()
    deadline = started + timeout_seconds

    def cleanup() -> None:
        if child.poll() is None:
            child.kill()
            child.wait()
        for stream in (child.stdin, child.stdout, child.stderr):
            close = getattr(stream, "close", None)
            if callable(close):
                try:
                    close()
                except OSError:
                    # Cleanup must not replace the timeout/crash result.
                    pass
        thread.join(timeout=0.1)

    try:
        first = output.get(timeout=max(0.0, deadline - clock()))
    except queue.Empty:
        cleanup()
        return {"status": "timeout"}
    if first != b"READY\n":
        cleanup()
        return {"status": "process_crash"}
    try:
        # READY is emitted after imports and before worker input is read.  Take
        # the baseline first so generation cannot enter the RSS interval.
        read_rss = sampler or _windows_peak_working_set
        rss_available = os.name == "nt" or sampler is not None
        try:
            baseline = read_rss(child.pid) if rss_available else None
        except OSError:
            return {"status": "rss_unavailable"}
        if child.stdin is None:
            raise OSError("worker stdin unavailable")
        child.stdin.write(canonical_json_bytes(request))
        child.stdin.close()
        peak = baseline
        response_line: bytes | None = None
        while child.poll() is None:
            if clock() >= deadline:
                child.kill()
                child.wait()
                return {"status": "timeout"}
            if rss_available:
                try:
                    peak = max(peak or 0, read_rss(child.pid))
                except OSError:
                    # Keep reading the worker response.  A valid worker with
                    # missing RSS is a representable stopped C2A result.
                    rss_available = False
            try:
                response_line = output.get_nowait()
                if response_line:
                    break
            except queue.Empty:
                sleeper(min(0.05, max(0.0, deadline - clock())))
        while child.poll() is None:
            if clock() >= deadline:
                child.kill()
                child.wait()
                return {"status": "timeout"}
            sleeper(min(0.05, max(0.0, deadline - clock())))
        while response_line is None:
            try:
                line = output.get_nowait()
            except queue.Empty:
                break
            if line:
                response_line = line
        if child.returncode:
            return {"status": "process_crash"}
        if response_line is None:
            return {"status": "process_crash"}
        try:
            response = json.loads(response_line.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return {"status": "process_crash"}
        if response.get("status") != "ok":
            return {"status": "contract_error"}
        if not rss_available:
            return {"status": "rss_unavailable", "worker": response}
        return {
            "status": "completed", "worker": response,
            "rss": {
                "status": "available", "baseline_rss_mib": _round_mib(baseline or 0),
                "peak_rss_mib": _round_mib(peak),
                "delta_peak_rss_mib": _reported_rss_delta(baseline or 0, peak),
            },
        }
    finally:
        cleanup()


def validate_campaign_inputs(config: dict[str, Any], envelope: dict[str, Any]) -> dict[str, Any]:
    """The C2B boundary: C2A's schema-first context must precede generation."""
    return validate_benchmark_campaign_context(config, envelope)


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="internal C2B benchmark runner")
    parser.add_argument("--worker", action="store_true")
    args = parser.parse_args(argv)
    if not args.worker:
        raise BenchmarkRunnerError("C2B accepts only structured internal worker requests")
    try:
        sys.stdout.buffer.write(b"READY\n")
        sys.stdout.buffer.flush()
        request = json.loads(sys.stdin.buffer.read().decode("utf-8"))
        sys.stdout.buffer.write(canonical_json_bytes(worker_payload(request)) + b"\n")
        return 0
    except (UnicodeDecodeError, json.JSONDecodeError, BenchmarkRunnerError):
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
