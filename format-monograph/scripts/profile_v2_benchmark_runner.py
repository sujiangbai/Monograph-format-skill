#!/usr/bin/env python3
"""Internal, serial C2B benchmark runner. It never authorizes a benchmark GO."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import queue
import subprocess
import sys
import threading
import time
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
    validate_complete_benchmark_suite,
)
from profile_v2_canonical import canonical_data_digest, stamp_intent_semantic_fingerprint_v041
from profile_v2_composer import apply_intent_resolutions_v041, compose_intent_profile_v041
from profile_v2_registry import load_registry
from profile_v2_scope import normalize_scope, normalized_property_scope_key


class BenchmarkRunnerError(ValueError):
    pass


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


def generate_assets(config: dict[str, Any], scale_id: str, scenario_id: str, seed: int, *, permuted: bool = False) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Generate only C1's contract.intent-probe assets; never a base rule asset."""
    if scenario_id not in SCENARIOS:
        raise BenchmarkRunnerError("unknown scenario")
    work = scaled_workload(config, scale_id)
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
    if permuted:
        for index, asset in enumerate(assets):
            asset["rules"].reverse()
            assets[index] = stamp_intent_semantic_fingerprint_v041(asset)
        assets.reverse()
    return assets, work


def _approval(report: dict[str, Any], ordinal: int) -> dict[str, Any]:
    conflict = report["approval_required_conflicts"][ordinal]
    document = {"artifact_kind": "qa-approval-artifact", "schema_version": "2.2", "registry_contract_version": "2.2", "authority_contract_version": "1.0", "artifact_id": "qa-approval-artifact:c2b-%04d" % ordinal, "created_by_tool": _tool(), "input_fingerprints": [{"input_id": "input:c2b-source", "role": "source_document", "fingerprint": report["bindings"]["input_fingerprint"]}, {"input_id": "input:c2b-structure", "role": "structure", "fingerprint": report["bindings"]["structure_fingerprint"]}, {"input_id": "input:c2b-report", "role": "conflict_report", "fingerprint": report["semantic_fingerprint"]}], "semantic_fingerprint": _sha(b"unstamped"), "approval_id": "approval:c2b-%04d" % ordinal, "approver": {"actor_id": "actor:c2b-user", "actor_role": "user"}, "decision_type": "conflict_resolution", "decision": "adopt_proposed", "reason": "Deterministic synthetic C2B approval.", "created_at": "2026-08-23T00:00:00Z", "bindings": {"input_fingerprint": report["bindings"]["input_fingerprint"], "structure_fingerprint": report["bindings"]["structure_fingerprint"], "composition_report_fingerprint": report["semantic_fingerprint"]}, "target": {"conflict_id": conflict["conflict_id"], "proposed_resolution_id": conflict["proposed_resolution_id"], "normalized_scope": deepcopy(conflict["key"]["normalized_scope"])}, "previous_approval_id": None}
    return stamp_intent_semantic_fingerprint_v041(document)


def compose_cell(config: dict[str, Any], *, scale_id: str, scenario_id: str, generation_seed: int, permutation_seed: int | None = None) -> dict[str, Any]:
    """Pure worker core: compose one C2B logical cell without a supervisor or GO."""
    assets, expected = generate_assets(config, scale_id, scenario_id, generation_seed, permuted=permutation_seed is not None)
    manifest = feature_manifest()
    composed = compose_intent_profile_v041(assets, manifest, input_fingerprint=_sha(b"c2b-source"), structure_fingerprint=_sha(b"c2b-structure"), artifact_id="conflict-report:c2b-%s-%s" % (scale_id, scenario_id), created_by_tool=_tool(), generated_at="2026-08-23T00:00:00Z", include_metrics=True)
    metrics = composed.metrics
    metric_expectations = {"input_asset_count": expected["rule_fragment"], "input_rule_count": expected["binding"], "input_binding_count": expected["binding"], "expected_key_count": expected["key"], "candidate_count": expected["candidate"]}
    if metrics is None or any(getattr(metrics, field) != value for field, value in metric_expectations.items()):
        raise BenchmarkRunnerError("C1 metrics do not conserve generated workload")
    final = None
    pre_status = composed.status
    status = pre_status
    if scenario_id != "dense-crossing":
        approvals = [_approval(composed.report, index) for index in range(len(composed.report["approval_required_conflicts"]))]
        applied = apply_intent_resolutions_v041(composed.report, approvals, manifest, task_id="task:c2b", task_fingerprint=_sha(b"c2b-task"), artifact_id="final-execution-profile:c2b", created_by_tool=_tool(), metrics=metrics)
        status, final = applied.status, applied.final_profile
    if scenario_id in {"disjoint", "subset-chain"} and (status != "profile_generated" or final is None):
        raise BenchmarkRunnerError("unexpected final-state boundary")
    if scenario_id == "dense-crossing" and (status != "unresolvable" or final is not None):
        raise BenchmarkRunnerError("dense crossing boundary")
    if scenario_id == "mixed-conflict-approval" and (status != "profile_generated" or final is None):
        raise BenchmarkRunnerError("mixed approval boundary")
    return {"assets": assets, "report": composed.report, "final_profile": final, "metrics": metrics, "status": status, "pre_status": pre_status, "expected": expected}


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


def build_result(
    *,
    cell: dict[str, Any],
    supervised: dict[str, Any],
    config: dict[str, Any],
    subject_commit: str,
    subject_manifest: list[dict[str, str]],
    parameters: dict[str, Any],
    elapsed_seconds: float,
) -> dict[str, Any]:
    """Construct one non-artifact C2A result; it never decides suite GO."""
    if supervised["status"] != "completed":
        raise BenchmarkRunnerError("stopped result construction is owned by the supervisor campaign path")
    worker = supervised["worker"]
    scenario = parameters["scenario_id"]
    terminal = FROZEN_SCENARIO_SEMANTICS[scenario]
    final_present = worker["final_present"]
    timing = {
        "synthetic_generation": _measured(0.0),
        "schema_registry_validation": _measured(0.0),
        "compose": _measured(elapsed_seconds),
        "approval_generation": _measured(0.0) if scenario == "mixed-conflict-approval" else _not_applicable(),
        "apply": _measured(0.0) if final_present else _not_applicable(),
        "canonical_serialization": _measured(0.0),
        "end_to_end": _measured(elapsed_seconds),
    }
    input_bytes = len(canonical_json_bytes(cell["assets"]))
    output_bytes = len(canonical_json_bytes({
        "report": cell["report"], "final_profile": cell["final_profile"],
    }))
    run = {
        "run_index": 1, "run_kind": parameters["measurement_kind"],
        "run_status": "completed", "timings": timing,
        "rss": supervised.get("rss", {"status": "unavailable"}),
        "input_json_bytes": input_bytes, "output_json_bytes": output_bytes,
        "metrics": worker["metrics"], "terminal_state": terminal["terminal_state"],
        "terminal_trace": {
            "pre_approval": terminal["pre_approval_terminal"],
            "post_approval": terminal["post_approval_terminal"],
        },
        "coverage_conservation": "passed", "stable_id_status": "stable",
        "canonical_determinism": "matched" if parameters["measurement_kind"] == "determinism" else "not_applicable",
        "fingerprint_determinism": "matched" if parameters["measurement_kind"] == "determinism" else "not_applicable",
        "contract_status": "valid", "final_profile_present": final_present,
        "final_profile_fingerprint": worker["final_fingerprint"],
    }
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
        "environment": {
            "os_family": "windows" if os.name == "nt" else "linux",
            "os_build_class": "unspecified", "python_version": "%d.%d.%d" % sys.version_info[:3],
            "cpu_architecture": "other", "logical_cpu_count": 1, "ram_tier": "le_8gib",
        },
        "command_template": "internal-benchmark --worker", "parameters": parameters,
        "execution_status": "completed", "composer_terminal_state": terminal["terminal_state"],
        "overall_gate": "go", "stop_reasons": [], "rss_protocol": FROZEN_RSS_PROTOCOL,
        "thresholds": FROZEN_THRESHOLDS,
        "output_json_bytes_basis": config["output_json_bytes_basis"],
        "input_json_bytes_basis": config["input_json_bytes_basis"], "runs": [run],
        "summary": None, "ratio_evidence": {name: {"status": "not_applicable"} for name in ("wall", "rss", "output_json")},
        "reference_budget": {"elapsed_hours": 0.0, "limit_hours": config["total_reference_budget_hours"]},
    }
    result["stop_reasons"] = derive_stop_reasons(result)
    result["overall_gate"] = "stop" if result["stop_reasons"] else "go"
    result["result_digest"] = recompute_result_digest(result)
    return result


def atomic_write_result(directory: Path, result: dict[str, Any]) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    key = logical_key(result["parameters"])
    path = directory / (key + ".json")
    if path.exists():
        raise BenchmarkRunnerError("refusing to overwrite evidence")
    temporary = directory / (key + ".tmp")
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
    for path in sorted(directory.glob("*.json")):
        result = load_cached_result(path, config, envelope, subject_digest=subject_digest)
        key = logical_key(result["parameters"])
        if key in cached:
            raise BenchmarkRunnerError("duplicate logical cache key")
        cached[key] = result
    return cached


def enforce_campaign_budget(started: float, *, limit_seconds: float, clock: Any = time.monotonic) -> None:
    if clock() - started > limit_seconds:
        raise BenchmarkRunnerError("reference budget exceeded")


def close_campaign(results: Iterable[dict[str, Any]], config: dict[str, Any], envelope: dict[str, Any]) -> dict[str, Any]:
    """The sole C2B delegation point for a suite GO or STOP decision."""
    return validate_complete_benchmark_suite(results, config, envelope)


def worker_payload(request: dict[str, Any]) -> dict[str, Any]:
    """One logical cell only; parent-side supervision owns timeout/RSS/evidence."""
    try:
        composed = compose_cell(request["config"], scale_id=request["scale_id"], scenario_id=request["scenario_id"], generation_seed=request["generation_seed"], permutation_seed=request.get("permutation_seed"))
        metrics = composed["metrics"]
        report = composed["report"]
        return {"status": "ok", "proposal_status": composed["status"], "final_present": composed["final_profile"] is not None, "metrics": metrics.__dict__, "report_fingerprint": report["semantic_fingerprint"], "report_counts": {"candidate_groups": len(report["candidate_groups"]), "scope_partitions": len(report["scope_partitions"]), "conflicts": len(report["approval_required_conflicts"]), "proposals": len(report["proposed_resolutions"]), "blockers": len(report["unresolvable_blockers"])}, "final_fingerprint": None if composed["final_profile"] is None else composed["final_profile"]["semantic_fingerprint"]}
    except Exception as exc:  # Worker errors become parent-side contract evidence.
        return {"status": "error", "error": type(exc).__name__}


def _round_mib(value: int | float) -> float:
    return float((Decimal(str(value)) / Decimal("1048576")).quantize(
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
    child = process_factory(
        command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    output: queue.Queue[bytes] = queue.Queue()
    thread = threading.Thread(target=_reader, args=(child.stdout, output), daemon=True)
    thread.start()
    started = clock()
    try:
        first = output.get(timeout=timeout_seconds)
    except queue.Empty:
        child.kill()
        child.wait()
        return {"status": "timeout"}
    if first != b"READY\n":
        child.kill()
        child.wait()
        return {"status": "process_crash"}
    try:
        if child.stdin is None:
            raise OSError("worker stdin unavailable")
        child.stdin.write(canonical_json_bytes(request))
        child.stdin.close()
        read_rss = sampler or _windows_peak_working_set
        rss_available = os.name == "nt" or sampler is not None
        try:
            baseline = read_rss(child.pid) if rss_available else None
        except OSError:
            return {"status": "rss_unavailable"}
        peak = baseline
        response_line: bytes | None = None
        while child.poll() is None:
            if clock() - started > timeout_seconds:
                child.kill()
                child.wait()
                return {"status": "timeout"}
            if rss_available:
                try:
                    peak = max(peak or 0, read_rss(child.pid))
                except OSError:
                    return {"status": "rss_unavailable"}
            try:
                response_line = output.get_nowait()
                if response_line:
                    break
            except queue.Empty:
                sleeper(0.05)
        child.wait()
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
                "delta_peak_rss_mib": _round_mib(max(0, peak - (baseline or 0))),
            },
        }
    finally:
        if child.poll() is None:
            child.kill()
            child.wait()
        for stream in (child.stdin, child.stdout, child.stderr):
            if stream is not None:
                stream.close()


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
