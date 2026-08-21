#!/usr/bin/env python3
"""Pure P3a-C2 benchmark contract helpers. No runner or runtime integration."""
from __future__ import annotations

import copy
import hashlib
import json
import math
import re
from decimal import Decimal, ROUND_HALF_EVEN
from statistics import median
from typing import Any, Iterable, Mapping, Sequence


class BenchmarkContractError(ValueError):
    """Raised when benchmark contract semantics are contradictory."""


SCALES = ("0.5x", "1.0x", "1.5x", "2.0x")
SCENARIOS = (
    "disjoint",
    "subset-chain",
    "dense-crossing",
    "mixed-conflict-approval",
)
DETERMINISM_SCALES = ("1.5x", "2.0x")
PERFORMANCE_CELLS = (
    ("0.5x", "mixed-conflict-approval"),
    ("1.0x", "mixed-conflict-approval"),
    *((scale, scenario) for scale in DETERMINISM_SCALES for scenario in SCENARIOS),
)
TIMING_STAGES = (
    "synthetic_generation",
    "schema_registry_validation",
    "compose",
    "approval_generation",
    "apply",
    "canonical_serialization",
    "end_to_end",
)
C1_METRIC_FIELDS = (
    "input_asset_count",
    "input_rule_count",
    "input_binding_count",
    "expected_key_count",
    "candidate_count",
    "candidate_group_count",
    "partition_count",
    "conflict_count",
    "proposal_count",
    "blocker_count",
    "max_candidates_per_key",
    "max_partition_width",
    "max_repartition_depth",
)
STOP_REASON_ORDER = (
    "threshold_exceeded",
    "wall_ratio_exceeded",
    "rss_ratio_exceeded",
    "output_ratio_exceeded",
    "timeout",
    "process_crash",
    "rss_unavailable",
    "coverage_conservation_failure",
    "stable_id_drift",
    "canonical_nondeterminism",
    "fingerprint_nondeterminism",
    "terminal_state_mismatch",
    "subject_stale",
    "contract_error",
    "reference_budget_exceeded",
)
STOP_REASON_INDEX = {value: index for index, value in enumerate(STOP_REASON_ORDER)}
SHA256_RE = re.compile(r"^sha256:[a-f0-9]{64}$")
COMMIT_RE = re.compile(r"^[a-f0-9]{40}$")
WINDOWS_DRIVE_RE = re.compile(r"^[A-Za-z]:")

FROZEN_SCALE_FACTORS = {
    "0.5x": 0.5,
    "1.0x": 1.0,
    "1.5x": 1.5,
    "2.0x": 2.0,
}
FROZEN_SCENARIO_SEMANTICS = {
    "disjoint": {
        "pre_approval_terminal": "final",
        "post_approval_terminal": "not_applicable",
        "terminal_state": "final",
        "final_requirement": "required",
    },
    "subset-chain": {
        "pre_approval_terminal": "final",
        "post_approval_terminal": "not_applicable",
        "terminal_state": "final",
        "final_requirement": "required",
    },
    "dense-crossing": {
        "pre_approval_terminal": "unresolvable",
        "post_approval_terminal": "not_applicable",
        "terminal_state": "unresolvable",
        "final_requirement": "forbidden",
    },
    "mixed-conflict-approval": {
        "pre_approval_terminal": "awaiting_approval",
        "post_approval_terminal": "final",
        "terminal_state": "final",
        "final_requirement": "conditional",
    },
}
FROZEN_RSS_PROTOCOL = {
    "measurement_method": "external_supervisor_child_peak_working_set",
    "sampling_interval_seconds": 0.05,
    "child_process_scope": "benchmark_child_process_only",
    "record_baseline": True,
    "record_delta": True,
    "unavailable_policy": "fail_closed",
    "delta_rounding": "mib_3_decimal_places_half_even",
}
FROZEN_THRESHOLDS = {
    "scale_limits": [
        {"scale_id": "0.5x", "max_wall_seconds": 60, "max_peak_rss_mib": 512},
        {"scale_id": "1.0x", "max_wall_seconds": 60, "max_peak_rss_mib": 512},
        {"scale_id": "1.5x", "max_wall_seconds": 60, "max_peak_rss_mib": 512},
        {"scale_id": "2.0x", "max_wall_seconds": 120, "max_peak_rss_mib": 1024},
    ],
    "wall_median_ratio_2x_to_1x": 6,
    "rss_ratio_2x_to_1x": 3,
    "output_json_ratio_2x_to_1x": 3,
}
FROZEN_REPETITIONS = {
    "coverage_min_runs_per_cell": 1,
    "performance_warmup_runs_per_cell": 1,
    "performance_measured_runs_per_cell": 3,
    "determinism_runs_per_seed": 1,
}
ABSOLUTE_GATE_RUN_KINDS = {"coverage", "performance_measured", "determinism"}
RATIO_SUMMARY_FIELDS = {
    "wall": "median_wall_seconds",
    "rss": "median_peak_rss_mib",
    "output_json": "median_output_json_bytes",
}
RATIO_THRESHOLD_FIELDS = {
    "wall": "wall_median_ratio_2x_to_1x",
    "rss": "rss_ratio_2x_to_1x",
    "output_json": "output_json_ratio_2x_to_1x",
}
RATIO_STOP_REASONS = {
    "wall": "wall_ratio_exceeded",
    "rss": "rss_ratio_exceeded",
    "output_json": "output_ratio_exceeded",
}


def _finite(value: Any) -> None:
    if isinstance(value, float) and not math.isfinite(value):
        raise BenchmarkContractError("Canonical JSON forbids NaN/Infinity")
    if isinstance(value, Mapping):
        for item in value.values():
            _finite(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            _finite(item)


def canonical_json_bytes(value: Any) -> bytes:
    _finite(value)
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise BenchmarkContractError(f"not canonical-JSON safe: {exc}") from exc


def canonical_sha256(value: Any) -> str:
    return "sha256:" + hashlib.sha256(canonical_json_bytes(value)).hexdigest()


def _self_digest(document: Mapping[str, Any], key: str) -> str:
    payload = copy.deepcopy(dict(document))
    payload.pop(key, None)
    return canonical_sha256(payload)


def recompute_config_digest(document: Mapping[str, Any]) -> str:
    return _self_digest(document, "config_digest")


def recompute_envelope_digest(document: Mapping[str, Any]) -> str:
    return _self_digest(document, "envelope_digest")


def recompute_result_digest(document: Mapping[str, Any]) -> str:
    return _self_digest(document, "result_digest")


def canonical_subject_path(path: str) -> str:
    if not isinstance(path, str) or not path:
        raise BenchmarkContractError("subject path must be a non-empty string")
    if "\\" in path or WINDOWS_DRIVE_RE.match(path) or path.startswith("/"):
        raise BenchmarkContractError("subject path must be POSIX relative")
    if path.endswith("/") or "//" in path:
        raise BenchmarkContractError("subject path must use canonical separators")
    parts = path.split("/")
    if any(part in {"", ".", ".."} for part in parts):
        raise BenchmarkContractError("subject path must be lexical-canonical")
    return "/".join(parts)


def recompute_subject_digest(manifest: Iterable[Mapping[str, Any]]) -> str:
    files = []
    paths = []
    for entry in manifest:
        path = canonical_subject_path(entry.get("path"))
        digest = entry.get("sha256")
        if not isinstance(digest, str) or not SHA256_RE.fullmatch(digest):
            raise BenchmarkContractError("invalid subject digest")
        paths.append(path)
        files.append({"path": path, "sha256": digest})
    if len(paths) != len(set(paths)):
        raise BenchmarkContractError("subject paths must be unique after canonicalization")
    if paths != sorted(paths):
        raise BenchmarkContractError("subject paths must be sorted")
    return canonical_sha256({"basis": "canonical_subject_manifest_v1", "files": files})


def expected_coverage_cells() -> set[tuple[str, str]]:
    return {(scale, scenario) for scale in SCALES for scenario in SCENARIOS}


def expected_performance_cells() -> set[tuple[str, str]]:
    return set(PERFORMANCE_CELLS)


def expected_determinism_cells(seeds: Iterable[int]) -> set[tuple[str, str, int]]:
    return {
        (scale, scenario, int(seed))
        for scale in DETERMINISM_SCALES
        for scenario in SCENARIOS
        for seed in seeds
    }


def _exact(actual: Sequence[Any], expected: set[Any], label: str) -> None:
    if len(actual) != len(set(actual)) or set(actual) != expected:
        raise BenchmarkContractError(f"{label} differs from frozen matrix")


def _config_scale_map(document: Mapping[str, Any]) -> dict[str, Any]:
    return {entry["scale_id"]: entry for entry in document["scales"]}


def _config_scenario_map(document: Mapping[str, Any]) -> dict[str, Any]:
    return {entry["scenario_id"]: entry for entry in document["scenarios"]}


def validate_benchmark_config_semantics(document: Mapping[str, Any]) -> None:
    seeds = list(document["generation"]["permutation_seeds"])
    if len(seeds) < 5 or len(seeds) != len(set(seeds)):
        raise BenchmarkContractError("five unique permutation seeds required")

    scales = _config_scale_map(document)
    if set(scales) != set(SCALES) or len(document["scales"]) != len(SCALES):
        raise BenchmarkContractError("frozen scale inventory mismatch")
    for scale_id, expected_factor in FROZEN_SCALE_FACTORS.items():
        entry = scales[scale_id]
        if entry["factor"] != expected_factor:
            raise BenchmarkContractError("frozen scale factor mismatch")
        if set(entry["scenario_ids"]) != set(SCENARIOS) or len(entry["scenario_ids"]) != len(SCENARIOS):
            raise BenchmarkContractError("coverage scenario set mismatch")

    scenarios = _config_scenario_map(document)
    if set(scenarios) != set(SCENARIOS) or len(document["scenarios"]) != len(SCENARIOS):
        raise BenchmarkContractError("frozen scenario inventory mismatch")
    for scenario_id, expected in FROZEN_SCENARIO_SEMANTICS.items():
        actual = scenarios[scenario_id]
        if actual["pre_approval_terminal"] != expected["pre_approval_terminal"]:
            raise BenchmarkContractError("frozen scenario pre-terminal mismatch")
        if actual["post_approval_terminal"] != expected["post_approval_terminal"]:
            raise BenchmarkContractError("frozen scenario post-terminal mismatch")
        if actual["final_requirement"] != expected["final_requirement"]:
            raise BenchmarkContractError("frozen scenario final requirement mismatch")

    matrices = document["matrices"]
    _exact(
        [(entry["scale_id"], entry["scenario_id"]) for entry in matrices["coverage_cells"]],
        expected_coverage_cells(),
        "coverage",
    )
    _exact(
        [(entry["scale_id"], entry["scenario_id"]) for entry in matrices["performance_cells"]],
        expected_performance_cells(),
        "performance",
    )
    _exact(
        [
            (entry["scale_id"], entry["scenario_id"], entry["permutation_seed"])
            for entry in matrices["determinism_cells"]
        ],
        expected_determinism_cells(seeds),
        "determinism",
    )

    if document["repetitions"] != FROZEN_REPETITIONS:
        raise BenchmarkContractError("repetition policy mismatch")
    if document["rss_protocol"] != FROZEN_RSS_PROTOCOL:
        raise BenchmarkContractError("frozen RSS protocol mismatch")
    if document["thresholds"] != FROZEN_THRESHOLDS:
        raise BenchmarkContractError("frozen thresholds mismatch")
    if document["total_reference_budget_hours"] != 3:
        raise BenchmarkContractError("reference budget mismatch")
    if document["config_digest"] != recompute_config_digest(document):
        raise BenchmarkContractError("config digest mismatch")


def _zero(counts: Mapping[str, Any]) -> bool:
    return all(value == 0 for value in counts.values())


def _synthetic_entry_ok(entry: Mapping[str, Any]) -> bool:
    return (
        entry["decision_id"].startswith("V040-SYN-")
        and entry["source_locator"].startswith("plan:synthetic_")
        and "synthetic" in entry["derivation_formula"].lower()
        and (
            "fixture" in entry["derivation_formula"].lower()
            or "not a production" in entry["derivation_formula"].lower()
        )
    )


def _formal_entry_uses_synthetic_namespace(entry: Mapping[str, Any]) -> bool:
    derivation = entry["derivation_formula"].lower()
    return (
        entry["decision_id"].startswith("V040-SYN-")
        or entry["source_locator"].startswith("plan:synthetic_")
        or "synthetic" in derivation
        or "not a production decision mapping" in derivation
    )


def validate_projected_envelope_semantics(document: Mapping[str, Any]) -> None:
    entries = list(document["entries"])
    ids = [entry["decision_id"] for entry in entries]
    if len(entries) != 171 or len(ids) != len(set(ids)):
        raise BenchmarkContractError("envelope must have 171 unique decision ids")

    counts = {"V0.4.1": 0, "V0.4.2": 0, "V0.4.3": 0}
    for entry in entries:
        counts[entry["primary_version"]] += 1
        projected = entry["projected_counts"]
        if (
            not entry["source_locator"].strip()
            or not entry["derivation_formula"].strip()
            or any(
                isinstance(value, bool) or not isinstance(value, int) or value < 0
                for value in projected.values()
            )
        ):
            raise BenchmarkContractError("invalid projection evidence")

        if entry["primary_version"] != "V0.4.1":
            protected_ok = (
                _zero(projected)
                and entry["implementation_owner"] == "future_primary"
                and entry["requirement_kind"] == "protected_boundary"
                and entry["scope_topology"] == "unmodeled"
                and entry["conflict_approval_assumption"] == "unmodeled"
                and entry["zero_load_reason"] == "protected_boundary"
                and not entry["blocked_projection"]
            )
            if not protected_ok:
                raise BenchmarkContractError("protected boundary mismatch")
        elif entry["blocked_projection"]:
            if (
                not _zero(projected)
                or entry["zero_load_reason"] != "none"
                or entry["scope_topology"] != "unmodeled"
                or entry["conflict_approval_assumption"] != "unmodeled"
            ):
                raise BenchmarkContractError("blocked projection mismatch")
        elif _zero(projected) and entry["zero_load_reason"] == "none":
            raise BenchmarkContractError("zero load reason required")
        elif not _zero(projected) and entry["zero_load_reason"] != "none":
            raise BenchmarkContractError("nonzero projection cannot have zero-load reason")

    actual = {
        "v041_primary": counts["V0.4.1"],
        "v042_protected": counts["V0.4.2"],
        "v043_protected": counts["V0.4.3"],
    }
    if actual != {"v041_primary": 150, "v042_protected": 20, "v043_protected": 1}:
        raise BenchmarkContractError("150/20/1 population mismatch")
    if document["decision_population_summary"] != actual:
        raise BenchmarkContractError("population summary mismatch")

    if document["projection_kind"] == "synthetic_contract_fixture":
        if not all(_synthetic_entry_ok(entry) for entry in entries):
            raise BenchmarkContractError("synthetic envelope must use synthetic namespaces")
    else:
        if any(_formal_entry_uses_synthetic_namespace(entry) for entry in entries):
            raise BenchmarkContractError("formal envelope cannot use synthetic namespaces")

    if document["envelope_digest"] != recompute_envelope_digest(document):
        raise BenchmarkContractError("envelope digest mismatch")


def _round(value: Any, quantum: str) -> Decimal:
    return Decimal(str(value)).quantize(Decimal(quantum), rounding=ROUND_HALF_EVEN)


def _ratio(numerator: Any, denominator: Any) -> Decimal:
    denominator_decimal = Decimal(str(denominator))
    if denominator_decimal <= 0:
        raise BenchmarkContractError("ratio baseline must be positive")
    return (Decimal(str(numerator)) / denominator_decimal).quantize(
        Decimal("0.000001"), rounding=ROUND_HALF_EVEN
    )


def _scale_limit(document: Mapping[str, Any]) -> Mapping[str, Any]:
    scale_id = document["parameters"]["scale_id"]
    return next(
        entry for entry in document["thresholds"]["scale_limits"]
        if entry["scale_id"] == scale_id
    )


def _expected_scenario(scenario_id: str) -> Mapping[str, Any]:
    return FROZEN_SCENARIO_SEMANTICS[scenario_id]


def _terminal_bad(document: Mapping[str, Any]) -> bool:
    expected = _expected_scenario(document["parameters"]["scenario_id"])
    for run in document["runs"]:
        if run["run_status"] != "completed":
            continue
        trace = run["terminal_trace"]
        expected_final = expected["terminal_state"] == "final"
        if (
            trace["pre_approval"] != expected["pre_approval_terminal"]
            or trace["post_approval"] != expected["post_approval_terminal"]
            or run["terminal_state"] != expected["terminal_state"]
            or run["final_profile_present"] != expected_final
        ):
            return True
        fingerprint_present = isinstance(run["final_profile_fingerprint"], str) and bool(
            SHA256_RE.fullmatch(str(run["final_profile_fingerprint"]))
        )
        if fingerprint_present != expected_final:
            return True
    return (
        document["execution_status"] == "completed"
        and document["composer_terminal_state"] != expected["terminal_state"]
    )


def _ratio_stop_reasons_from_document(document: Mapping[str, Any]) -> set[str]:
    reasons: set[str] = set()
    for metric, threshold_field in RATIO_THRESHOLD_FIELDS.items():
        evidence = document["ratio_evidence"][metric]
        if evidence["status"] != "measured":
            continue
        measured = _ratio(evidence["observed_2x"], evidence["baseline_1x"])
        if measured > Decimal(str(document["thresholds"][threshold_field])):
            reasons.add(RATIO_STOP_REASONS[metric])
    return reasons


def derive_stop_reasons(document: Mapping[str, Any]) -> list[str]:
    reasons: set[str] = set()
    subject_status = document["subject_digest_status"]
    scale_limit = _scale_limit(document)
    if subject_status["state"] == "stale":
        reasons.add("subject_stale")

    for run in document["runs"]:
        status = run["run_status"]
        if status == "timeout":
            reasons.add("timeout")
        if status == "process_crash":
            reasons.add("process_crash")
        if run["rss"]["status"] == "unavailable":
            reasons.add("rss_unavailable")
        for field, bad_value, reason in (
            ("coverage_conservation", "failed", "coverage_conservation_failure"),
            ("stable_id_status", "drift", "stable_id_drift"),
            ("canonical_determinism", "mismatched", "canonical_nondeterminism"),
            ("fingerprint_determinism", "mismatched", "fingerprint_nondeterminism"),
            ("contract_status", "error", "contract_error"),
        ):
            if run[field] == bad_value:
                reasons.add(reason)

        if run["run_kind"] in ABSOLUTE_GATE_RUN_KINDS and status == "completed":
            end_to_end = run["timings"]["end_to_end"]
            if end_to_end["status"] != "measured":
                reasons.add("contract_error")
            elif Decimal(str(end_to_end["wall_seconds"])) > Decimal(
                str(scale_limit["max_wall_seconds"])
            ):
                reasons.add("threshold_exceeded")
            if (
                run["rss"]["status"] == "available"
                and Decimal(str(run["rss"]["peak_rss_mib"]))
                > Decimal(str(scale_limit["max_peak_rss_mib"]))
            ):
                reasons.add("threshold_exceeded")

    if _terminal_bad(document):
        reasons.add("terminal_state_mismatch")
    reasons.update(_ratio_stop_reasons_from_document(document))
    if Decimal(str(document["reference_budget"]["elapsed_hours"])) > Decimal(
        str(document["reference_budget"]["limit_hours"])
    ):
        reasons.add("reference_budget_exceeded")
    return sorted(reasons, key=STOP_REASON_INDEX.__getitem__)


def _subject_ok(document: Mapping[str, Any]) -> None:
    observed = recompute_subject_digest(list(document["subject_manifest"]))
    status = document["subject_digest_status"]
    if status["observed_subject_digest"] != observed:
        raise BenchmarkContractError("observed subject digest mismatch")
    current = observed == document["benchmark_subject_digest"]
    if current != (status["state"] == "current"):
        raise BenchmarkContractError("subject state mismatch")
    if status["revalidation_required"] != (not current):
        raise BenchmarkContractError("subject revalidation flag mismatch")


def _runs_ok(document: Mapping[str, Any]) -> None:
    runs = list(document["runs"])
    kind = document["parameters"]["measurement_kind"]
    seed = document["parameters"]["permutation_seed"]
    if [run["run_index"] for run in runs] != list(range(1, len(runs) + 1)):
        raise BenchmarkContractError("run indices invalid")

    interrupted_indexes = [
        index for index, run in enumerate(runs)
        if run["run_status"] in {"timeout", "process_crash"}
    ]
    if interrupted_indexes:
        if interrupted_indexes != [len(runs) - 1]:
            raise BenchmarkContractError("timeout/crash must be the final run")
        if document["execution_status"] != "stopped":
            raise BenchmarkContractError("interrupted invocation must be stopped")
        if document["composer_terminal_state"] != "not_reached":
            raise BenchmarkContractError("interrupted invocation cannot claim composer terminal")

    if kind == "determinism":
        if isinstance(seed, bool) or not isinstance(seed, int):
            raise BenchmarkContractError("determinism result requires a seed")
        if len(runs) > 1 or any(run["run_kind"] != "determinism" for run in runs):
            raise BenchmarkContractError("determinism run shape invalid")
        if document["execution_status"] == "completed" and len(runs) != 1:
            raise BenchmarkContractError("completed determinism requires one run")
    elif seed is not None:
        raise BenchmarkContractError("non-determinism result cannot bind permutation seed")

    if kind == "performance":
        expected = ["performance_warmup", "performance_measured", "performance_measured", "performance_measured"]
        actual = [run["run_kind"] for run in runs]
        if actual != expected[: len(actual)]:
            raise BenchmarkContractError("performance run prefix invalid")
        if document["execution_status"] == "completed" and actual != expected:
            raise BenchmarkContractError("completed performance run shape invalid")
    if kind == "coverage":
        if any(run["run_kind"] != "coverage" for run in runs):
            raise BenchmarkContractError("coverage run kind invalid")
        if document["execution_status"] == "completed" and not runs:
            raise BenchmarkContractError("completed coverage requires a run")


def _timing_statuses(run: Mapping[str, Any]) -> dict[str, str]:
    return {stage: run["timings"][stage]["status"] for stage in TIMING_STAGES}


def _validate_completed_timing(run: Mapping[str, Any], scenario_id: str) -> None:
    statuses = _timing_statuses(run)
    for stage in (
        "synthetic_generation",
        "schema_registry_validation",
        "compose",
        "canonical_serialization",
        "end_to_end",
    ):
        if statuses[stage] != "measured":
            raise BenchmarkContractError("completed run missing required timing evidence")
    if run["final_profile_present"]:
        if statuses["apply"] != "measured":
            raise BenchmarkContractError("completed final run requires apply timing")
    elif statuses["apply"] not in {"measured", "not_applicable"}:
        raise BenchmarkContractError("non-final apply timing invalid")
    if scenario_id == "mixed-conflict-approval":
        if statuses["approval_generation"] != "measured":
            raise BenchmarkContractError("mixed approval requires approval-generation timing")
    elif statuses["approval_generation"] != "not_applicable":
        raise BenchmarkContractError("approval timing must be not_applicable outside mixed scenario")


def _validate_interrupted_timing(run: Mapping[str, Any]) -> None:
    statuses = _timing_statuses(run)
    if any(status not in {"measured", "not_reached"} for status in statuses.values()):
        raise BenchmarkContractError("interrupted timing must be measured/not_reached")
    if statuses["end_to_end"] != "measured":
        raise BenchmarkContractError("interrupted run requires measured elapsed time")
    ordered = [statuses[stage] for stage in TIMING_STAGES[:-1]]
    saw_not_reached = False
    for status in ordered:
        if status == "not_reached":
            saw_not_reached = True
        elif saw_not_reached:
            raise BenchmarkContractError("interrupted timing cannot resume after not_reached")


def _validate_metrics(metrics: Any) -> None:
    if not isinstance(metrics, Mapping):
        raise BenchmarkContractError("completed run requires C1 metrics")
    if set(metrics) != set(C1_METRIC_FIELDS):
        raise BenchmarkContractError("C1 metrics key set invalid")
    if any(
        isinstance(value, bool) or not isinstance(value, int) or value < 0
        for value in metrics.values()
    ):
        raise BenchmarkContractError("C1 metrics invalid")


def _run_evidence_ok(document: Mapping[str, Any]) -> None:
    scenario_id = document["parameters"]["scenario_id"]
    measurement_kind = document["parameters"]["measurement_kind"]
    for run in document["runs"]:
        if set(run["timings"]) != set(TIMING_STAGES):
            raise BenchmarkContractError("seven-stage timing key set invalid")

        rss = run["rss"]
        if rss["status"] == "available":
            baseline, peak, delta = map(
                lambda key: Decimal(str(rss[key])),
                ("baseline_rss_mib", "peak_rss_mib", "delta_peak_rss_mib"),
            )
            if peak < baseline or delta != _round(peak - baseline, "0.001"):
                raise BenchmarkContractError("RSS delta mismatch")

        if run["run_status"] == "completed":
            _validate_completed_timing(run, scenario_id)
            _validate_metrics(run["metrics"])
            if run["coverage_conservation"] not in {"passed", "failed"}:
                raise BenchmarkContractError("completed run requires coverage evidence")
            if run["stable_id_status"] not in {"stable", "drift"}:
                raise BenchmarkContractError("completed run requires stable-id evidence")
            if measurement_kind == "determinism":
                if run["canonical_determinism"] not in {"matched", "mismatched"}:
                    raise BenchmarkContractError("completed determinism requires canonical evidence")
                if run["fingerprint_determinism"] not in {"matched", "mismatched"}:
                    raise BenchmarkContractError("completed determinism requires fingerprint evidence")
            else:
                if run["canonical_determinism"] != "not_applicable":
                    raise BenchmarkContractError("non-determinism canonical evidence must be not_applicable")
                if run["fingerprint_determinism"] != "not_applicable":
                    raise BenchmarkContractError("non-determinism fingerprint evidence must be not_applicable")
        else:
            _validate_interrupted_timing(run)
            if (
                run["terminal_state"] != "not_reached"
                or run["terminal_trace"] != {"pre_approval": "not_reached", "post_approval": "not_reached"}
                or run["final_profile_present"]
                or run["final_profile_fingerprint"] is not None
                or run["metrics"] is not None
                or run["input_json_bytes"] is not None
                or run["output_json_bytes"] is not None
                or run["coverage_conservation"] != "not_reached"
                or run["stable_id_status"] != "not_reached"
                or run["canonical_determinism"] != "not_reached"
                or run["fingerprint_determinism"] != "not_reached"
            ):
                raise BenchmarkContractError("interrupted run carries reached/final evidence")


def _summary_ok(document: Mapping[str, Any]) -> None:
    summary = document["summary"]
    if (
        document["parameters"]["measurement_kind"] != "performance"
        or document["execution_status"] != "completed"
    ):
        if summary is not None:
            raise BenchmarkContractError("summary only allowed for completed performance")
        return

    measured_runs = [
        run for run in document["runs"] if run["run_kind"] == "performance_measured"
    ]
    if any(run["rss"]["status"] != "available" for run in measured_runs):
        raise BenchmarkContractError("completed performance summary requires RSS evidence")
    expected = {
        "median_wall_seconds": median(
            [run["timings"]["end_to_end"]["wall_seconds"] for run in measured_runs]
        ),
        "max_wall_seconds": max(
            run["timings"]["end_to_end"]["wall_seconds"] for run in measured_runs
        ),
        "median_peak_rss_mib": median(
            [run["rss"]["peak_rss_mib"] for run in measured_runs]
        ),
        "max_peak_rss_mib": max(run["rss"]["peak_rss_mib"] for run in measured_runs),
        "median_output_json_bytes": int(
            median([run["output_json_bytes"] for run in measured_runs])
        ),
        "max_output_json_bytes": max(run["output_json_bytes"] for run in measured_runs),
    }
    if summary != expected:
        raise BenchmarkContractError("summary does not reconstruct")


def validate_benchmark_result_semantics(document: Mapping[str, Any]) -> None:
    if not COMMIT_RE.fullmatch(document["benchmark_subject_commit"]):
        raise BenchmarkContractError("invalid subject commit")
    if document["benchmark_subject_digest_basis"] != "canonical_subject_manifest_v1":
        raise BenchmarkContractError("subject digest basis mismatch")
    if document["rss_protocol"] != FROZEN_RSS_PROTOCOL:
        raise BenchmarkContractError("frozen result RSS protocol mismatch")
    if document["thresholds"] != FROZEN_THRESHOLDS:
        raise BenchmarkContractError("frozen result thresholds mismatch")
    if document["reference_budget"]["limit_hours"] != 3:
        raise BenchmarkContractError("frozen result budget mismatch")

    _subject_ok(document)
    _runs_ok(document)
    _run_evidence_ok(document)
    _summary_ok(document)

    for evidence in document["ratio_evidence"].values():
        if evidence["status"] == "measured" and Decimal(str(evidence["ratio"])) != _ratio(
            evidence["observed_2x"], evidence["baseline_1x"]
        ):
            raise BenchmarkContractError("ratio evidence arithmetic mismatch")

    reasons = derive_stop_reasons(document)
    if list(document["stop_reasons"]) != reasons:
        raise BenchmarkContractError("stop reasons not evidence-derived")
    if document["overall_gate"] == "go" and (
        document["execution_status"] != "completed" or reasons
    ):
        raise BenchmarkContractError("GO state inconsistent")
    if document["overall_gate"] == "stop" and not reasons:
        raise BenchmarkContractError("STOP requires reason")
    if document["execution_status"] == "completed" and any(
        run["run_status"] != "completed" for run in document["runs"]
    ):
        raise BenchmarkContractError("completed execution contains interrupted run")
    if document["result_digest"] != recompute_result_digest(document):
        raise BenchmarkContractError("result digest mismatch")


def _matrix_sets(config: Mapping[str, Any]) -> tuple[set[Any], set[Any], set[Any]]:
    matrices = config["matrices"]
    coverage = {(entry["scale_id"], entry["scenario_id"]) for entry in matrices["coverage_cells"]}
    performance = {(entry["scale_id"], entry["scenario_id"]) for entry in matrices["performance_cells"]}
    determinism = {
        (entry["scale_id"], entry["scenario_id"], entry["permutation_seed"])
        for entry in matrices["determinism_cells"]
    }
    return coverage, performance, determinism


def _validate_repetition_against_config(result: Mapping[str, Any], config: Mapping[str, Any]) -> None:
    repetitions = config["repetitions"]
    runs = list(result["runs"])
    kind = result["parameters"]["measurement_kind"]
    if kind == "performance":
        expected = ["performance_warmup"] * repetitions["performance_warmup_runs_per_cell"] + [
            "performance_measured"
        ] * repetitions["performance_measured_runs_per_cell"]
        actual = [run["run_kind"] for run in runs]
        if actual != expected[: len(actual)]:
            raise BenchmarkContractError("performance repetition does not match config")
        if result["execution_status"] == "completed" and actual != expected:
            raise BenchmarkContractError("completed performance repetitions incomplete")
    elif kind == "determinism":
        required = repetitions["determinism_runs_per_seed"]
        if result["execution_status"] == "completed" and len(runs) != required:
            raise BenchmarkContractError("determinism repetitions do not match config")
    else:
        minimum = repetitions["coverage_min_runs_per_cell"]
        if result["execution_status"] == "completed" and len(runs) < minimum:
            raise BenchmarkContractError("coverage repetitions do not match config")


def validate_benchmark_result_against_config(
    result: Mapping[str, Any], config: Mapping[str, Any]
) -> None:
    """Validate a result and bind every frozen execution parameter to its config."""
    validate_benchmark_config_semantics(config)
    validate_benchmark_result_semantics(result)

    if result["benchmark_config_digest"] != recompute_config_digest(config):
        raise BenchmarkContractError("result is not bound to supplied config digest")
    if result["parameters"]["generation_seed"] != config["generation"]["generation_seed"]:
        raise BenchmarkContractError("generation seed differs from config")
    if result["rss_protocol"] != config["rss_protocol"]:
        raise BenchmarkContractError("RSS protocol differs from config")
    if result["thresholds"] != config["thresholds"]:
        raise BenchmarkContractError("thresholds differ from config")
    if result["output_json_bytes_basis"] != config["output_json_bytes_basis"]:
        raise BenchmarkContractError("output byte basis differs from config")
    if result["input_json_bytes_basis"] != config["input_json_bytes_basis"]:
        raise BenchmarkContractError("input byte basis differs from config")
    if result["reference_budget"]["limit_hours"] != config["total_reference_budget_hours"]:
        raise BenchmarkContractError("reference budget differs from config")

    coverage, performance, determinism = _matrix_sets(config)
    params = result["parameters"]
    cell = (params["scale_id"], params["scenario_id"])
    kind = params["measurement_kind"]
    if kind == "coverage":
        if cell not in coverage:
            raise BenchmarkContractError("coverage result cell not in config matrix")
    elif kind == "performance":
        if cell not in performance:
            raise BenchmarkContractError("performance result cell not in config matrix")
    else:
        seed = params["permutation_seed"]
        if seed not in set(config["generation"]["permutation_seeds"]):
            raise BenchmarkContractError("determinism seed not frozen by config")
        if params["scale_id"] not in DETERMINISM_SCALES:
            raise BenchmarkContractError("determinism scale outside frozen scales")
        if (params["scale_id"], params["scenario_id"], seed) not in determinism:
            raise BenchmarkContractError("determinism result cell not in config matrix")

    if kind == "performance" and params["scale_id"] == "2.0x":
        scenario = params["scenario_id"]
        if ("1.0x", scenario) in performance and ("2.0x", scenario) in performance:
            if any(evidence["status"] != "measured" for evidence in result["ratio_evidence"].values()):
                raise BenchmarkContractError("required 2x ratio evidence cannot be not_applicable")

    _validate_repetition_against_config(result, config)


def _expected_ratio_values(
    result_1x: Mapping[str, Any], result_2x: Mapping[str, Any]
) -> dict[str, tuple[Any, Any, Decimal]]:
    if result_1x["summary"] is None or result_2x["summary"] is None:
        raise BenchmarkContractError("ratio pair requires completed performance summaries")
    values: dict[str, tuple[Any, Any, Decimal]] = {}
    for metric, summary_field in RATIO_SUMMARY_FIELDS.items():
        baseline = result_1x["summary"][summary_field]
        observed = result_2x["summary"][summary_field]
        values[metric] = (baseline, observed, _ratio(observed, baseline))
    return values


def validate_ratio_evidence(
    result_1x: Mapping[str, Any] | None,
    result_2x: Mapping[str, Any] | None,
    config: Mapping[str, Any],
) -> None:
    """Bind the 2x ratio record to real 1x/2x performance result summaries."""
    if result_1x is None or result_2x is None:
        raise BenchmarkContractError("required ratio pair is incomplete")
    validate_benchmark_result_against_config(result_1x, config)
    validate_benchmark_result_against_config(result_2x, config)

    p1 = result_1x["parameters"]
    p2 = result_2x["parameters"]
    if p1["measurement_kind"] != "performance" or p2["measurement_kind"] != "performance":
        raise BenchmarkContractError("ratio pair must use performance results")
    if p1["scale_id"] != "1.0x" or p2["scale_id"] != "2.0x":
        raise BenchmarkContractError("ratio pair must bind 1.0x and 2.0x")
    if p1["scenario_id"] != p2["scenario_id"]:
        raise BenchmarkContractError("ratio pair scenario mismatch")
    if result_1x["execution_status"] != "completed" or result_2x["execution_status"] != "completed":
        raise BenchmarkContractError("ratio pair requires completed results")

    performance_cells = _matrix_sets(config)[1]
    scenario = p1["scenario_id"]
    if ("1.0x", scenario) not in performance_cells or ("2.0x", scenario) not in performance_cells:
        raise BenchmarkContractError("ratio pair is not required by frozen matrix")

    expected_values = _expected_ratio_values(result_1x, result_2x)
    for metric, (baseline, observed, ratio) in expected_values.items():
        evidence = result_2x["ratio_evidence"][metric]
        if evidence["status"] != "measured":
            raise BenchmarkContractError("required ratio cannot be not_applicable")
        if Decimal(str(evidence["baseline_1x"])) != Decimal(str(baseline)):
            raise BenchmarkContractError("ratio baseline is not bound to 1x summary")
        if Decimal(str(evidence["observed_2x"])) != Decimal(str(observed)):
            raise BenchmarkContractError("ratio observed value is not bound to 2x summary")
        if Decimal(str(evidence["ratio"])) != ratio:
            raise BenchmarkContractError("ratio value does not recompute from summaries")

    expected_ratio_reasons = {
        RATIO_STOP_REASONS[metric]
        for metric, (_, _, ratio) in expected_values.items()
        if ratio > Decimal(str(config["thresholds"][RATIO_THRESHOLD_FIELDS[metric]]))
    }
    actual_ratio_reasons = set(result_2x["stop_reasons"]) & set(RATIO_STOP_REASONS.values())
    if actual_ratio_reasons != expected_ratio_reasons:
        raise BenchmarkContractError("ratio stop reasons do not match recomputed pair")
    if expected_ratio_reasons and result_2x["overall_gate"] != "stop":
        raise BenchmarkContractError("ratio threshold breach must STOP")
    if not expected_ratio_reasons and result_2x["overall_gate"] == "stop":
        intrinsic = set(result_2x["stop_reasons"]) - set(RATIO_STOP_REASONS.values())
        if not intrinsic:
            raise BenchmarkContractError("ratio-passing result cannot STOP without another reason")


def validate_benchmark_result_set(
    results: Iterable[Mapping[str, Any]], config: Mapping[str, Any]
) -> None:
    """Validate logical uniqueness and required 1x/2x performance ratio bindings."""
    validate_benchmark_config_semantics(config)
    result_list = list(results)
    index: dict[tuple[Any, ...], Mapping[str, Any]] = {}
    for result in result_list:
        validate_benchmark_result_against_config(result, config)
        params = result["parameters"]
        key = (
            params["measurement_kind"],
            params["scale_id"],
            params["scenario_id"],
            params["permutation_seed"],
        )
        if key in index:
            raise BenchmarkContractError("duplicate benchmark result logical key")
        index[key] = result

    performance_cells = _matrix_sets(config)[1]
    ratio_scenarios = {
        scenario
        for scenario in SCENARIOS
        if ("1.0x", scenario) in performance_cells and ("2.0x", scenario) in performance_cells
    }
    for scenario in ratio_scenarios:
        one = index.get(("performance", "1.0x", scenario, None))
        two = index.get(("performance", "2.0x", scenario, None))
        if (one is None) != (two is None):
            raise BenchmarkContractError("required 1x/2x performance baseline pair missing")
        if one is not None and two is not None:
            validate_ratio_evidence(one, two, config)
