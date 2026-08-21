import copy
import hashlib
import importlib.util
import json
import math
import unittest
from decimal import Decimal, ROUND_HALF_EVEN
from pathlib import Path
from statistics import median

from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
BENCH = ROOT / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2"
FIX = ROOT / "tests" / "fixtures" / "v0412" / "p3a_c2"
MOD = ROOT / "format-monograph" / "scripts" / "profile_v2_benchmark.py"

spec = importlib.util.spec_from_file_location("profile_v2_benchmark", MOD)
b = importlib.util.module_from_spec(spec)
spec.loader.exec_module(b)

SCHEMAS = {
    "ENVELOPE": json.loads((BENCH / "projected-envelope.schema.json").read_text()),
    "CONFIG": json.loads((BENCH / "benchmark-config.schema.json").read_text()),
    "RESULT": json.loads((BENCH / "benchmark-result.schema.json").read_text()),
}
VALIDATORS = {key: Draft202012Validator(value) for key, value in SCHEMAS.items()}
ASSERTION_IDS = tuple(f"T412-C2A-XR-{i:03d}" for i in range(1, 111))


def load(name):
    return json.loads((FIX / name).read_text())


def schema_valid(kind, document):
    return not list(VALIDATORS[kind].iter_errors(document))


def semantic_ok(function, *args):
    try:
        function(*args)
        return True
    except b.BenchmarkContractError:
        return False


def independent_digest(document, key):
    payload = copy.deepcopy(document)
    payload.pop(key, None)
    data = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def digest_only(document):
    document["result_digest"] = independent_digest(document, "result_digest")
    return document


def settle(document):
    document["stop_reasons"] = b.derive_stop_reasons(document)
    document["overall_gate"] = "stop" if document["stop_reasons"] else "go"
    return digest_only(document)


def stamp_config(document):
    document["config_digest"] = independent_digest(document, "config_digest")
    return document


def stamp_envelope(document):
    document["envelope_digest"] = independent_digest(document, "envelope_digest")
    return document


def expected_terminal(scenario):
    return b.FROZEN_SCENARIO_SEMANTICS[scenario]


def apply_scenario(run, scenario):
    expected = expected_terminal(scenario)
    final = expected["terminal_state"] == "final"
    run["terminal_state"] = expected["terminal_state"]
    run["terminal_trace"] = {
        "pre_approval": expected["pre_approval_terminal"],
        "post_approval": expected["post_approval_terminal"],
    }
    run["final_profile_present"] = final
    run["final_profile_fingerprint"] = "sha256:" + "d" * 64 if final else None
    run["timings"]["approval_generation"] = (
        {"status": "measured", "wall_seconds": 0.03}
        if scenario == "mixed-conflict-approval"
        else {"status": "not_applicable"}
    )
    run["timings"]["apply"] = (
        {"status": "measured", "wall_seconds": 0.04}
        if final
        else {"status": "not_applicable"}
    )
    return run


def rebuild_summary(document):
    measured = [run for run in document["runs"] if run["run_kind"] == "performance_measured"]
    document["summary"] = {
        "median_wall_seconds": median(
            [run["timings"]["end_to_end"]["wall_seconds"] for run in measured]
        ),
        "max_wall_seconds": max(
            run["timings"]["end_to_end"]["wall_seconds"] for run in measured
        ),
        "median_peak_rss_mib": median([run["rss"]["peak_rss_mib"] for run in measured]),
        "max_peak_rss_mib": max(run["rss"]["peak_rss_mib"] for run in measured),
        "median_output_json_bytes": int(median([run["output_json_bytes"] for run in measured])),
        "max_output_json_bytes": max(run["output_json_bytes"] for run in measured),
    }
    return document


def make_performance(scale, scenario, walls=None, rss=None, outputs=None):
    document = load("benchmark-result.valid.json")
    document["parameters"].update(
        {
            "scale_id": scale,
            "scenario_id": scenario,
            "measurement_kind": "performance",
            "permutation_seed": None,
        }
    )
    document["composer_terminal_state"] = expected_terminal(scenario)["terminal_state"]
    walls = walls or [0.45, 0.3, 0.4, 0.5]
    rss = rss or [29.0, 30.0, 31.0, 32.0]
    outputs = outputs or [1200, 1200, 1200, 1200]
    for index, run in enumerate(document["runs"]):
        run["run_index"] = index + 1
        run["run_kind"] = "performance_warmup" if index == 0 else "performance_measured"
        run["run_status"] = "completed"
        run["timings"]["end_to_end"] = {"status": "measured", "wall_seconds": walls[index]}
        run["rss"] = {
            "status": "available",
            "baseline_rss_mib": 20.0,
            "peak_rss_mib": rss[index],
            "delta_peak_rss_mib": round(rss[index] - 20.0, 3),
        }
        run["output_json_bytes"] = outputs[index]
        run["coverage_conservation"] = "passed"
        run["stable_id_status"] = "stable"
        run["canonical_determinism"] = "not_applicable"
        run["fingerprint_determinism"] = "not_applicable"
        apply_scenario(run, scenario)
    document["execution_status"] = "completed"
    document["ratio_evidence"] = {
        "wall": {"status": "not_applicable"},
        "rss": {"status": "not_applicable"},
        "output_json": {"status": "not_applicable"},
    }
    rebuild_summary(document)
    return settle(document)


def make_coverage(scale="0.5x", scenario="disjoint"):
    document = make_performance("1.5x", scenario)
    run = copy.deepcopy(document["runs"][0])
    run["run_index"] = 1
    run["run_kind"] = "coverage"
    document["parameters"].update(
        {
            "scale_id": scale,
            "scenario_id": scenario,
            "measurement_kind": "coverage",
            "permutation_seed": None,
        }
    )
    document["runs"] = [run]
    document["summary"] = None
    document["composer_terminal_state"] = expected_terminal(scenario)["terminal_state"]
    return settle(document)


def make_determinism(scale="1.5x", scenario="disjoint", seed=7):
    document = make_performance("1.5x", scenario)
    run = copy.deepcopy(document["runs"][0])
    run["run_index"] = 1
    run["run_kind"] = "determinism"
    run["canonical_determinism"] = "matched"
    run["fingerprint_determinism"] = "matched"
    document["parameters"].update(
        {
            "scale_id": scale,
            "scenario_id": scenario,
            "measurement_kind": "determinism",
            "permutation_seed": seed,
        }
    )
    document["runs"] = [run]
    document["summary"] = None
    document["composer_terminal_state"] = expected_terminal(scenario)["terminal_state"]
    return settle(document)


def make_interrupted(status="timeout", interrupt_index=2):
    document = make_performance("1.5x", "disjoint")
    runs = document["runs"][:interrupt_index]
    interrupted = runs[-1]
    interrupted["run_status"] = status
    interrupted["terminal_state"] = "not_reached"
    interrupted["terminal_trace"] = {
        "pre_approval": "not_reached",
        "post_approval": "not_reached",
    }
    interrupted["final_profile_present"] = False
    interrupted["final_profile_fingerprint"] = None
    interrupted["metrics"] = None
    interrupted["input_json_bytes"] = None
    interrupted["output_json_bytes"] = None
    interrupted["coverage_conservation"] = "not_reached"
    interrupted["stable_id_status"] = "not_reached"
    interrupted["canonical_determinism"] = "not_reached"
    interrupted["fingerprint_determinism"] = "not_reached"
    interrupted["timings"] = {
        "synthetic_generation": {"status": "measured", "wall_seconds": 0.01},
        "schema_registry_validation": {"status": "measured", "wall_seconds": 0.02},
        "compose": {"status": "not_reached"},
        "approval_generation": {"status": "not_reached"},
        "apply": {"status": "not_reached"},
        "canonical_serialization": {"status": "not_reached"},
        "end_to_end": {"status": "measured", "wall_seconds": 30.0},
    }
    document["runs"] = runs
    document["execution_status"] = "stopped"
    document["composer_terminal_state"] = "not_reached"
    document["summary"] = None
    return settle(document)


def bind_ratio(one, two, fake_wall=None):
    fields = {
        "wall": "median_wall_seconds",
        "rss": "median_peak_rss_mib",
        "output_json": "median_output_json_bytes",
    }
    for metric, field in fields.items():
        baseline = one["summary"][field]
        observed = two["summary"][field]
        if metric == "wall" and fake_wall is not None:
            baseline, observed = fake_wall
        ratio = (Decimal(str(observed)) / Decimal(str(baseline))).quantize(
            Decimal("0.000001"), rounding=ROUND_HALF_EVEN
        )
        two["ratio_evidence"][metric] = {
            "status": "measured",
            "baseline_1x": baseline,
            "observed_2x": observed,
            "ratio": float(ratio),
        }
    return settle(two)


def make_envelope():
    entries = []
    for i in range(1, 151):
        blocked = i == 150
        entries.append(
            {
                "decision_id": f"V040-SYN-V041-{i:03d}",
                "source_locator": f"plan:synthetic_v041_{i:03d}",
                "primary_version": "V0.4.1",
                "implementation_owner": "p3a_c",
                "requirement_kind": "system_projection",
                "projected_counts": {
                    "rule_fragment": 0 if blocked else 1,
                    "binding": 0 if blocked else 2,
                    "key": 0 if blocked else 1,
                    "candidate": 0 if blocked else 2,
                },
                "scope_topology": "unmodeled" if blocked else "single_core_probe",
                "conflict_approval_assumption": "unmodeled" if blocked else "none",
                "derivation_formula": "synthetic contract fixture only; not a production decision mapping",
                "confidence": "low",
                "zero_load_reason": "none",
                "blocked_projection": blocked,
            }
        )
    for version, count, label in [("V0.4.2", 20, "V042"), ("V0.4.3", 1, "V043")]:
        for i in range(1, count + 1):
            entries.append(
                {
                    "decision_id": f"V040-SYN-{label}-{i:03d}",
                    "source_locator": f"plan:synthetic_{label.lower()}_{i:03d}",
                    "primary_version": version,
                    "implementation_owner": "future_primary",
                    "requirement_kind": "protected_boundary",
                    "projected_counts": {"rule_fragment": 0, "binding": 0, "key": 0, "candidate": 0},
                    "scope_topology": "unmodeled",
                    "conflict_approval_assumption": "unmodeled",
                    "derivation_formula": "synthetic protected-boundary fixture only; zero benchmark load",
                    "confidence": "low",
                    "zero_load_reason": "protected_boundary",
                    "blocked_projection": False,
                }
            )
    document = {
        "document_kind": "p3a_c2_projected_envelope",
        "contract_version": "1.0",
        "projection_kind": "synthetic_contract_fixture",
        "projection_status": "planning_projection",
        "compose_projection_strategy": "single_core_probe",
        "source_plan_digest": "sha256:" + "a" * 64,
        "envelope_digest": "sha256:" + "0" * 64,
        "envelope_digest_basis": "canonical_json_excluding_envelope_digest",
        "derivation_policy": "manual_prediction_with_later_mechanical_validation",
        "unmodeled_dimensions": ["docx_runtime", "reference_hardware"],
        "decision_population_summary": {"v041_primary": 150, "v042_protected": 20, "v043_protected": 1},
        "entries": entries,
    }
    return stamp_envelope(document)


def formalize(document):
    formal = copy.deepcopy(document)
    formal["projection_kind"] = "formal_planning_projection"
    for index, entry in enumerate(formal["entries"], 1):
        entry["decision_id"] = f"V040-FORMAL-{index:03d}"
        entry["source_locator"] = f"plan:formal_projection_{index:03d}"
        entry["derivation_formula"] = "formal planning projection boundary placeholder"
    return stamp_envelope(formal)


class C2AXRClosureTests(unittest.TestCase):
    executed = set()

    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        expected = set(ASSERTION_IDS)
        if cls.executed != expected:
            missing = sorted(expected - cls.executed)
            extra = sorted(cls.executed - expected)
            raise AssertionError(f"assertion execution mismatch: missing={missing}, extra={extra}")

    def check_range(self, start, checks):
        self.assertEqual(10, len(checks))
        for offset, condition in enumerate(checks):
            assertion_id = f"T412-C2A-XR-{start + offset:03d}"
            self.assertIn(assertion_id, ASSERTION_IDS)
            self.assertNotIn(assertion_id, type(self).executed)
            type(self).executed.add(assertion_id)
            self.assertTrue(condition, assertion_id)

    def test_01_xr001_result_config_binding(self):
        config = load("benchmark-config.valid.json")
        valid = load("benchmark-result.valid.json")
        disallowed = copy.deepcopy(valid)
        disallowed["parameters"]["scale_id"] = "0.5x"
        digest_only(disallowed)
        bad_det_scale = make_determinism("1.0x", seed=7)
        unknown_seed = make_determinism("1.5x", seed=999)
        wrong_digest = copy.deepcopy(valid)
        wrong_digest["benchmark_config_digest"] = "sha256:" + "e" * 64
        digest_only(wrong_digest)
        wrong_generation = copy.deepcopy(valid)
        wrong_generation["parameters"]["generation_seed"] = 42
        digest_only(wrong_generation)
        coverage = make_coverage()
        self.check_range(
            1,
            [
                schema_valid("RESULT", disallowed),
                semantic_ok(b.validate_benchmark_result_semantics, disallowed),
                not semantic_ok(b.validate_benchmark_result_against_config, disallowed, config),
                schema_valid("RESULT", bad_det_scale) and not semantic_ok(b.validate_benchmark_result_against_config, bad_det_scale, config),
                schema_valid("RESULT", unknown_seed) and not semantic_ok(b.validate_benchmark_result_against_config, unknown_seed, config),
                schema_valid("RESULT", wrong_digest) and not semantic_ok(b.validate_benchmark_result_against_config, wrong_digest, config),
                schema_valid("RESULT", valid) and semantic_ok(b.validate_benchmark_result_against_config, valid, config),
                not semantic_ok(b.validate_benchmark_result_against_config, wrong_generation, config),
                semantic_ok(b.validate_benchmark_result_against_config, coverage, config),
                valid["benchmark_config_digest"] == independent_digest(config, "config_digest"),
            ],
        )

    def test_02_xr002_mandatory_evidence(self):
        det_na = make_determinism()
        det_na["runs"][0]["canonical_determinism"] = "not_applicable"
        digest_only(det_na)
        coverage_cons = make_coverage()
        coverage_cons["runs"][0]["coverage_conservation"] = "not_reached"
        digest_only(coverage_cons)
        coverage_stable = make_coverage()
        coverage_stable["runs"][0]["stable_id_status"] = "not_reached"
        digest_only(coverage_stable)
        mixed = make_coverage(scenario="mixed-conflict-approval")
        mixed["runs"][0]["timings"]["approval_generation"] = {"status": "not_applicable"}
        digest_only(mixed)
        interrupted = make_interrupted()
        missing_metrics = make_coverage()
        missing_metrics["runs"][0]["metrics"] = None
        digest_only(missing_metrics)
        compose_not_reached = make_coverage()
        compose_not_reached["runs"][0]["timings"]["compose"] = {"status": "not_reached"}
        digest_only(compose_not_reached)
        self.check_range(
            11,
            [
                schema_valid("RESULT", det_na) and not semantic_ok(b.validate_benchmark_result_semantics, det_na),
                schema_valid("RESULT", coverage_cons) and not semantic_ok(b.validate_benchmark_result_semantics, coverage_cons),
                schema_valid("RESULT", coverage_stable) and not semantic_ok(b.validate_benchmark_result_semantics, coverage_stable),
                schema_valid("RESULT", mixed) and not semantic_ok(b.validate_benchmark_result_semantics, mixed),
                schema_valid("RESULT", interrupted) and semantic_ok(b.validate_benchmark_result_semantics, interrupted),
                interrupted["runs"][-1]["timings"]["compose"]["status"] == "not_reached",
                schema_valid("RESULT", missing_metrics) and not semantic_ok(b.validate_benchmark_result_semantics, missing_metrics),
                schema_valid("RESULT", compose_not_reached) and not semantic_ok(b.validate_benchmark_result_semantics, compose_not_reached),
                make_determinism()["runs"][0]["coverage_conservation"] == "passed",
                make_coverage(scenario="mixed-conflict-approval")["runs"][0]["timings"]["approval_generation"]["status"] == "measured",
            ],
        )

    def test_03_xr003_ratio_pair_binding(self):
        config = load("benchmark-config.valid.json")
        one = make_performance("1.0x", "mixed-conflict-approval", walls=[5, 10, 10, 10], rss=[80, 100, 100, 100], outputs=[800, 1000, 1000, 1000])
        two = make_performance("2.0x", "mixed-conflict-approval", walls=[10, 20, 20, 20], rss=[160, 200, 200, 200], outputs=[1600, 2000, 2000, 2000])
        pair_two = bind_ratio(one, copy.deepcopy(two))
        not_applicable = copy.deepcopy(two)
        fake_one = make_performance("1.0x", "mixed-conflict-approval", walls=[5, 10, 10, 10], rss=[80, 100, 100, 100], outputs=[800, 1000, 1000, 1000])
        fake_two = make_performance("2.0x", "mixed-conflict-approval", walls=[60, 70, 70, 70], rss=[160, 200, 200, 200], outputs=[1600, 2000, 2000, 2000])
        fake_two = bind_ratio(fake_one, fake_two, fake_wall=(20, 70))
        over_two = make_performance("2.0x", "mixed-conflict-approval", walls=[60, 70, 70, 70], rss=[160, 200, 200, 200], outputs=[1600, 2000, 2000, 2000])
        over_two = bind_ratio(fake_one, over_two)
        independent_wall = (Decimal("70") / Decimal("10")).quantize(Decimal("0.000001"))
        self.check_range(
            21,
            [
                schema_valid("RESULT", not_applicable) and not semantic_ok(b.validate_ratio_evidence, one, not_applicable, config),
                schema_valid("RESULT", fake_two) and semantic_ok(b.validate_benchmark_result_semantics, fake_two),
                not semantic_ok(b.validate_ratio_evidence, fake_one, fake_two, config),
                semantic_ok(b.validate_ratio_evidence, one, pair_two, config),
                semantic_ok(b.validate_benchmark_result_set, [one, pair_two], config),
                semantic_ok(b.validate_ratio_evidence, fake_one, over_two, config) and over_two["overall_gate"] == "stop",
                "wall_ratio_exceeded" in over_two["stop_reasons"],
                not semantic_ok(b.validate_benchmark_result_set, [over_two], config),
                Decimal(str(over_two["ratio_evidence"]["wall"]["ratio"])) == independent_wall,
                b.RATIO_SUMMARY_FIELDS == {"wall": "median_wall_seconds", "rss": "median_peak_rss_mib", "output_json": "median_output_json_bytes"},
            ],
        )

    def test_04_xr004_interrupted_run_closure(self):
        timeout = make_interrupted("timeout")
        timeout_final = copy.deepcopy(timeout)
        timeout_final["runs"][-1]["final_profile_present"] = True
        timeout_final["runs"][-1]["final_profile_fingerprint"] = "sha256:" + "d" * 64
        digest_only(timeout_final)
        crash_final = make_interrupted("process_crash")
        crash_final["runs"][-1]["terminal_state"] = "final"
        digest_only(crash_final)
        later = make_interrupted("timeout")
        later["runs"].append(copy.deepcopy(load("benchmark-result.valid.json")["runs"][2]))
        later["runs"][-1]["run_index"] = 3
        digest_only(later)
        completed = copy.deepcopy(timeout)
        completed["execution_status"] = "completed"
        digest_only(completed)
        self.check_range(
            31,
            [
                schema_valid("RESULT", timeout_final) and not semantic_ok(b.validate_benchmark_result_semantics, timeout_final),
                schema_valid("RESULT", crash_final) and not semantic_ok(b.validate_benchmark_result_semantics, crash_final),
                schema_valid("RESULT", later) and not semantic_ok(b.validate_benchmark_result_semantics, later),
                semantic_ok(b.validate_benchmark_result_semantics, timeout),
                schema_valid("RESULT", completed) and not semantic_ok(b.validate_benchmark_result_semantics, completed),
                timeout["runs"][-1]["terminal_state"] == "not_reached",
                timeout["runs"][-1]["final_profile_present"] is False and timeout["runs"][-1]["final_profile_fingerprint"] is None,
                timeout["runs"][-1]["terminal_trace"] == {"pre_approval": "not_reached", "post_approval": "not_reached"},
                timeout["execution_status"] == "stopped" and timeout["stop_reasons"] == ["timeout"],
                make_interrupted("process_crash")["stop_reasons"] == ["process_crash"],
            ],
        )

    def test_05_xr005_absolute_threshold_applicability(self):
        config = load("benchmark-config.valid.json")
        coverage_wall = make_coverage("0.5x")
        coverage_wall["runs"][0]["timings"]["end_to_end"]["wall_seconds"] = 61
        settle(coverage_wall)
        det_rss = make_determinism("1.5x")
        det_rss["runs"][0]["rss"] = {"status": "available", "baseline_rss_mib": 20.0, "peak_rss_mib": 600.0, "delta_peak_rss_mib": 580.0}
        settle(det_rss)
        normal = make_coverage("0.5x")
        two_over = make_coverage("2.0x")
        two_over["runs"][0]["timings"]["end_to_end"]["wall_seconds"] = 121
        settle(two_over)
        two_equal = make_coverage("2.0x")
        two_equal["runs"][0]["timings"]["end_to_end"]["wall_seconds"] = 120
        settle(two_equal)
        warmup = make_performance("1.5x", "disjoint")
        warmup["runs"][0]["timings"]["end_to_end"]["wall_seconds"] = 999
        settle(warmup)
        det_wall = make_determinism("1.5x")
        det_wall["runs"][0]["timings"]["end_to_end"]["wall_seconds"] = 61
        settle(det_wall)
        self.check_range(
            41,
            [
                semantic_ok(b.validate_benchmark_result_against_config, coverage_wall, config) and coverage_wall["overall_gate"] == "stop",
                "threshold_exceeded" in coverage_wall["stop_reasons"],
                semantic_ok(b.validate_benchmark_result_against_config, det_rss, config) and det_rss["overall_gate"] == "stop",
                semantic_ok(b.validate_benchmark_result_against_config, normal, config) and normal["overall_gate"] == "go",
                semantic_ok(b.validate_benchmark_result_against_config, two_over, config) and two_over["overall_gate"] == "stop",
                semantic_ok(b.validate_benchmark_result_against_config, two_equal, config) and two_equal["overall_gate"] == "go",
                semantic_ok(b.validate_benchmark_result_against_config, warmup, config) and warmup["overall_gate"] == "go",
                b.ABSOLUTE_GATE_RUN_KINDS == {"coverage", "performance_measured", "determinism"},
                semantic_ok(b.validate_benchmark_result_against_config, det_wall, config) and det_wall["overall_gate"] == "stop",
                b.FROZEN_THRESHOLDS["scale_limits"][-1] == {"scale_id": "2.0x", "max_wall_seconds": 120, "max_peak_rss_mib": 1024},
            ],
        )

    def test_06_xr006_config_frozen_mappings(self):
        config = load("benchmark-config.valid.json")
        bad_factor = copy.deepcopy(config)
        bad_factor["scales"][0]["factor"] = 2.0
        stamp_config(bad_factor)
        bad_dense = copy.deepcopy(config)
        bad_dense["scenarios"][2]["final_requirement"] = "required"
        stamp_config(bad_dense)
        bad_mixed = copy.deepcopy(config)
        bad_mixed["scenarios"][3]["pre_approval_terminal"] = "final"
        stamp_config(bad_mixed)
        bad_disjoint = copy.deepcopy(config)
        bad_disjoint["scenarios"][0]["pre_approval_terminal"] = "unresolvable"
        stamp_config(bad_disjoint)
        self.check_range(
            51,
            [
                schema_valid("CONFIG", bad_factor) and not semantic_ok(b.validate_benchmark_config_semantics, bad_factor),
                schema_valid("CONFIG", bad_dense) and not semantic_ok(b.validate_benchmark_config_semantics, bad_dense),
                schema_valid("CONFIG", bad_mixed) and not semantic_ok(b.validate_benchmark_config_semantics, bad_mixed),
                semantic_ok(b.validate_benchmark_config_semantics, config),
                schema_valid("CONFIG", bad_disjoint) and not semantic_ok(b.validate_benchmark_config_semantics, bad_disjoint),
                b.FROZEN_SCALE_FACTORS == {"0.5x": 0.5, "1.0x": 1.0, "1.5x": 1.5, "2.0x": 2.0},
                b.FROZEN_SCENARIO_SEMANTICS["dense-crossing"]["terminal_state"] == "unresolvable",
                b.FROZEN_SCENARIO_SEMANTICS["mixed-conflict-approval"]["pre_approval_terminal"] == "awaiting_approval",
                {(x["scale_id"], x["scenario_id"]) for x in config["matrices"]["performance_cells"]} == {("0.5x", "mixed-conflict-approval"), ("1.0x", "mixed-conflict-approval"), *((s, c) for s in ("1.5x", "2.0x") for c in b.SCENARIOS)},
                config["config_digest"] == independent_digest(config, "config_digest"),
            ],
        )

    def test_07_xr007_json_object_order_independence(self):
        valid = load("benchmark-result.valid.json")
        timing_reordered = copy.deepcopy(valid)
        timing_reordered["runs"][0]["timings"] = dict(reversed(list(timing_reordered["runs"][0]["timings"].items())))
        digest_only(timing_reordered)
        metrics_reordered = copy.deepcopy(valid)
        metrics_reordered["runs"][0]["metrics"] = dict(reversed(list(metrics_reordered["runs"][0]["metrics"].items())))
        digest_only(metrics_reordered)
        subject_reordered = copy.deepcopy(valid)
        subject_reordered["subject_manifest"][0] = {"sha256": subject_reordered["subject_manifest"][0]["sha256"], "path": subject_reordered["subject_manifest"][0]["path"]}
        digest_only(subject_reordered)
        missing_timing = copy.deepcopy(valid)
        del missing_timing["runs"][0]["timings"]["compose"]
        digest_only(missing_timing)
        extra_timing = copy.deepcopy(valid)
        extra_timing["runs"][0]["timings"]["extra"] = {"status": "not_reached"}
        digest_only(extra_timing)
        missing_metric = copy.deepcopy(valid)
        del missing_metric["runs"][0]["metrics"]["candidate_count"]
        digest_only(missing_metric)
        extra_metric = copy.deepcopy(valid)
        extra_metric["runs"][0]["metrics"]["extra"] = 1
        digest_only(extra_metric)
        self.check_range(
            61,
            [
                schema_valid("RESULT", timing_reordered) and semantic_ok(b.validate_benchmark_result_semantics, timing_reordered),
                schema_valid("RESULT", metrics_reordered) and semantic_ok(b.validate_benchmark_result_semantics, metrics_reordered),
                schema_valid("RESULT", subject_reordered) and semantic_ok(b.validate_benchmark_result_semantics, subject_reordered),
                independent_digest(valid, "result_digest") == independent_digest(json.loads(json.dumps(valid, sort_keys=True)), "result_digest"),
                not schema_valid("RESULT", missing_timing) and not semantic_ok(b.validate_benchmark_result_semantics, missing_timing),
                not schema_valid("RESULT", extra_timing) and not semantic_ok(b.validate_benchmark_result_semantics, extra_timing),
                not schema_valid("RESULT", missing_metric) and not semantic_ok(b.validate_benchmark_result_semantics, missing_metric),
                not schema_valid("RESULT", extra_metric) and not semantic_ok(b.validate_benchmark_result_semantics, extra_metric),
                set(valid["runs"][0]["timings"]) == set(b.TIMING_STAGES),
                set(valid["runs"][0]["metrics"]) == set(b.C1_METRIC_FIELDS),
            ],
        )

    def test_08_xr008_synthetic_formal_envelope_boundary(self):
        synthetic = make_envelope()
        renamed = copy.deepcopy(synthetic)
        renamed["projection_kind"] = "formal_planning_projection"
        stamp_envelope(renamed)
        formal = formalize(synthetic)
        formal_one_id = copy.deepcopy(formal)
        formal_one_id["entries"][0]["decision_id"] = "V040-SYN-V041-001"
        stamp_envelope(formal_one_id)
        formal_source = copy.deepcopy(formal)
        formal_source["entries"][0]["source_locator"] = "plan:synthetic_v041_001"
        stamp_envelope(formal_source)
        formal_formula = copy.deepcopy(formal)
        formal_formula["entries"][0]["derivation_formula"] = "synthetic contract fixture only; not a production decision mapping"
        stamp_envelope(formal_formula)
        synthetic_id = copy.deepcopy(synthetic)
        synthetic_id["entries"][0]["decision_id"] = "V040-FORMAL-001"
        stamp_envelope(synthetic_id)
        synthetic_source = copy.deepcopy(synthetic)
        synthetic_source["entries"][0]["source_locator"] = "plan:formal_projection_001"
        stamp_envelope(synthetic_source)
        synthetic_formula = copy.deepcopy(synthetic)
        synthetic_formula["entries"][0]["derivation_formula"] = "formal planning projection boundary placeholder"
        stamp_envelope(synthetic_formula)
        self.check_range(
            71,
            [
                schema_valid("ENVELOPE", synthetic) and semantic_ok(b.validate_projected_envelope_semantics, synthetic),
                schema_valid("ENVELOPE", renamed) and not semantic_ok(b.validate_projected_envelope_semantics, renamed),
                schema_valid("ENVELOPE", formal) and semantic_ok(b.validate_projected_envelope_semantics, formal),
                schema_valid("ENVELOPE", formal_one_id) and not semantic_ok(b.validate_projected_envelope_semantics, formal_one_id),
                schema_valid("ENVELOPE", formal_source) and not semantic_ok(b.validate_projected_envelope_semantics, formal_source),
                schema_valid("ENVELOPE", formal_formula) and not semantic_ok(b.validate_projected_envelope_semantics, formal_formula),
                schema_valid("ENVELOPE", synthetic_id) and not semantic_ok(b.validate_projected_envelope_semantics, synthetic_id),
                schema_valid("ENVELOPE", synthetic_source) and not semantic_ok(b.validate_projected_envelope_semantics, synthetic_source),
                schema_valid("ENVELOPE", synthetic_formula) and not semantic_ok(b.validate_projected_envelope_semantics, synthetic_formula),
                all(entry["source_locator"].startswith("plan:synthetic_") for entry in synthetic["entries"]),
            ],
        )

    def test_09_xr009_subject_path_canonicalization(self):
        valid = load("benchmark-result.valid.json")
        alias_dot = copy.deepcopy(valid)
        alias_dot["subject_manifest"][0]["path"] = "a/./b"
        digest_only(alias_dot)
        alias_slashes = copy.deepcopy(valid)
        alias_slashes["subject_manifest"][0]["path"] = "a//b"
        digest_only(alias_slashes)
        trailing = copy.deepcopy(valid)
        trailing["subject_manifest"][0]["path"] = "a/b/"
        digest_only(trailing)
        duplicate_alias = copy.deepcopy(valid)
        duplicate_alias["subject_manifest"] = [
            {"path": "a/b", "sha256": "sha256:" + "1" * 64},
            {"path": "a/./b", "sha256": "sha256:" + "2" * 64},
        ]
        digest_only(duplicate_alias)
        parent = copy.deepcopy(valid)
        parent["subject_manifest"][0]["path"] = "../a"
        digest_only(parent)
        absolute = copy.deepcopy(valid)
        absolute["subject_manifest"][0]["path"] = "/a/b"
        digest_only(absolute)
        windows = copy.deepcopy(valid)
        windows["subject_manifest"][0]["path"] = "C:\\a\\b"
        digest_only(windows)
        reordered_manifest = [
            {"sha256": entry["sha256"], "path": entry["path"]}
            for entry in valid["subject_manifest"]
        ]
        self.check_range(
            81,
            [
                b.canonical_subject_path("format-monograph/scripts/profile_v2_benchmark.py") == "format-monograph/scripts/profile_v2_benchmark.py",
                schema_valid("RESULT", alias_dot) and not semantic_ok(b.validate_benchmark_result_semantics, alias_dot),
                schema_valid("RESULT", alias_slashes) and not semantic_ok(b.validate_benchmark_result_semantics, alias_slashes),
                schema_valid("RESULT", trailing) and not semantic_ok(b.validate_benchmark_result_semantics, trailing),
                schema_valid("RESULT", duplicate_alias) and not semantic_ok(b.validate_benchmark_result_semantics, duplicate_alias),
                not schema_valid("RESULT", parent) and not semantic_ok(b.validate_benchmark_result_semantics, parent),
                not schema_valid("RESULT", absolute) and not semantic_ok(b.validate_benchmark_result_semantics, absolute),
                not schema_valid("RESULT", windows) and not semantic_ok(b.validate_benchmark_result_semantics, windows),
                b.recompute_subject_digest(valid["subject_manifest"]) == b.recompute_subject_digest(reordered_manifest),
                valid["subject_manifest"] == sorted(valid["subject_manifest"], key=lambda entry: entry["path"]),
            ],
        )

    def test_10_baseline_boundary_and_independent_recompute(self):
        config = load("benchmark-config.valid.json")
        valid = load("benchmark-result.valid.json")
        forbidden = {
            "artifact_kind",
            "semantic_fingerprint",
            "input_fingerprints",
            "delivery",
            "final_ready",
            "runtime_eligible",
            "execution_eligibility",
            "evidence_commit",
        }
        roots = set().union(*(set(schema["properties"]) for schema in SCHEMAS.values()))
        source = MOD.read_text()
        py311 = copy.deepcopy(valid)
        py311["environment"]["python_version"] = "3.11.9"
        digest_only(py311)
        refs = []
        def walk(node):
            if isinstance(node, dict):
                for key, value in node.items():
                    if key == "$ref": refs.append(value)
                    walk(value)
            elif isinstance(node, list):
                for value in node: walk(value)
        for schema in SCHEMAS.values(): walk(schema)
        measured = valid["runs"][1:]
        independent_summary = {
            "median_wall_seconds": median([run["timings"]["end_to_end"]["wall_seconds"] for run in measured]),
            "max_wall_seconds": max(run["timings"]["end_to_end"]["wall_seconds"] for run in measured),
            "median_peak_rss_mib": median([run["rss"]["peak_rss_mib"] for run in measured]),
            "max_peak_rss_mib": max(run["rss"]["peak_rss_mib"] for run in measured),
            "median_output_json_bytes": int(median([run["output_json_bytes"] for run in measured])),
            "max_output_json_bytes": max(run["output_json_bytes"] for run in measured),
        }
        self.check_range(
            91,
            [
                all(Draft202012Validator.check_schema(schema) is None for schema in SCHEMAS.values()),
                all(ref.startswith("#") for ref in refs),
                not (forbidden & roots),
                "subprocess" not in source and "run_benchmark" not in source and "def main(" not in source,
                schema_valid("RESULT", valid) and semantic_ok(b.validate_benchmark_result_against_config, valid, config),
                not schema_valid("RESULT", py311),
                config["config_digest"] == independent_digest(config, "config_digest"),
                valid["result_digest"] == independent_digest(valid, "result_digest"),
                valid["summary"] == independent_summary,
                len(ASSERTION_IDS) == 110 and ASSERTION_IDS[0].endswith("001") and ASSERTION_IDS[-1].endswith("110"),
            ],
        )

    def test_11_cross_document_frozen_values(self):
        config = load("benchmark-config.valid.json")
        valid = load("benchmark-result.valid.json")
        wrong_basis = copy.deepcopy(valid)
        wrong_basis["output_json_bytes_basis"] = "other_basis"
        digest_only(wrong_basis)
        wrong_rss = copy.deepcopy(valid)
        wrong_rss["rss_protocol"]["sampling_interval_seconds"] = 0.1
        digest_only(wrong_rss)
        self.check_range(
            101,
            [
                valid["rss_protocol"] == config["rss_protocol"],
                valid["thresholds"] == config["thresholds"],
                valid["output_json_bytes_basis"] == config["output_json_bytes_basis"],
                valid["input_json_bytes_basis"] == config["input_json_bytes_basis"],
                valid["reference_budget"]["limit_hours"] == config["total_reference_budget_hours"],
                valid["parameters"]["generation_seed"] == config["generation"]["generation_seed"],
                valid["benchmark_config_digest"] == config["config_digest"],
                semantic_ok(b.validate_benchmark_result_against_config, valid, config),
                not schema_valid("RESULT", wrong_basis) and not semantic_ok(b.validate_benchmark_result_against_config, wrong_basis, config),
                not schema_valid("RESULT", wrong_rss) and not semantic_ok(b.validate_benchmark_result_against_config, wrong_rss, config),
            ],
        )


if __name__ == "__main__":
    unittest.main()
