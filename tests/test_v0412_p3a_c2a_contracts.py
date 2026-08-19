import copy
import hashlib
import json
import re
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator


ROOT = Path(__file__).resolve().parents[1]
BENCHMARKS = ROOT / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2"
FIXTURES = ROOT / "tests" / "fixtures" / "v0412" / "p3a_c2a"
SCHEMA_FILES = {
    "ENVELOPE": "projected-envelope.schema.json",
    "CONFIG": "benchmark-config.schema.json",
    "RESULT": "benchmark-result.schema.json",
}
RANGES = {"ENVELOPE": 18, "CONFIG": 20, "RESULT": 26, "BOUNDARY": 16}
ASSERTION_IDS = tuple(
    f"T412-C2A-{family}-{index:03d}"
    for family, count in RANGES.items()
    for index in range(1, count + 1)
)


def load_json(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def canonical_result_digest(document: dict) -> str:
    payload = copy.deepcopy(document)
    payload.pop("result_digest", None)
    data = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def result_subject_status_is_consistent(document: dict) -> bool:
    status = document["subject_digest_status"]
    same_digest = status["observed_subject_digest"] == document["benchmark_subject_digest"]
    if status["state"] == "current":
        return same_digest and not status["revalidation_required"]
    return not same_digest and status["revalidation_required"]


def iter_refs(node: object):
    if isinstance(node, dict):
        for key, value in node.items():
            if key == "$ref":
                yield value
            yield from iter_refs(value)
    elif isinstance(node, list):
        for value in node:
            yield from iter_refs(value)


class C2AContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.schemas = {
            family: json.loads((BENCHMARKS / filename).read_text(encoding="utf-8"))
            for family, filename in SCHEMA_FILES.items()
        }
        self.validators = {
            family: Draft202012Validator(schema)
            for family, schema in self.schemas.items()
        }
        self.executed: set[str] = set()

    def tearDown(self) -> None:
        expected = set(ASSERTION_IDS)
        self.assertTrue(self.executed <= expected)

    def _record(self, assertion_id: str, condition: bool) -> None:
        self.assertIn(assertion_id, ASSERTION_IDS)
        self.assertNotIn(assertion_id, self.executed)
        self.executed.add(assertion_id)
        self.assertTrue(condition, assertion_id)

    def _valid(self, family: str, document: dict) -> bool:
        return not list(self.validators[family].iter_errors(document))

    def _invalid(self, family: str, document: dict) -> bool:
        return bool(list(self.validators[family].iter_errors(document)))

    def _check_family(self, family: str, checks: list[bool]) -> None:
        self.assertEqual(RANGES[family], len(checks))
        for index, condition in enumerate(checks, 1):
            self._record(f"T412-C2A-{family}-{index:03d}", condition)

    def test_envelope_contract(self) -> None:
        valid = load_json("projected-envelope.valid.json")
        invalid_fixture = load_json("projected-envelope.invalid.json")
        protected = valid["entries"][1]
        summary_bad = copy.deepcopy(valid)
        summary_bad["decision_population_summary"]["v041_primary"] = 149
        strategy_bad = copy.deepcopy(valid)
        strategy_bad["compose_projection_strategy"] = "multi_probe"
        protected_owner_bad = copy.deepcopy(valid)
        protected_owner_bad["entries"][1]["implementation_owner"] = "p3b_b"
        protected_kind_bad = copy.deepcopy(valid)
        protected_kind_bad["entries"][1]["requirement_kind"] = "base_appearance"
        protected_count_bad = copy.deepcopy(valid)
        protected_count_bad["entries"][1]["projected_counts"]["binding"] = 1
        protected_load_bad = copy.deepcopy(valid)
        protected_load_bad["entries"][1]["zero_load_reason"] = "none"
        artifact_bad = copy.deepcopy(valid)
        artifact_bad["artifact_kind"] = "conflict-report"
        property_bad = copy.deepcopy(valid)
        property_bad["entries"][0]["property_id"] = "format.title"
        missing_bad = copy.deepcopy(valid)
        del missing_bad["entries"][0]["derivation_formula"]
        unknown_bad = copy.deepcopy(valid)
        unknown_bad["entries"][0]["unexpected"] = True
        blocked = copy.deepcopy(valid)
        blocked["entries"][0].update({"blocked_projection": True, "zero_load_reason": "none", "scope_topology": "unmodeled", "conflict_approval_assumption": "unmodeled"})
        blocked_bad = copy.deepcopy(blocked)
        blocked_bad["entries"][0]["scope_topology"] = "single_core_probe"
        checks = [
            self._valid("ENVELOPE", valid),
            self.schemas["ENVELOPE"]["$schema"] == "https://json-schema.org/draft/2020-12/schema",
            self.schemas["ENVELOPE"]["properties"]["document_kind"]["const"] == "p3a_c2_projected_envelope",
            self.schemas["ENVELOPE"]["properties"]["projection_status"]["const"] == "planning_projection",
            self.schemas["ENVELOPE"]["properties"]["compose_projection_strategy"]["const"] == "single_core_probe",
            valid["decision_population_summary"] == {"v041_primary": 150, "v042_protected": 20, "v043_protected": 1},
            protected["projected_counts"] == {"rule_fragment": 0, "binding": 0, "key": 0, "candidate": 0},
            self._invalid("ENVELOPE", summary_bad),
            self._invalid("ENVELOPE", strategy_bad),
            self._invalid("ENVELOPE", protected_owner_bad),
            self._invalid("ENVELOPE", protected_kind_bad),
            self._invalid("ENVELOPE", protected_count_bad),
            self._invalid("ENVELOPE", protected_load_bad),
            self._valid("ENVELOPE", blocked),
            self._invalid("ENVELOPE", blocked_bad),
            self._invalid("ENVELOPE", artifact_bad),
            self._invalid("ENVELOPE", property_bad),
            self._invalid("ENVELOPE", missing_bad) and self._invalid("ENVELOPE", unknown_bad) and self._invalid("ENVELOPE", invalid_fixture),
        ]
        self._check_family("ENVELOPE", checks)

    def test_config_contract(self) -> None:
        valid = load_json("benchmark-config.valid.json")
        invalid_fixture = load_json("benchmark-config.invalid.json")
        missing_scale = copy.deepcopy(valid)
        missing_scale["scales"].pop()
        wrong_scale = copy.deepcopy(valid)
        wrong_scale["scales"][2]["factor"] = 1.0
        wrong_matrix = copy.deepcopy(valid)
        wrong_matrix["scales"][0]["scenario_ids"].pop()
        wrong_disjoint = copy.deepcopy(valid)
        wrong_disjoint["scenarios"][0]["final_requirement"] = "forbidden"
        wrong_dense = copy.deepcopy(valid)
        wrong_dense["scenarios"][2]["post_approval_terminal"] = "final"
        wrong_mixed = copy.deepcopy(valid)
        wrong_mixed["scenarios"][3]["pre_approval_terminal"] = "final"
        wrong_warmup = copy.deepcopy(valid)
        wrong_warmup["repetitions"]["warmup_runs"] = 0
        wrong_measured = copy.deepcopy(valid)
        wrong_measured["repetitions"]["measured_runs"] = 2
        wrong_rss = copy.deepcopy(valid)
        wrong_rss["rss_protocol"] = "tracemalloc_as_rss"
        wrong_limit = copy.deepcopy(valid)
        wrong_limit["thresholds"]["scale_limits"][3]["max_peak_rss_mib"] = 512
        wrong_ratio = copy.deepcopy(valid)
        wrong_ratio["thresholds"]["wall_median_ratio_2x_to_1x"] = 7
        wrong_budget = copy.deepcopy(valid)
        wrong_budget["total_reference_budget_hours"] = 3.1
        artifact_bad = copy.deepcopy(valid)
        artifact_bad["semantic_fingerprint"] = "sha256:" + "a" * 64
        checks = [
            self._valid("CONFIG", valid),
            self.schemas["CONFIG"]["properties"]["document_kind"]["const"] == "p3a_c2_benchmark_config",
            valid["headroom_factor"] == 1.5 and valid["support_ceiling"] == 2.0,
            [item["scale_id"] for item in valid["scales"]] == ["0.5x", "1.0x", "1.5x", "2.0x"],
            all(len(item["scenario_ids"]) == 4 for item in valid["scales"]),
            [item["scenario_id"] for item in valid["scenarios"]] == ["disjoint", "subset-chain", "dense-crossing", "mixed-conflict-approval"],
            valid["scenarios"][0]["post_approval_terminal"] == "final",
            valid["scenarios"][2]["final_requirement"] == "forbidden",
            valid["scenarios"][3]["pre_approval_terminal"] == "awaiting_approval",
            valid["scenarios"][3]["post_approval_terminal"] == "final",
            valid["generation"]["rounding_policy"] == "round_half_even",
            valid["repetitions"]["warmup_runs"] == 1 and valid["repetitions"]["measured_runs"] == 3,
            valid["rss_protocol"] == "external_supervisor_child_peak_working_set",
            valid["thresholds"]["scale_limits"][0]["max_wall_seconds"] == 60,
            valid["thresholds"]["scale_limits"][3] == {"scale_id": "2.0x", "max_wall_seconds": 120, "max_peak_rss_mib": 1024},
            self._invalid("CONFIG", missing_scale) and self._invalid("CONFIG", wrong_scale),
            self._invalid("CONFIG", wrong_matrix) and self._invalid("CONFIG", wrong_disjoint),
            self._invalid("CONFIG", wrong_dense) and self._invalid("CONFIG", wrong_mixed),
            self._invalid("CONFIG", wrong_warmup) and self._invalid("CONFIG", wrong_measured),
            self._invalid("CONFIG", wrong_rss) and self._invalid("CONFIG", wrong_limit) and self._invalid("CONFIG", wrong_ratio) and self._invalid("CONFIG", wrong_budget) and self._invalid("CONFIG", artifact_bad) and self._invalid("CONFIG", invalid_fixture),
        ]
        self._check_family("CONFIG", checks)

    def test_result_contract(self) -> None:
        valid = load_json("benchmark-result.valid.json")
        invalid_fixture = load_json("benchmark-result.invalid.json")
        dense = copy.deepcopy(valid)
        dense["parameters"]["scenario_id"] = "dense-crossing"
        dense.update({"expected_terminal_state": "unresolvable", "actual_terminal_state": "unresolvable", "final_present": False, "final_bytes": None, "final_fingerprint": None, "final_determinism": "not_applicable"})
        for run in dense["raw_runs"]:
            run["terminal_state"] = "unresolvable"
        dense_bad_final = copy.deepcopy(dense)
        dense_bad_final["final_present"] = True
        dense_bad_final["final_bytes"] = 1
        dense_bad_final["final_fingerprint"] = "sha256:" + "e" * 64
        dense_bad_final["final_determinism"] = "matched"
        complete_bad = copy.deepcopy(valid)
        complete_bad.update({"final_present": False, "final_bytes": None, "final_fingerprint": None, "final_determinism": "not_applicable"})
        pre_approval = copy.deepcopy(valid)
        pre_approval["parameters"]["scenario_id"] = "mixed-conflict-approval"
        pre_approval.update({"resolution_phase": "pre_approval", "expected_terminal_state": "awaiting_approval", "actual_terminal_state": "awaiting_approval", "final_present": False, "final_bytes": None, "final_fingerprint": None, "final_determinism": "not_applicable"})
        for run in pre_approval["raw_runs"]:
            run["terminal_state"] = "awaiting_approval"
        post_approval = copy.deepcopy(valid)
        post_approval["parameters"]["scenario_id"] = "mixed-conflict-approval"
        post_approval["resolution_phase"] = "post_deterministic_approval"
        bad_commit = copy.deepcopy(valid)
        bad_commit["benchmark_subject_commit"] = "not-a-commit"
        bad_digest = copy.deepcopy(valid)
        bad_digest["benchmark_subject_digest"] = "sha256:" + "Z" * 64
        bad_evidence_commit = copy.deepcopy(valid)
        bad_evidence_commit["evidence_commit"] = "short"
        stale = copy.deepcopy(valid)
        stale["subject_digest_status"].update({"state": "stale", "observed_subject_digest": "sha256:" + "f" * 64, "revalidation_required": True})
        stale_bad = copy.deepcopy(stale)
        stale_bad["subject_digest_status"]["revalidation_required"] = False
        current_mismatch = copy.deepcopy(valid)
        current_mismatch["subject_digest_status"]["observed_subject_digest"] = "sha256:" + "f" * 64
        path_bad = copy.deepcopy(valid)
        path_bad["command_template"] = "C:\\\\private\\\\benchmark.exe"
        host_bad = copy.deepcopy(valid)
        host_bad["environment"]["hostname"] = "private-host"
        artifact_bad = copy.deepcopy(valid)
        artifact_bad["delivery"] = True
        missing_bad = copy.deepcopy(valid)
        del missing_bad["raw_runs"]
        raw_bad = copy.deepcopy(valid)
        raw_bad["raw_runs"][0]["run_phase"] = "measured"
        checks = [
            self._valid("RESULT", valid),
            valid["document_kind"] == "p3a_c2_benchmark_result" and valid["evidence_kind"] == "non_artifact_benchmark",
            valid["evidence_commit"] is None,
            bool(re.fullmatch(r"[a-f0-9]{40}", valid["benchmark_subject_commit"])),
            valid["result_digest_basis"] == "canonical_json_excluding_result_digest",
            valid["environment"]["cpu_architecture"] == "x86_64" and "hostname" not in valid["environment"],
            valid["parameters"]["rss_protocol"] == "external_supervisor_child_peak_working_set",
            len(valid["raw_runs"]) == 4 and valid["raw_runs"][0] == {**valid["raw_runs"][0], "run_index": 1, "run_phase": "warmup"},
            [(run["run_index"], run["run_phase"]) for run in valid["raw_runs"][1:]] == [(2, "measured"), (3, "measured"), (4, "measured")],
            set(valid["summary"]) == {"median_wall_seconds", "max_wall_seconds", "median_peak_rss_mib", "max_peak_rss_mib", "median_output_json_bytes", "max_output_json_bytes"},
            valid["report_determinism"] == {"canonical_bytes_equal": True, "fingerprint_equal": True},
            valid["final_present"] and valid["final_determinism"] == "matched",
            result_subject_status_is_consistent(valid),
            self._valid("RESULT", dense),
            self._invalid("RESULT", dense_bad_final),
            self._invalid("RESULT", complete_bad),
            self._valid("RESULT", pre_approval),
            self._valid("RESULT", post_approval),
            self._invalid("RESULT", bad_commit),
            self._invalid("RESULT", bad_digest) and self._invalid("RESULT", bad_evidence_commit),
            self._valid("RESULT", stale) and result_subject_status_is_consistent(stale) and self._invalid("RESULT", stale_bad) and self._valid("RESULT", current_mismatch) and not result_subject_status_is_consistent(current_mismatch),
            self._invalid("RESULT", path_bad) and self._invalid("RESULT", host_bad),
            self._invalid("RESULT", artifact_bad) and self._invalid("RESULT", missing_bad) and self._invalid("RESULT", invalid_fixture),
            self._invalid("RESULT", raw_bad),
            valid["result_digest"] == canonical_result_digest(valid) and canonical_result_digest(valid) == canonical_result_digest({**valid, "result_digest": "sha256:" + "e" * 64}),
            canonical_result_digest(valid) != canonical_result_digest({**valid, "benchmark_subject_digest": "sha256:" + "e" * 64}),
        ]
        self._check_family("RESULT", checks)

    def test_boundary_and_offline_contracts(self) -> None:
        forbidden = {
            "artifact_kind", "schema_version", "semantic_fingerprint", "input_fingerprints",
            "delivery", "final_ready", "runtime_eligible", "execution_eligibility",
        }
        schema_ids = [schema["$id"] for schema in self.schemas.values()]
        isolated = all(all(ref.startswith("#") for ref in iter_refs(schema)) for schema in self.schemas.values())
        root_properties = set().union(*(set(schema["properties"]) for schema in self.schemas.values()))
        checks = [
            len(self.schemas) == 3,
            len(schema_ids) == len(set(schema_ids)),
            all(schema["$schema"] == "https://json-schema.org/draft/2020-12/schema" for schema in self.schemas.values()),
            all(not list(Draft202012Validator.check_schema(schema) or []) for schema in self.schemas.values()),
            isolated,
            not (forbidden & root_properties),
            all("artifact_kind" not in schema.get("required", []) for schema in self.schemas.values()),
            all("semantic_fingerprint" not in schema.get("required", []) for schema in self.schemas.values()),
            "result_digest" in self.schemas["RESULT"]["properties"],
            "result_digest" not in self.schemas["ENVELOPE"]["properties"],
            "result_digest" not in self.schemas["CONFIG"]["properties"],
            "benchmark_subject_commit" not in self.schemas["ENVELOPE"]["properties"],
            "benchmark_subject_commit" not in self.schemas["CONFIG"]["properties"],
            all("tracemalloc_as_rss" not in json.dumps(schema) for schema in self.schemas.values()),
            all("execution-evidence-artifact" not in json.dumps(schema) for schema in self.schemas.values()),
            all("final_ready_eligible" not in json.dumps(schema) for schema in self.schemas.values()),
        ]
        self._check_family("BOUNDARY", checks)


if __name__ == "__main__":
    unittest.main()
