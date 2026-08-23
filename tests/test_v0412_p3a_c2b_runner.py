"""T412-P3AC-BENCH-C2B-001..032 C2B runner behavior contracts."""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import io
import json
import tempfile
import time
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
SPEC = importlib.util.spec_from_file_location("c2b_runner", SCRIPTS / "profile_v2_benchmark_runner.py")
runner = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(runner)

CHECK_IDS = tuple("T412-P3AC-BENCH-C2B-%03d" % number for number in range(1, 33))
ASSERTION_ID_MAP = {
    "schema_first_context": CHECK_IDS[0:3],
    "subject_manifest": CHECK_IDS[3:6],
    "workload_and_source_keys": CHECK_IDS[6:11],
    "four_scenario_smoke": CHECK_IDS[11:15],
    "seed_and_permutation": CHECK_IDS[15:17],
    "worker_supervisor": CHECK_IDS[17:22],
    "campaign_budget": CHECK_IDS[22:23],
    "atomic_cache_resume": CHECK_IDS[23:29],
    "suite_closure_and_runtime_boundary": CHECK_IDS[29:32],
}
MICRO_CONFIG = {
    "projection_binding": {
        "aggregate_counts": {"rule_fragment": 2, "binding": 4, "key": 2, "candidate": 4}
    }
}


class _Input(io.BytesIO):
    def close(self) -> None:
        self.closed_by_runner = True


class _Process:
    def __init__(self, lines: bytes, returncode: int = 0) -> None:
        self.stdin = _Input()
        self.stdout = io.BytesIO(lines)
        self.stderr = io.BytesIO()
        self.returncode = returncode
        self.pid = 1
        self.killed = False

    def poll(self):
        return self.returncode

    def wait(self):
        return self.returncode

    def kill(self):
        self.killed = True
        self.returncode = -9


class _DelayedEmpty:
    def readline(self) -> bytes:
        time.sleep(0.02)
        return b""


def _sha(text: str) -> str:
    return "sha256:" + hashlib.sha256(text.encode("ascii")).hexdigest()


def _fake_process(lines: bytes, returncode: int = 0):
    return lambda *args, **kwargs: _Process(lines, returncode)


class C2BRunnerTests(unittest.TestCase):
    seen: set[str] = set()

    @classmethod
    def tearDownClass(cls) -> None:
        assert len(CHECK_IDS) == 32 and len(set(CHECK_IDS)) == 32

    def check(self, number: int, condition: bool) -> None:
        identifier = CHECK_IDS[number - 1]
        self.assertNotIn(identifier, self.seen)
        self.seen.add(identifier)
        self.assertTrue(condition, identifier)

    def test_001_003_schema_first_campaign_context(self) -> None:
        valid_config = {"config_digest": "sha256:" + "0" * 64}
        valid_envelope = {"envelope_digest": "sha256:" + "1" * 64}
        with patch.object(runner, "validate_benchmark_campaign_context", return_value={"ok": True}) as campaign:
            self.check(1, runner.validate_campaign_inputs(valid_config, valid_envelope) == {"ok": True})
            campaign.assert_called_once_with(valid_config, valid_envelope)
        with patch.object(runner, "validate_benchmark_campaign_context", side_effect=runner.BenchmarkContractError("schema")):
            with self.assertRaises(runner.BenchmarkContractError):
                runner.validate_campaign_inputs(valid_config, valid_envelope)
            self.check(2, True)
        self.check(3, "validate_benchmark_campaign_context" in (SCRIPTS / "profile_v2_benchmark_runner.py").read_text(encoding="utf-8"))

    def test_004_006_subject_manifest_contract(self) -> None:
        class Completed:
            def __init__(self, path: str) -> None:
                self.returncode = 0
                self.stdout = path.encode("ascii")
        with patch.object(runner.subprocess, "run", side_effect=lambda args, **kwargs: Completed(args[-1])):
            manifest = runner.build_subject_manifest("a" * 40, repository=ROOT)
        self.check(4, [entry["path"] for entry in manifest] == sorted(runner.SUBJECT_PATHS))
        self.check(5, len(manifest) == len(set(runner.SUBJECT_PATHS)) and runner.recompute_subject_digest(manifest).startswith("sha256:"))
        with self.assertRaises(runner.BenchmarkRunnerError):
            runner.build_subject_manifest("A" * 40, repository=ROOT)
        self.check(6, True)

    def test_007_011_workload_and_normalized_source_key_contract(self) -> None:
        self.check(7, runner.scaled_workload(MICRO_CONFIG, "1.0x") == MICRO_CONFIG["projection_binding"]["aggregate_counts"])
        half = runner.scaled_workload({"projection_binding": {"aggregate_counts": {"rule_fragment": 3, "binding": 3, "key": 3, "candidate": 3}}}, "0.5x")
        self.check(8, half == {"rule_fragment": 2, "binding": 2, "key": 2, "candidate": 2})
        with self.assertRaises(runner.BenchmarkRunnerError):
            runner.scaled_workload({"projection_binding": {"aggregate_counts": {"rule_fragment": 1, "binding": 2, "key": 1, "candidate": 1}}}, "1.0x")
        self.check(9, True)
        assets, work = runner.generate_assets(MICRO_CONFIG, "1.0x", "dense-crossing", 1)
        self.check(10, runner.normalized_source_key_count(assets) == work["key"] == 2)
        reversed_scope = copy.deepcopy(assets[0]["rules"][0]["scope"])
        reversed_scope["selectors"][0]["selector_ids"].reverse()
        self.check(11, runner.normalize_scope(reversed_scope)["scope_id"] == runner.normalize_scope(assets[0]["rules"][0]["scope"])["scope_id"])

    def test_012_015_four_real_micro_scenario_smokes_once_each(self) -> None:
        # Exactly one child compose/apply call per scenario.  The explicit
        # table is independent of the generator's workload helper.
        results = {
            scenario: runner.supervise_worker(
                {"config": MICRO_CONFIG, "scale_id": "1.0x", "scenario_id": scenario, "generation_seed": 11},
                timeout_seconds=30.0, python_executable=sys.executable,
            )
            for scenario in runner.SCENARIOS
        }
        for number, scenario in enumerate(runner.SCENARIOS, 12):
            result = results[scenario]
            expected = runner.MICRO_SCENARIO_EXPECTATIONS[scenario]
            worker = result.get("worker", {})
            metrics = worker.get("metrics", {})
            report_counts = worker.get("report_counts", {})
            self.check(number, (
                result["status"] in {"completed", "rss_unavailable"}
                and metrics.get("expected_key_count") == expected["source_keys"]
                and worker.get("proposal_status") == expected["terminal"]
                and report_counts.get("conflicts") == expected["conflicts"]
                and report_counts.get("proposals") == expected["proposals"]
                and report_counts.get("blockers") == expected["blockers"]
                and (worker.get("final_present") is False) == (scenario == "dense-crossing")
            ))

    def test_016_017_permutation_determinism_without_extra_compose(self) -> None:
        normal, _ = runner.generate_assets(MICRO_CONFIG, "1.0x", "mixed-conflict-approval", 5)
        permuted, _ = runner.generate_assets(MICRO_CONFIG, "1.0x", "mixed-conflict-approval", 5, permuted=True)
        self.check(16, runner.normalized_source_key_count(normal) == runner.normalized_source_key_count(permuted) == 2)
        self.check(17, sorted(item["semantic_fingerprint"] for item in normal) == sorted(item["semantic_fingerprint"] for item in permuted))

    def test_018_022_supervisor_statuses(self) -> None:
        response = json.dumps({"status": "ok", "metrics": {}, "final_present": False}).encode("utf-8")
        completed = runner.supervise_worker({}, timeout_seconds=0.1, process_factory=_fake_process(b"READY\n" + response + b"\n"), sampler=lambda pid: 1024)
        self.check(18, completed["status"] == "completed" and completed["rss"]["status"] == "available")
        delayed = _Process(b"")
        delayed.stdout = _DelayedEmpty()
        timeout = runner.supervise_worker({}, timeout_seconds=0.001, process_factory=lambda *args, **kwargs: delayed)
        self.check(19, timeout["status"] == "timeout")
        crash = runner.supervise_worker({}, timeout_seconds=0.1, process_factory=_fake_process(b"BROKEN\n", 1))
        self.check(20, crash["status"] == "process_crash")
        contract = runner.supervise_worker({}, timeout_seconds=0.1, process_factory=_fake_process(b"READY\n{\"status\":\"error\"}\n"), sampler=lambda pid: 1024)
        self.check(21, contract["status"] == "contract_error")
        unavailable = runner.supervise_worker({}, timeout_seconds=0.1, process_factory=_fake_process(b"READY\n" + response + b"\n"), sampler=lambda pid: (_ for _ in ()).throw(OSError("rss")))
        self.check(22, unavailable["status"] == "rss_unavailable")

    def test_023_budget(self) -> None:
        with self.assertRaises(runner.BenchmarkRunnerError):
            runner.enforce_campaign_budget(0.0, limit_seconds=1.0, clock=lambda: 1.1)
        self.check(23, True)

    def test_024_029_atomic_cache_and_resume_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            result = {"parameters": {"measurement_kind": "coverage", "scale_id": "0.5x", "scenario_id": "disjoint", "permutation_seed": None}}
            path = runner.atomic_write_result(directory, result)
            self.check(24, path.exists() and not list(directory.glob("*.tmp")))
            with self.assertRaises(runner.BenchmarkRunnerError):
                runner.atomic_write_result(directory, result)
            self.check(25, True)
            with self.assertRaises(runner.BenchmarkRunnerError):
                runner.load_cached_result(directory / "missing.tmp", {}, {})
            self.check(26, True)
            (directory / "corrupt.json").write_text("{", encoding="utf-8")
            with self.assertRaises(runner.BenchmarkRunnerError):
                runner.load_cached_result(directory / "corrupt.json", {}, {})
            self.check(27, True)
            with patch.object(runner, "validate_benchmark_result_context", return_value={}), patch.object(runner, "logical_key", return_value=path.stem):
                result["benchmark_subject_digest"] = _sha("subject")
                path.write_text(json.dumps(result), encoding="utf-8")
                self.check(28, runner.load_cached_result(path, {}, {}, subject_digest=_sha("subject"))["benchmark_subject_digest"] == _sha("subject"))
                with self.assertRaises(runner.BenchmarkRunnerError):
                    runner.load_cached_result(path, {}, {}, subject_digest=_sha("other"))
                self.check(29, True)

    def test_030_032_closure_and_default_boundary(self) -> None:
        with patch.object(runner, "validate_complete_benchmark_suite", return_value={"overall_gate": "stop"}) as close:
            self.check(30, runner.close_campaign([], {}, {}) == {"overall_gate": "stop"})
            close.assert_called_once()
        self.check(31, "validate_complete_benchmark_suite" in (SCRIPTS / "profile_v2_benchmark_runner.py").read_text(encoding="utf-8"))
        runtime = (ROOT / "format-monograph" / "scripts" / "run_monograph.py").read_text(encoding="utf-8")
        self.check(32, "profile_v2_benchmark_runner" not in runtime and "profile_v2_composer" not in runtime and set().union(*map(set, ASSERTION_ID_MAP.values())) == set(CHECK_IDS))


if __name__ == "__main__":
    unittest.main()
