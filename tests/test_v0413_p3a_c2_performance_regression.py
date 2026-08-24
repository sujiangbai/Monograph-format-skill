"""Structural C2P-FIX regression checks; this is not a performance campaign."""
from __future__ import annotations

import copy
import hashlib
import json
import sys
import time
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import profile_v2_benchmark_runner as runner
import profile_v2_composer as composer
from profile_v2_canonical import stamp_intent_semantic_fingerprint_v041
from profile_v2_composer import ComposerContractError, ComposerDisabledError


MICRO_CONFIG = {
    "projection_binding": {
        "aggregate_counts": {
            "rule_fragment": 2,
            "binding": 4,
            "key": 2,
            "candidate": 4,
        }
    }
}
GOLDEN = ROOT / "tests" / "fixtures" / "v0413" / "p3a_c2" / "old-h-micro-golden.json"
FORMAL_CONFIG = ROOT / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2" / "benchmark-config.json"


def _digest(value: str) -> str:
    return "sha256:" + hashlib.sha256(value.encode("utf-8")).hexdigest()


def _tool() -> dict[str, str]:
    return {"tool_id": "format-monograph", "version": "0.4.1"}


def _approval_for(report: dict, ordinal: int) -> dict:
    conflict = report["approval_required_conflicts"][ordinal]
    approval = {
        "artifact_kind": "qa-approval-artifact",
        "schema_version": "2.2",
        "registry_contract_version": "2.2",
        "authority_contract_version": "1.0",
        "artifact_id": f"qa-approval-artifact:c2p-{ordinal}",
        "created_by_tool": _tool(),
        "input_fingerprints": [
            {"input_id": "input:c2p-source", "role": "source_document", "fingerprint": report["bindings"]["input_fingerprint"]},
            {"input_id": "input:c2p-structure", "role": "structure", "fingerprint": report["bindings"]["structure_fingerprint"]},
            {"input_id": "input:c2p-report", "role": "conflict_report", "fingerprint": report["semantic_fingerprint"]},
        ],
        "semantic_fingerprint": _digest("unstamped-c2p-approval"),
        "approval_id": f"approval:c2p-{ordinal}",
        "approver": {"actor_id": "actor:c2p", "actor_role": "user"},
        "decision_type": "conflict_resolution",
        "decision": "adopt_proposed",
        "reason": "Structural C2P regression approval.",
        "created_at": "2026-08-24T00:00:00Z",
        "bindings": {
            "input_fingerprint": report["bindings"]["input_fingerprint"],
            "structure_fingerprint": report["bindings"]["structure_fingerprint"],
            "composition_report_fingerprint": report["semantic_fingerprint"],
        },
        "target": {
            "conflict_id": conflict["conflict_id"],
            "proposed_resolution_id": conflict["proposed_resolution_id"],
            "normalized_scope": copy.deepcopy(conflict["key"]["normalized_scope"]),
        },
        "previous_approval_id": None,
    }
    return stamp_intent_semantic_fingerprint_v041(approval)


class C2PPreparedSnapshotRegressionTests(unittest.TestCase):
    """Keep C2P's validation ownership observable without naming private helpers."""

    def _compose(self, assets: list[dict], manifest: dict, *, metrics: bool = True):
        return composer.compose_intent_profile_v041(
            assets,
            manifest,
            input_fingerprint=_digest("c2p-source"),
            structure_fingerprint=_digest("c2p-structure"),
            artifact_id="conflict-report:c2p",
            created_by_tool=_tool(),
            generated_at="2026-08-24T00:00:00Z",
            include_metrics=metrics,
        )

    def test_snapshot_isolated_and_asset_validation_is_single_boundary(self) -> None:
        assets, expected = runner.generate_assets(
            MICRO_CONFIG, "1.0x", "mixed-conflict-approval", 41
        )
        manifest = runner.feature_manifest()
        original = copy.deepcopy(assets)
        calls: list[str] = []
        mutate_once = {"done": False}
        validate = composer.validate_intent_artifact_v041

        def observed_validate(document: dict) -> None:
            calls.append(document.get("artifact_kind", ""))
            if not mutate_once["done"]:
                mutate_once["done"] = True
                assets[0]["activation"] = "rejected"
                assets[0]["rules"][0]["status"] = "rejected"
                assets[0]["rules"][0]["properties"][0]["value"]["value"] = "caller-mutated"
                manifest["features"]["monograph_base_v041"] = False
            validate(document)

        with patch.object(composer, "validate_intent_artifact_v041", side_effect=observed_validate):
            composed = self._compose(assets, manifest)
            approvals = [
                _approval_for(composed.report, ordinal)
                for ordinal in range(len(composed.report["approval_required_conflicts"]))
            ]
            applied = composer.apply_intent_resolutions_v041(
                composed.report,
                approvals,
                runner.feature_manifest(),
                task_id="task:c2p",
                task_fingerprint=_digest("c2p-task"),
                artifact_id="final-execution-profile:c2p",
                created_by_tool=_tool(),
                metrics=composed.metrics,
            )

        self.assertEqual("awaiting_approval", composed.status)
        self.assertEqual("profile_generated", applied.status)
        self.assertIsNotNone(applied.final_profile)
        self.assertEqual(original[0]["semantic_fingerprint"], composed.report["bindings"]["rule_asset_fingerprints"][0])
        self.assertEqual("rejected", assets[0]["activation"])
        self.assertEqual("rejected", assets[0]["rules"][0]["status"])
        self.assertEqual("caller-mutated", assets[0]["rules"][0]["properties"][0]["value"]["value"])
        self.assertFalse(manifest["features"]["monograph_base_v041"])
        self.assertEqual(len(original), calls.count("layered-rule-asset"))
        self.assertGreaterEqual(calls.count("conflict-report"), 2)
        self.assertGreaterEqual(calls.count("final-execution-profile"), 1)
        self.assertEqual(expected["rule_fragment"], composed.metrics.input_asset_count)

    def test_direct_composer_remains_fail_closed_without_runner(self) -> None:
        assets, _ = runner.generate_assets(MICRO_CONFIG, "1.0x", "disjoint", 41)
        disabled = runner.feature_manifest()
        disabled["features"]["monograph_base_v041"] = False
        disabled = stamp_intent_semantic_fingerprint_v041(disabled)
        with self.assertRaises(ComposerDisabledError):
            self._compose(assets, disabled, metrics=False)

    def test_content_bound_prepared_snapshot_rejects_post_validation_tamper(self) -> None:
        assets, _ = runner.generate_assets(MICRO_CONFIG, "1.0x", "disjoint", 41)
        adapter = composer._intent_adapter_v041()
        prepared = composer._prepare_intent_snapshot_v041(
            assets, runner.feature_manifest(), adapter
        )
        prepared.rule_assets[0]["rules"][0]["properties"][0]["value"]["value"] = "tampered-after-validation"
        with self.assertRaises(ComposerContractError):
            composer._build_report_documents(
                prepared.rule_assets,
                prepared.feature_manifest,
                adapter=adapter,
                input_fingerprint=_digest("c2p-source"),
                structure_fingerprint=_digest("c2p-structure"),
                artifact_id="conflict-report:c2p-tamper",
                created_by_tool=_tool(),
                generated_at="2026-08-24T00:00:00Z",
                intent_contract=True,
                prepared_intent=prepared,
            )

    def test_committed_old_h_micro_golden_matches_h2(self) -> None:
        """The expected values were generated once by H, never by this implementation."""
        golden = json.loads(GOLDEN.read_text(encoding="utf-8"))
        self.assertEqual("3c8d40d6c04113845ad7fa6a776cb500e3834c94", golden["source_commit"])
        self.assertEqual(
            "canonical_composition_report_plus_final_profile_if_present_v1",
            golden["canonical_basis"],
        )
        request = golden["input"]
        config = {"projection_binding": {"aggregate_counts": request["aggregate_counts"]}}
        cell = runner.compose_cell(
            config,
            scale_id=request["scale_id"],
            scenario_id=request["scenario_id"],
            generation_seed=request["generation_seed"],
            permutation_seed=request["permutation_seed"],
        )
        report = cell["report"]
        final = cell["final_profile"]
        expected = golden["expected"]
        coverage = report["coverage_evidence"]
        actual = {
            "terminal": cell["status"],
            "report_canonical_sha256": "sha256:" + hashlib.sha256(runner.canonical_json_bytes(report)).hexdigest(),
            "final_canonical_sha256": None if final is None else "sha256:" + hashlib.sha256(runner.canonical_json_bytes(final)).hexdigest(),
            "report_semantic_fingerprint": report["semantic_fingerprint"],
            "final_semantic_fingerprint": None if final is None else final["semantic_fingerprint"],
            "metrics": cell["metrics"].as_dict(),
            "coverage": {
                "expected_count": coverage["expected_binding_count"],
                "consumed_count": coverage["consumed_binding_count"],
                "expected_digest": coverage["expected_inventory_digest"],
                "consumed_digest": coverage["consumed_inventory_digest"],
            },
            "stable_ids": {
                "candidate_group_ids": sorted(item["candidate_group_id"] for item in report["candidate_groups"]),
                "blocker_ids": sorted(item["blocker_id"] for item in report["unresolvable_blockers"]),
            },
        }
        self.assertEqual(expected, actual)

    def test_frozen_formal_shape_is_structural_invariant_only(self) -> None:
        """Structural invariant only: not a Pilot, campaign, or 60/120-second gate."""
        config = json.loads(FORMAL_CONFIG.read_text(encoding="utf-8"))
        assets, expected = runner.generate_assets(
            config, "0.5x", "dense-crossing", 41, permutation_seed=None
        )
        calls: list[str] = []
        validate = composer.validate_intent_artifact_v041

        def observed_validate(document: dict) -> None:
            calls.append(document.get("artifact_kind", ""))
            validate(document)

        started = time.monotonic()
        with patch.object(composer, "validate_intent_artifact_v041", side_effect=observed_validate):
            composed = self._compose(assets, runner.feature_manifest())
        elapsed = time.monotonic() - started
        self.assertGreater(elapsed, 0.0)
        self.assertEqual("unresolvable", composed.status)
        self.assertIsNone(composed.final_profile)
        self.assertEqual(expected["rule_fragment"], composed.metrics.input_asset_count)
        self.assertEqual(expected["binding"], composed.metrics.input_rule_count)
        self.assertEqual(expected["binding"], composed.metrics.input_binding_count)
        self.assertEqual(expected["key"], composed.metrics.expected_key_count)
        self.assertEqual(expected["candidate"], composed.metrics.candidate_count)
        coverage = composed.report["coverage_evidence"]
        self.assertEqual(coverage["expected_binding_count"], coverage["consumed_binding_count"])
        self.assertTrue(coverage["expected_inventory_digest"].startswith("sha256:"))
        self.assertTrue(coverage["consumed_inventory_digest"].startswith("sha256:"))
        self.assertEqual(len(assets), calls.count("layered-rule-asset"))
        self.assertGreaterEqual(calls.count("conflict-report"), 1)
        self.assertEqual(1, len(composed.report["unresolvable_blockers"]))
        stable_ids = [item["blocker_id"] for item in composed.report["unresolvable_blockers"]]
        self.assertEqual(len(stable_ids), len(set(stable_ids)))

    def test_runner_records_no_duplicate_validation_stage(self) -> None:
        payload = runner.worker_payload(
            {
                "config": MICRO_CONFIG,
                "scale_id": "1.0x",
                "scenario_id": "disjoint",
                "generation_seed": 41,
                "permutation_seed": None,
            }
        )
        self.assertEqual("ok", payload["status"])
        self.assertIsNone(payload["stage_seconds"]["schema_registry_validation"])
        config = json.loads((ROOT / "tests" / "fixtures" / "v0412" / "p3a_c2" / "benchmark-config.valid.json").read_text(encoding="utf-8"))
        manifest = [{"path": "format-monograph/scripts/profile_v2_composer.py", "sha256": _digest("subject")}]
        result = runner.build_result(
            cell=None,
            supervised={
                "status": "completed",
                "worker": payload,
                "rss": {"status": "available", "baseline_rss_mib": 1.0, "peak_rss_mib": 1.1, "delta_rss_mib": 0.1},
            },
            config=config,
            subject_commit="a" * 40,
            subject_manifest=manifest,
            parameters={
                "measurement_kind": "coverage",
                "scale_id": "1.0x",
                "scenario_id": "disjoint",
                "generation_seed": 41,
                "permutation_seed": None,
            },
            elapsed_seconds=1.0,
            environment={"os_family": "windows", "os_build_class": "unspecified", "python_version": "3.12.13", "cpu_architecture": "x86_64", "logical_cpu_count": 1, "ram_tier": "le_8gib"},
        )
        timings = result["runs"][0]["timings"]
        self.assertEqual("not_applicable", timings["schema_registry_validation"]["status"])
        self.assertEqual("measured", timings["compose"]["status"])
        self.assertEqual("measured", timings["end_to_end"]["status"])
