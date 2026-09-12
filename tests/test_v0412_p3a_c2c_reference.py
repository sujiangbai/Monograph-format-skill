"""H1c oracle for formal C2C planning inputs; it never runs a campaign."""
from __future__ import annotations

import ast
import copy
import hashlib
import inspect
import io
import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "format-monograph" / "scripts"))

from profile_v2_benchmark import (  # noqa: E402
    BenchmarkContractError,
    aggregate_projected_envelope_counts,
    canonical_json_bytes,
    canonical_sha256,
    recompute_config_digest,
    recompute_envelope_digest,
    recompute_result_digest,
    recompute_subject_digest,
    validate_complete_benchmark_suite,
    validate_benchmark_config_against_envelope,
    validate_benchmark_config_semantics,
    validate_benchmark_result_context,
    validate_projected_envelope_semantics,
)
from profile_v2_benchmark_runner import BenchmarkRunnerError, SUBJECT_PATHS, logical_key, scan_cache  # noqa: E402
import profile_v2_benchmark_reference as reference  # noqa: E402

BASE = ROOT / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2"
FIXTURES = ROOT / "tests" / "fixtures" / "v0412" / "p3a_c2"
PLAN_PATHS = (
    "docs/plans/v0.4.0-profile-system-and-long-document-reliability.md",
    "docs/plans/v0.4.1-profile-foundation-and-grouped-execution.md",
    "docs/plans/v0.4.1.2-p3-capability-and-asset-freeze.md",
)
FUTURE_42 = {
    *(f"V040-G-{number:03d}" for number in range(1, 6)),
    *(f"V040-L-{number:03d}" for number in range(6, 8)),
    "V040-N-005",
    *(f"V040-Q-{number:03d}" for number in range(1, 6)),
    "V040-R-005",
    *(f"V040-S-{number:03d}" for number in range(2, 6)),
    "V040-W-006", "V040-Y-004",
}
FUTURE_43 = {"V040-Z-008"}
DECISION_RE = re.compile(r"V040-[A-Z]-[0-9]{3}")


def _git_blob(path: str) -> bytes:
    return subprocess.check_output(["git", "-C", str(ROOT), "show", f"HEAD:{path}"])


def _section(blob: bytes, heading: str) -> list[str]:
    lines = blob.decode("utf-8").splitlines()
    if lines.count(heading) != 1:
        raise AssertionError(f"heading is not unique: {heading}")
    start = lines.index(heading) + 1
    end = next((index for index in range(start, len(lines)) if lines[index].startswith("## ")), None)
    if end is None:
        raise AssertionError(f"section has no next heading: {heading}")
    return lines[start:end]


def _table(rows: list[str]) -> list[list[str]]:
    parsed = []
    for row in rows:
        cells = [cell.strip() for cell in row.strip().split("|")]
        if len(cells) >= 4 and DECISION_RE.fullmatch(cells[1]):
            parsed.append(cells[1:])
    return parsed


def _owner(decision_id: str, pr: str) -> str:
    letter, number = decision_id.split("-")[1:]
    if "P3" not in pr:
        if "P6" in pr:
            return "p6"
        if "P7" in pr:
            return "p7"
        raise AssertionError(f"unmapped non-P3 decision: {decision_id} {pr}")
    if "A" <= letter <= "F" or "H" <= letter <= "J":
        return "p3b_b"
    if "K" <= letter <= "P" or letter == "R" or (letter == "S" and number == "001") or "T" <= letter <= "Y" or letter == "Z":
        return "p3b_o"
    raise AssertionError(f"unmapped P3 decision: {decision_id} {pr}")


def _expected_from_blobs(blobs: dict[str, bytes]) -> tuple[list[dict], str]:
    v040_rows = _table(_section(blobs[PLAN_PATHS[0]], "## 5. 规范性冻结格式与决策追踪"))
    ids = [row[0] for row in v040_rows]
    if len(ids) != 171 or len(set(ids)) != 171:
        raise AssertionError(f"V0.4.0 decisions raw={len(ids)} unique={len(set(ids))}")
    v041_rows = _table(_section(blobs[PLAN_PATHS[1]], "## 9. V0.4.1 逐项映射"))
    mapping = {row[0]: {"contract": row[1], "pr": row[2]} for row in v041_rows}
    if len(v041_rows) != 150 or len(mapping) != 150:
        raise AssertionError(f"V0.4.1 mapping raw={len(v041_rows)} unique={len(mapping)}")
    if set(ids) - FUTURE_42 - FUTURE_43 != set(mapping):
        raise AssertionError("150/20/1 mapping partition drift")
    manifest = [{"path": path, "sha256": "sha256:" + hashlib.sha256(blobs[path]).hexdigest()} for path in sorted(PLAN_PATHS)]
    entries = []
    for decision_id in ids:
        common = {"decision_id": decision_id, "source_locator": "plan:" + decision_id.lower(), "blocked_projection": False}
        if decision_id in FUTURE_42 or decision_id in FUTURE_43:
            version = "V0.4.2" if decision_id in FUTURE_42 else "V0.4.3"
            common.update({"primary_version": version, "implementation_owner": "future_primary", "requirement_kind": "protected_boundary", "projected_counts": {"rule_fragment": 0, "binding": 0, "key": 0, "candidate": 0}, "scope_topology": "unmodeled", "conflict_approval_assumption": "unmodeled", "derivation_formula": f"{version} protected future-primary decision; no V0.4.1 base appearance workload is projected.", "confidence": "high", "zero_load_reason": "protected_boundary"})
        else:
            contract, pr = mapping[decision_id]["contract"], mapping[decision_id]["pr"]
            if "P3" in pr:
                common.update({"primary_version": "V0.4.1", "implementation_owner": _owner(decision_id, pr), "requirement_kind": "base_appearance" if "MB-" in contract else "system_projection", "projected_counts": {"rule_fragment": 1, "binding": 2, "key": 1, "candidate": 2}, "scope_topology": "single_core_probe", "conflict_approval_assumption": "none", "derivation_formula": "V0.4.1 §9 PR includes P3; one decision-level fragment, two probe bindings/candidates, one normalized key.", "confidence": "medium", "zero_load_reason": "none"})
            else:
                common.update({"primary_version": "V0.4.1", "implementation_owner": _owner(decision_id, pr), "requirement_kind": "system_projection", "projected_counts": {"rule_fragment": 0, "binding": 0, "key": 0, "candidate": 0}, "scope_topology": "unmodeled", "conflict_approval_assumption": "unmodeled", "derivation_formula": "V0.4.1 §9 PR only assigns P6/P7 state/finalization responsibility; no P3 rule assets.", "confidence": "high", "zero_load_reason": "no_base_appearance"})
        entries.append(common)
    return entries, canonical_sha256(manifest)


def _expected_evidence_parameters(config: dict) -> dict[str, dict]:
    generation_seed = config["generation"]["generation_seed"]
    expected = {}
    for kind, cells in (("coverage", config["matrices"]["coverage_cells"]), ("performance", config["matrices"]["performance_cells"])):
        for cell in cells:
            parameters = {"measurement_kind": kind, "scale_id": cell["scale_id"], "scenario_id": cell["scenario_id"], "generation_seed": generation_seed, "permutation_seed": None}
            expected[logical_key(parameters)] = parameters
    for cell in config["matrices"]["determinism_cells"]:
        parameters = {"measurement_kind": "determinism", "scale_id": cell["scale_id"], "scenario_id": cell["scenario_id"], "generation_seed": generation_seed, "permutation_seed": cell["permutation_seed"]}
        expected[logical_key(parameters)] = parameters
    if len(expected) != 66:
        raise AssertionError("formal matrix did not construct 66 logical keys")
    return expected


def _load_complete_evidence_inventory(directory: Path, config: dict) -> dict[str, dict]:
    expected = _expected_evidence_parameters(config)
    if not directory.is_dir():
        raise AssertionError("evidence path is not a directory")
    children = sorted(directory.iterdir(), key=lambda child: child.name)
    if len(children) != 66 or any(not child.is_file() or child.suffix != ".json" or child.name.endswith(".tmp") for child in children):
        raise AssertionError("evidence inventory is not exactly 66 final JSON files")
    documents, observed = {}, set()
    for child in children:
        try:
            document = json.loads(child.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise AssertionError("evidence JSON is unreadable") from exc
        if not isinstance(document, dict):
            raise AssertionError("evidence JSON is not an object")
        try:
            key = logical_key(document["parameters"])
        except (KeyError, TypeError):
            raise AssertionError("evidence JSON has no logical key") from None
        if key in observed:
            raise AssertionError("duplicate evidence logical key")
        observed.add(key)
        if child.stem != key:
            raise AssertionError("evidence filename does not equal logical key")
        documents[key] = document
    if observed != set(expected):
        raise AssertionError("evidence logical key set differs from formal matrix")
    return documents


def _git_blob_at(commit: str, path: str) -> bytes:
    return subprocess.check_output(["git", "-C", str(ROOT), "show", f"{commit}:{path}"])


def _parent_formal_inputs(parent_h: str, blob_reader=_git_blob_at) -> tuple[dict, dict]:
    paths = (
        "format-monograph/references/benchmarks/v0412/p3a-c2/benchmark-config.json",
        "format-monograph/references/benchmarks/v0412/p3a-c2/projected-envelope.json",
    )
    try:
        config, envelope = (json.loads(blob_reader(parent_h, path).decode("utf-8")) for path in paths)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise AssertionError("parent H formal input is unreadable") from exc
    if not isinstance(config, dict) or not isinstance(envelope, dict):
        raise AssertionError("parent H formal input is not an object")
    return config, envelope


def _assert_worktree_formal_bytes_match_parent(parent_h: str, blob_reader=_git_blob_at) -> None:
    for path in ("benchmark-config.json", "projected-envelope.json"):
        parent = blob_reader(parent_h, f"format-monograph/references/benchmarks/v0412/p3a-c2/{path}")
        if (BASE / path).read_bytes() != parent:
            raise AssertionError("E changed a parent-H formal input")


def _expected_evidence_paths(config: dict) -> set[str]:
    prefix = "format-monograph/references/benchmarks/v0412/p3a-c2/"
    paths = {prefix + "reference-results/" + key + ".json" for key in _expected_evidence_parameters(config)}
    return paths | {prefix + name for name in ("suite-summary.json", "reference-environment.json", "SUMMARY.md")}


def _assert_evidence_only_additions(raw: bytes, expected_paths: set[str]) -> None:
    tokens = [token.decode("utf-8") for token in raw.split(b"\0") if token]
    records, index = [], 0
    while index < len(tokens):
        status = tokens[index]; index += 1
        if status.startswith(("R", "C")):
            if index + 1 >= len(tokens):
                raise AssertionError("truncated rename/copy diff record")
            records.append((status, tokens[index], tokens[index + 1])); index += 2
        else:
            if index >= len(tokens):
                raise AssertionError("truncated diff record")
            records.append((status, tokens[index])); index += 1
    if any(record[0] != "A" or len(record) != 2 for record in records):
        raise AssertionError("E diff contains non-addition status")
    actual = {record[1] for record in records}
    if len(actual) != len(records) or actual != expected_paths:
        raise AssertionError("E diff is not the exact evidence whitelist")


def _assert_frozen_result_environment(results: list[dict]) -> None:
    if not results:
        raise AssertionError("no result environments")
    environments = [result.get("environment") for result in results]
    if any(not isinstance(environment, dict) for environment in environments) or any(environment != environments[0] for environment in environments[1:]):
        raise AssertionError("result environments differ")
    expected = {"os_family": "windows", "os_build_class": "frozen_reference", "python_version": "3.12.13", "cpu_architecture": "x86_64", "logical_cpu_count": 24}
    if any(environments[0].get(key) != value for key, value in expected.items()):
        raise AssertionError("result environment does not meet frozen gates")


class FormalInputTests(unittest.TestCase):
    def setUp(self) -> None:
        self.blobs = {path: _git_blob(path) for path in PLAN_PATHS}
        self.expected_entries, self.expected_source_digest = _expected_from_blobs(self.blobs)
        self.envelope = json.loads((BASE / "projected-envelope.json").read_text(encoding="utf-8"))
        self.config = json.loads((BASE / "benchmark-config.json").read_text(encoding="utf-8"))

    def _assert_formal_oracle(self, envelope: dict) -> None:
        self.assertEqual(envelope["projection_kind"], "formal_planning_projection")
        self.assertNotIn("fixture_kind", envelope)
        self.assertNotIn("synthetic_recipe", envelope)
        self.assertEqual(envelope["source_plan_digest"], self.expected_source_digest)
        self.assertEqual({entry["decision_id"]: entry for entry in envelope["entries"]}, {entry["decision_id"]: entry for entry in self.expected_entries})

    def test_source_oracle_closes_all_171_entries_and_preserves_decision_level_proxy_counts(self) -> None:
        self._assert_formal_oracle(self.envelope)
        entries = self.envelope["entries"]
        # These are decision-level probe counts, not a production asset estimate.
        self.assertEqual(self.envelope["decision_population_summary"], {"v041_primary": 150, "v042_protected": 20, "v043_protected": 1})
        self.assertEqual({owner: sum(entry["implementation_owner"] == owner for entry in entries) for owner in ("p3b_b", "p3b_o", "p6", "future_primary")}, {"p3b_b": 54, "p3b_o": 94, "p6": 2, "future_primary": 21})
        self.assertEqual((sum(entry["zero_load_reason"] != "none" for entry in entries), sum(entry["zero_load_reason"] == "none" for entry in entries)), (23, 148))
        self.assertEqual(aggregate_projected_envelope_counts(self.envelope), {"rule_fragment": 148, "binding": 296, "key": 148, "candidate": 296})

    def test_section_10_independently_closes_the_frozen_future_20_plus_1(self) -> None:
        rows = _table(_section(self.blobs[PLAN_PATHS[1]], "## 10. V0.4.2/V0.4.3 主责决策覆盖清单"))
        ids = [row[0] for row in rows]
        self.assertEqual((len(ids), len(set(ids))), (21, 21))
        v042, v043 = set(), set()
        for row in rows:
            if row[1].startswith("V0.4.2"):
                v042.add(row[0])
            elif row[1] == "V0.4.3":
                v043.add(row[0])
            else:
                self.fail(f"unexpected future ownership cell: {row[1]}")
        self.assertEqual(v042, FUTURE_42)
        self.assertEqual(v043, FUTURE_43)

    def test_formal_assets_pass_schema_semantics_and_context(self) -> None:
        for name, document in (("projected-envelope", self.envelope), ("benchmark-config", self.config)):
            schema = json.loads((BASE / f"{name}.schema.json").read_text(encoding="utf-8"))
            Draft202012Validator.check_schema(schema)
            self.assertEqual(list(Draft202012Validator(schema).iter_errors(document)), [], name)
        validate_projected_envelope_semantics(self.envelope)
        validate_benchmark_config_semantics(self.config)
        validate_benchmark_config_against_envelope(self.config, self.envelope)
        self.assertEqual(self.envelope["envelope_digest"], recompute_envelope_digest(self.envelope))
        self.assertEqual(self.config["config_digest"], recompute_config_digest(self.config))
        matrices = self.config["matrices"]
        self.assertEqual((len(matrices["coverage_cells"]), len(matrices["performance_cells"]), len(matrices["determinism_cells"])), (16, 10, 40))
        result = json.loads((FIXTURES / "benchmark-result.valid.json").read_text(encoding="utf-8"))
        result["benchmark_config_digest"] = self.config["config_digest"]
        result["result_digest"] = recompute_result_digest(result)
        context = validate_benchmark_result_context(result, self.config, self.envelope)
        self.assertEqual(context["projection_kind"], "formal_planning_projection")
        self.assertFalse(context["production_representative"])
        self.assertTrue(context["revalidation_required"])
        self.assertEqual(context["aggregate_counts"], {"rule_fragment": 148, "binding": 296, "key": 148, "candidate": 296})
        self.assertEqual(self.envelope["derivation_policy"], "manual_prediction_with_later_mechanical_validation")
        self.assertEqual(self.config["projection_binding"]["representativeness_scope"], "non_production_single_core_probe")

    def test_owner_source_and_entry_mutations_fail_independent_oracle(self) -> None:
        owner = copy.deepcopy(self.envelope)
        owner["entries"][0]["implementation_owner"] = "p3b_o"
        with self.assertRaises(AssertionError): self._assert_formal_oracle(owner)
        source = copy.deepcopy(self.envelope)
        source["source_plan_digest"] = "sha256:" + "0" * 64
        with self.assertRaises(AssertionError): self._assert_formal_oracle(source)
        for field, value in (("derivation_formula", "tampered"), ("source_locator", "plan:tampered"), ("projected_counts", {"rule_fragment": 9, "binding": 9, "key": 9, "candidate": 9})):
            altered = copy.deepcopy(self.envelope)
            altered["entries"][0][field] = value
            with self.subTest(field=field), self.assertRaises(AssertionError): self._assert_formal_oracle(altered)

    def test_plan_blob_and_projection_kind_mutations_fail_formal_oracle(self) -> None:
        altered_blobs = dict(self.blobs)
        altered_blobs[PLAN_PATHS[2]] += b"\n"
        _, altered_digest = _expected_from_blobs(altered_blobs)
        self.assertNotEqual(altered_digest, self.envelope["source_plan_digest"])
        altered = copy.deepcopy(self.envelope)
        altered["source_plan_digest"] = altered_digest
        with self.assertRaises(AssertionError): self._assert_formal_oracle(altered)
        synthetic = copy.deepcopy(self.envelope)
        synthetic["projection_kind"] = "synthetic_contract_fixture"
        synthetic["envelope_digest"] = recompute_envelope_digest(synthetic)
        with self.assertRaises(AssertionError): self._assert_formal_oracle(synthetic)

    def test_config_misbinding_and_blocked_projection_fail_validators(self) -> None:
        config = copy.deepcopy(self.config)
        config["projection_binding"]["projected_envelope_digest"] = "sha256:" + "1" * 64
        config["config_digest"] = recompute_config_digest(config)
        with self.assertRaises(ValueError): validate_benchmark_config_against_envelope(config, self.envelope)
        blocked = copy.deepcopy(self.envelope)
        blocked["entries"][0]["blocked_projection"] = True
        blocked["envelope_digest"] = recompute_envelope_digest(blocked)
        with self.assertRaises(ValueError): validate_projected_envelope_semantics(blocked)

    def test_schema_and_semantic_legal_blocked_variant_still_fails_formal_oracle(self) -> None:
        blocked = copy.deepcopy(self.envelope)
        entry = next(item for item in blocked["entries"] if item["primary_version"] == "V0.4.1" and item["projected_counts"]["rule_fragment"] > 0)
        entry.update({
            "blocked_projection": True,
            "projected_counts": {"rule_fragment": 0, "binding": 0, "key": 0, "candidate": 0},
            "scope_topology": "unmodeled",
            "conflict_approval_assumption": "unmodeled",
            "zero_load_reason": "none",
        })
        blocked["envelope_digest"] = recompute_envelope_digest(blocked)
        schema = json.loads((BASE / "projected-envelope.schema.json").read_text(encoding="utf-8"))
        self.assertEqual(list(Draft202012Validator(schema).iter_errors(blocked)), [])
        validate_projected_envelope_semantics(blocked)
        with self.assertRaises(AssertionError):
            self._assert_formal_oracle(blocked)


class _RegistryKey:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False


class _FakeWinreg:
    KEY_READ = 1
    KEY_WOW64_64KEY = 2
    HKEY_LOCAL_MACHINE = object()

    def __init__(self, values):
        self.values = values

    def OpenKey(self, *_args):
        return _RegistryKey()

    def QueryValueEx(self, _key, name):
        if name not in self.values:
            raise FileNotFoundError(name)
        return self.values[name], 1


class ReferenceDriverTests(unittest.TestCase):
    SUBJECT = "a" * 40

    def _manifest(self):
        return [{"path": path, "sha256": "sha256:" + (f"{index:064x}")} for index, path in enumerate(sorted(SUBJECT_PATHS), 1)]

    def _environment(self, values=None, *, system="Windows", machine="AMD64", version=(3, 12, 13), cpus=24):
        values = values or {"DisplayVersion": "25H2", "UBR": 9168, "CurrentBuild": "26200", "CurrentBuildNumber": "26200"}
        with mock.patch.dict(sys.modules, {"winreg": _FakeWinreg(values)}), \
             mock.patch.object(reference.platform, "system", return_value=system), \
             mock.patch.object(reference.platform, "machine", return_value=machine), \
             mock.patch.object(reference.sys, "version_info", version), \
             mock.patch.object(reference.os, "cpu_count", return_value=cpus):
            reference._reference_environment()

    def test_driver_imports_are_bounded_and_winreg_is_lazy(self):
        self.assertIsNotNone(reference)
        tree = ast.parse(Path(reference.__file__).read_text(encoding="utf-8"))
        imports = {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom) and node.module}
        direct_imports = {name.name for node in tree.body if isinstance(node, ast.Import) for name in node.names}
        self.assertEqual(imports - {"__future__", "pathlib", "typing", "profile_v2_benchmark", "profile_v2_benchmark_runner"}, set())
        self.assertEqual(direct_imports - {"argparse", "json", "os", "platform", "subprocess", "sys"}, set())
        self.assertFalse(any(name in {"profile_v2_composer", "profile_v2_registry", "profile_v2_artifacts", "profile_v2_scope", "profile_v2_values", "profile_v2_authority", "run_monograph"} for name in imports))
        runner_import = next(node for node in ast.walk(tree) if isinstance(node, ast.ImportFrom) and node.module == "profile_v2_benchmark_runner")
        self.assertTrue(all(not item.name.startswith("_") for item in runner_import.names))
        top_level_winreg = [node for node in tree.body if isinstance(node, (ast.Import, ast.ImportFrom)) and "winreg" in ast.dump(node)]
        self.assertEqual(top_level_winreg, [])
        self.assertIn("import winreg", inspect.getsource(reference._reference_environment))
        self.assertEqual(reference.REFERENCE_TIMEOUT_SECONDS, 180.0)
        self.assertEqual(reference.RESULTS_DIRECTORY.name, "reference-results")

    def test_reference_environment_accepts_only_the_frozen_machine_contract(self):
        self._environment()
        cases = (
            ("non_windows", {"system": "Linux"}), ("arch", {"machine": "arm64"}),
            ("python", {"version": (3, 12, 12)}), ("cpu", {"cpus": 23}),
            ("display", {"values": {"DisplayVersion": "25H1", "UBR": 9168, "CurrentBuild": "26200"}}),
            ("ubr", {"values": {"DisplayVersion": "25H2", "UBR": 1, "CurrentBuild": "26200"}}),
            ("no_build", {"values": {"DisplayVersion": "25H2", "UBR": 9168}}),
            ("bad_second_build", {"values": {"DisplayVersion": "25H2", "UBR": 9168, "CurrentBuild": "26200", "CurrentBuildNumber": "1"}}),
        )
        for name, kwargs in cases:
            with self.subTest(name=name), self.assertRaises(reference.ReferencePreflightError):
                self._environment(**kwargs)
        self._environment({"DisplayVersion": "25H2", "UBR": 9168, "CurrentBuild": "26200"})

    def test_git_worktree_gates_fail_closed_but_allow_cache_entries(self):
        cache = reference.RESULTS_DIRECTORY.resolve(strict=False)

        def responder(*args, text=True):
            data = b"" if not text else ""
            if args[:2] == ("rev-parse", "HEAD"):
                data = (self.SUBJECT + "\n") if text else (self.SUBJECT + "\n").encode()
            return subprocess.CompletedProcess(args, 0, stdout=data, stderr=data)

        with mock.patch.object(reference, "_git", side_effect=responder):
            reference._require_clean_subject_worktree(self.SUBJECT, cache)
        for name, mutate in (("head", lambda args: args[:2] == ("rev-parse", "HEAD")), ("unstaged", lambda args: args[:1] == ("diff",) and "--cached" not in args), ("staged", lambda args: args[:1] == ("diff",) and "--cached" in args)):
            def bad(*args, text=True, _mutate=mutate):
                completed = responder(*args, text=text)
                if _mutate(args):
                    return subprocess.CompletedProcess(args, 1, stdout=completed.stdout, stderr=completed.stderr)
                return completed
            with self.subTest(name=name), mock.patch.object(reference, "_git", side_effect=bad), self.assertRaises(reference.ReferencePreflightError):
                reference._require_clean_subject_worktree(self.SUBJECT, cache)
        for name, entry, ignored in (("untracked", b"outside.tmp\0", False), ("ignored", b"outside.tmp\0", True)):
            with self.subTest(name=name), mock.patch.object(reference, "_relative_worktree_entries", side_effect=lambda ignored=False, _entry=entry, _want=ignored: {Path(_entry[:-1].decode())} if ignored == _want else set()), mock.patch.object(reference, "_git", side_effect=responder), self.assertRaises(reference.ReferencePreflightError):
                reference._require_clean_subject_worktree(self.SUBJECT, cache)
        inside = Path("format-monograph/references/benchmarks/v0412/p3a-c2/reference-results/cached.json")
        with mock.patch.object(reference, "_relative_worktree_entries", return_value={inside}), mock.patch.object(reference, "_git", side_effect=responder):
            reference._require_clean_subject_worktree(self.SUBJECT, cache)

    def test_subject_manifest_and_cache_preflight_contracts(self):
        manifest = self._manifest()
        digest = recompute_subject_digest(manifest)
        with mock.patch.object(reference, "build_subject_manifest", return_value=manifest):
            self.assertEqual(reference._require_subject_manifest(self.SUBJECT), (manifest, digest))
        for bad_manifest in (manifest[:-1], manifest[:-1] + [manifest[0]]):
            with self.subTest(manifest=bad_manifest), mock.patch.object(reference, "build_subject_manifest", return_value=bad_manifest), self.assertRaises(reference.ReferencePreflightError):
                reference._require_subject_manifest(self.SUBJECT)
        with mock.patch.object(reference, "build_subject_manifest", side_effect=BenchmarkContractError("bad")), self.assertRaises(reference.ReferencePreflightError):
            reference._require_subject_manifest(self.SUBJECT)
        result = {"parameters": {"measurement_kind": "coverage", "scale_id": "0.5x", "scenario_id": "disjoint", "permutation_seed": None}, "benchmark_subject_commit": self.SUBJECT, "subject_manifest": manifest, "benchmark_subject_digest": digest, "subject_digest_status": {"state": "current", "observed_subject_digest": digest, "revalidation_required": False}}
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary) / "cache"
            reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
            directory.mkdir()
            reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
            for name, create in (("non_json", lambda: (directory / "note.txt").write_text("x")), ("tmp", lambda: (directory / "orphan.tmp").write_text("x")), ("subdirectory", lambda: (directory / "nested").mkdir())):
                create()
                with self.subTest(name=name), self.assertRaises(reference.ReferencePreflightError): reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
                for child in directory.iterdir():
                    if child.is_dir(): child.rmdir()
                    else: child.unlink()
            (directory / "bad.json").write_text("{")
            with mock.patch.object(reference, "scan_cache", side_effect=json.JSONDecodeError("bad", "{", 1)), self.assertRaises(reference.ReferencePreflightError): reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
            (directory / "bad.json").unlink(); (directory / "valid.json").write_text("{}")
            with mock.patch.object(reference, "scan_cache", return_value={"one": result}): reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
            for field, value in (("benchmark_subject_commit", "b" * 40), ("subject_manifest", []), ("benchmark_subject_digest", "sha256:" + "0" * 64), ("subject_digest_status", {})):
                bad = copy.deepcopy(result); bad[field] = value
                with self.subTest(field=field), mock.patch.object(reference, "scan_cache", return_value={"one": bad}), self.assertRaises(reference.ReferencePreflightError): reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)
            for error in (BenchmarkContractError("bad"), BenchmarkRunnerError("bad"), OSError("bad")):
                with self.subTest(error=type(error).__name__), mock.patch.object(reference, "scan_cache", side_effect=error), self.assertRaises(reference.ReferencePreflightError): reference._preflight_cache(directory, {}, {}, self.SUBJECT, manifest, digest)

    def _run_with_stage_failure(self, stage: str, error: Exception):
        config, envelope, cache = {"config": True}, {"envelope": True}, Path(tempfile.gettempdir()) / "c2c-cache"
        manifest, digest = self._manifest(), "sha256:" + "f" * 64
        with mock.patch.object(reference, "_fixed_results_directory", return_value=cache), \
             mock.patch.object(reference, "_load_json", side_effect=[config, envelope]), \
             mock.patch.object(reference, "validate_campaign_inputs") as validate, \
             mock.patch.object(reference, "_require_clean_subject_worktree") as clean, \
             mock.patch.object(reference, "_require_subject_manifest", return_value=(manifest, digest)) as subject, \
             mock.patch.object(reference, "_reference_environment") as environment, \
             mock.patch.object(reference, "_preflight_cache") as cache_preflight, \
             mock.patch.object(reference, "run_benchmark_campaign") as runner:
            {"validate": validate, "clean": clean, "manifest": subject, "environment": environment, "cache": cache_preflight}[stage].side_effect = error
            with self.assertRaises(reference.ReferencePreflightError): reference.run_reference_campaign(self.SUBJECT)
            self.assertEqual(runner.call_count, 0)

    def test_orchestration_calls_runner_once_only_after_every_preflight(self):
        config, envelope, cache, closure = {"config": True}, {"envelope": True}, Path(tempfile.gettempdir()) / "c2c-cache", {"overall_gate": "stop"}
        manifest, digest = self._manifest(), "sha256:" + "f" * 64
        with mock.patch.object(reference, "_fixed_results_directory", return_value=cache), \
             mock.patch.object(reference, "_load_json", side_effect=[config, envelope]) as load, \
             mock.patch.object(reference, "validate_campaign_inputs"), \
             mock.patch.object(reference, "_require_clean_subject_worktree"), \
             mock.patch.object(reference, "_require_subject_manifest", return_value=(manifest, digest)), \
             mock.patch.object(reference, "_reference_environment"), \
             mock.patch.object(reference, "_preflight_cache"), \
             mock.patch.object(reference, "run_benchmark_campaign", return_value=closure) as runner:
            self.assertIs(reference.run_reference_campaign(self.SUBJECT), closure)
            self.assertEqual(load.call_args_list, [mock.call(reference.CONFIG_PATH), mock.call(reference.ENVELOPE_PATH)])
            runner.assert_called_once_with(config, envelope, benchmark_subject_commit=self.SUBJECT, cache_directory=cache, timeout_seconds=180.0, os_build_class="frozen_reference")
        for stage, error in (("validate", BenchmarkContractError("bad")), ("clean", reference.ReferencePreflightError("bad")), ("manifest", reference.ReferencePreflightError("bad")), ("environment", reference.ReferencePreflightError("bad")), ("cache", reference.ReferencePreflightError("bad"))):
            with self.subTest(stage=stage): self._run_with_stage_failure(stage, error)
        with mock.patch.object(reference, "run_benchmark_campaign") as runner, self.assertRaises(reference.ReferencePreflightError):
            reference.run_reference_campaign("invalid")
        self.assertEqual(runner.call_count, 0)

    def test_evidence_inventory_rejects_noncomplete_or_noncanonical_directories_before_suite_validation(self):
        config = json.loads((BASE / "benchmark-config.json").read_text(encoding="utf-8"))
        expected = _expected_evidence_parameters(config)

        def write_entries(directory, keys):
            for key in keys:
                (directory / f"{key}.json").write_text(json.dumps({"parameters": expected[key]}), encoding="utf-8")

        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary) / "evidence"; directory.mkdir()
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            write_entries(directory, list(expected)[:-1])
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            for child in directory.iterdir(): child.unlink()
            write_entries(directory, expected); (directory / "extra.json").write_text("{}")
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            (directory / "extra.json").unlink(); (directory / "stale.tmp").write_text("x")
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            (directory / "stale.tmp").unlink(); (directory / "nested").mkdir()
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            (directory / "nested").rmdir()
            first = next(iter(sorted(expected)))
            (directory / f"{first}.json").write_text("{", encoding="utf-8")
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)
            (directory / f"{first}.json").write_text(json.dumps({"parameters": expected[first]}), encoding="utf-8")
            second = list(sorted(expected))[1]
            (directory / f"{second}.json").write_text(json.dumps({"parameters": expected[first]}), encoding="utf-8")
            with self.assertRaisesRegex(AssertionError, "duplicate"): _load_complete_evidence_inventory(directory, config)
            (directory / f"{second}.json").write_text(json.dumps({"parameters": expected[second]}), encoding="utf-8")
            (directory / f"{first}.json").rename(directory / "wrong-name.json")
            with self.assertRaises(AssertionError): _load_complete_evidence_inventory(directory, config)

    def test_actual_reference_evidence_is_all_or_nothing(self):
        directory = reference.RESULTS_DIRECTORY
        if not directory.exists():
            self.skipTest("reference-results is absent during H")
        commits = subprocess.check_output(["git", "-C", str(ROOT), "rev-list", "--parents", "-n", "1", "HEAD"], text=True).split()
        self.assertEqual(len(commits), 2)
        parent_h = commits[1]
        config, envelope = _parent_formal_inputs(parent_h)
        _assert_worktree_formal_bytes_match_parent(parent_h)
        diff = subprocess.check_output(["git", "-C", str(ROOT), "diff-tree", "-r", "--no-commit-id", "--name-status", "-z", parent_h, "HEAD"])
        _assert_evidence_only_additions(diff, _expected_evidence_paths(config))
        documents = _load_complete_evidence_inventory(directory, config)
        subjects = {result["benchmark_subject_commit"] for result in documents.values()}
        self.assertEqual(subjects, {parent_h})
        subject = subjects.pop()
        _assert_frozen_result_environment(list(documents.values()))
        manifest = reference.build_subject_manifest(subject, repository=ROOT)
        digest = recompute_subject_digest(manifest)
        reference._preflight_cache(directory, config, envelope, subject, manifest, digest)
        cached = scan_cache(directory, config, envelope, subject_digest=digest)
        self.assertEqual(set(cached), set(_expected_evidence_parameters(config)))
        for result in cached.values(): validate_benchmark_result_context(result, config, envelope)
        closure = validate_complete_benchmark_suite(list(cached.values()), config, envelope)
        self.assertTrue(closure["structurally_complete"])
        self.assertEqual(closure["overall_gate"], "go")
        summary_path = directory.parent / "suite-summary.json"
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        self.assertIsInstance(summary, dict)
        self.assertEqual(summary, closure)
        self.assertEqual(canonical_json_bytes(summary), canonical_json_bytes(closure))

    def test_parent_input_evidence_diff_and_environment_helpers_fail_closed(self):
        config = json.loads((BASE / "benchmark-config.json").read_text(encoding="utf-8"))
        expected_paths = _expected_evidence_paths(config)
        raw = b"".join(b"A\0" + path.encode("utf-8") + b"\0" for path in sorted(expected_paths))
        _assert_evidence_only_additions(raw, expected_paths)
        for label, altered in (("extra", raw + b"A\0extra.txt\0"), ("missing", b"".join(b"A\0" + path.encode("utf-8") + b"\0" for path in sorted(expected_paths)[1:])), ("modified", b"M\0" + sorted(expected_paths)[0].encode("utf-8") + b"\0"), ("renamed", b"R100\0old\0new\0")):
            with self.subTest(diff=label), self.assertRaises(AssertionError): _assert_evidence_only_additions(altered, expected_paths)
        config_path = "format-monograph/references/benchmarks/v0412/p3a-c2/benchmark-config.json"
        envelope_path = "format-monograph/references/benchmarks/v0412/p3a-c2/projected-envelope.json"
        parent_bytes = {config_path: (BASE / "benchmark-config.json").read_bytes(), envelope_path: (BASE / "projected-envelope.json").read_bytes()}
        _parent_formal_inputs("h", lambda _commit, path: parent_bytes[path])
        _assert_worktree_formal_bytes_match_parent("h", lambda _commit, path: parent_bytes[path])
        bad_bytes = dict(parent_bytes); bad_bytes[config_path] = b"{}"
        with self.assertRaises(AssertionError): _assert_worktree_formal_bytes_match_parent("h", lambda _commit, path: bad_bytes[path])
        environment = {"os_family": "windows", "os_build_class": "frozen_reference", "python_version": "3.12.13", "cpu_architecture": "x86_64", "logical_cpu_count": 24, "ram_tier": "16_to_32gib"}
        _assert_frozen_result_environment([{"environment": environment}, {"environment": dict(environment)}])
        for field, value in (("os_family", "linux"), ("os_build_class", "public_ci"), ("python_version", "3.12.12"), ("cpu_architecture", "arm64"), ("logical_cpu_count", 23)):
            bad = dict(environment); bad[field] = value
            with self.subTest(field=field), self.assertRaises(AssertionError): _assert_frozen_result_environment([{"environment": bad}])
        ram = dict(environment); ram["ram_tier"] = "32_to_64gib"
        with self.assertRaises(AssertionError): _assert_frozen_result_environment([{"environment": environment}, {"environment": ram}])

    def test_main_fail_closed_hides_contract_and_runner_error_details(self):
        for error in (reference.ReferencePreflightError("C:\\private\\detail"), BenchmarkContractError("contract detail"), BenchmarkRunnerError("runner detail")):
            with self.subTest(error=type(error).__name__), mock.patch.object(reference, "run_reference_campaign", side_effect=error), mock.patch.object(reference, "run_benchmark_campaign") as runner, mock.patch.object(reference.sys, "stderr", new_callable=io.StringIO) as stderr:
                self.assertEqual(reference.main(["--benchmark-subject-commit", self.SUBJECT]), 2)
                self.assertEqual(stderr.getvalue(), "reference benchmark preflight failed\n")
                runner.assert_not_called()
