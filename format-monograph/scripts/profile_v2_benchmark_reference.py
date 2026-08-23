#!/usr/bin/env python3
"""Internal C2C reference-driver preflight; never a public runtime command."""
from __future__ import annotations

import argparse
import json
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any

from profile_v2_benchmark import BenchmarkContractError, canonical_json_bytes, recompute_subject_digest
from profile_v2_benchmark_runner import (
    SUBJECT_PATHS,
    BenchmarkRunnerError,
    build_subject_manifest,
    run_benchmark_campaign,
    scan_cache,
    validate_campaign_inputs,
)


class ReferencePreflightError(ValueError):
    """Reference-only preflight failed before the runner may start a child."""


REPOSITORY = Path(__file__).resolve().parents[2]
BENCHMARK_DIRECTORY = REPOSITORY / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2"
CONFIG_PATH = BENCHMARK_DIRECTORY / "benchmark-config.json"
ENVELOPE_PATH = BENCHMARK_DIRECTORY / "projected-envelope.json"
RESULTS_DIRECTORY = BENCHMARK_DIRECTORY / "reference-results"
REFERENCE_TIMEOUT_SECONDS = 180.0


def _require_commit(value: str) -> str:
    if len(value) != 40 or any(character not in "0123456789abcdef" for character in value):
        raise ReferencePreflightError("benchmark subject must be a lowercase SHA-1")
    return value


def _git(*args: str, text: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(REPOSITORY), *args], capture_output=True, text=text, check=False)


def _git_required(*args: str, text: bool = True) -> subprocess.CompletedProcess:
    completed = _git(*args, text=text)
    if completed.returncode:
        raise ReferencePreflightError("Git preflight failed")
    return completed


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReferencePreflightError("reference input is unreadable") from exc
    if not isinstance(value, dict):
        raise ReferencePreflightError("reference input is not an object")
    return value


def _fixed_results_directory() -> Path:
    repository = REPOSITORY.resolve()
    directory = RESULTS_DIRECTORY.resolve(strict=False)
    try:
        directory.relative_to(repository)
    except ValueError as exc:
        raise ReferencePreflightError("reference cache escaped repository") from exc
    if directory != (repository / "format-monograph" / "references" / "benchmarks" / "v0412" / "p3a-c2" / "reference-results"):
        raise ReferencePreflightError("reference cache is not the fixed directory")
    return directory


def _relative_worktree_entries(*, ignored: bool) -> set[Path]:
    args = ["ls-files", "--others", "--exclude-standard", "-z"]
    if ignored:
        args.insert(2, "--ignored")
    output = _git_required(*args, text=False).stdout
    return {Path(item.decode("utf-8")) for item in output.split(b"\0") if item}


def _require_clean_subject_worktree(subject_commit: str, cache_directory: Path) -> None:
    if _git_required("rev-parse", "HEAD").stdout.strip() != subject_commit:
        raise ReferencePreflightError("HEAD is not the reference subject")
    for arguments in (("diff", "--quiet", "--no-ext-diff", subject_commit, "--"), ("diff", "--cached", "--quiet", "--no-ext-diff", subject_commit, "--")):
        if _git(*arguments).returncode:
            raise ReferencePreflightError("tracked worktree differs from reference subject")
    for relative in _relative_worktree_entries(ignored=False) | _relative_worktree_entries(ignored=True):
        candidate = (REPOSITORY / relative).resolve(strict=False)
        try:
            candidate.relative_to(cache_directory)
        except ValueError as exc:
            raise ReferencePreflightError("untracked or ignored worktree entry outside reference cache") from exc


def _reference_environment() -> None:
    if platform.system() != "Windows":
        raise ReferencePreflightError("reference platform is not Windows")
    machine = platform.machine().lower().replace("-", "_")
    if machine not in {"amd64", "x86_64", "x64"}:
        raise ReferencePreflightError("reference CPU architecture differs")
    if sys.version_info[:3] != (3, 12, 13) or os.cpu_count() != 24:
        raise ReferencePreflightError("reference Python or CPU count differs")
    try:
        import winreg
        access = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion", 0, access) as key:
            display = str(winreg.QueryValueEx(key, "DisplayVersion")[0])
            ubr = str(winreg.QueryValueEx(key, "UBR")[0])
            builds = []
            for name in ("CurrentBuild", "CurrentBuildNumber"):
                try:
                    builds.append(str(winreg.QueryValueEx(key, name)[0]))
                except FileNotFoundError:
                    pass
    except (ImportError, OSError) as exc:
        raise ReferencePreflightError("reference Windows registry is unavailable") from exc
    if display != "25H2" or ubr != "9168" or not builds or any(build != "26200" for build in builds):
        raise ReferencePreflightError("reference Windows build differs")


def _require_subject_manifest(subject_commit: str) -> tuple[list[dict[str, str]], str]:
    try:
        manifest = build_subject_manifest(subject_commit, repository=REPOSITORY)
        digest = recompute_subject_digest(manifest)
    except (BenchmarkContractError, BenchmarkRunnerError) as exc:
        raise ReferencePreflightError("reference subject manifest is invalid") from exc
    if len(manifest) != 25 or tuple(item["path"] for item in manifest) != tuple(sorted(SUBJECT_PATHS)):
        raise ReferencePreflightError("reference subject manifest inventory differs")
    if len({item["path"] for item in manifest}) != len(manifest):
        raise ReferencePreflightError("reference subject manifest duplicates a path")
    return manifest, digest


def _result_key(result: dict[str, Any]) -> tuple[Any, ...]:
    parameters = result.get("parameters")
    if not isinstance(parameters, dict):
        raise ReferencePreflightError("cache result has no logical key")
    return tuple(parameters.get(name) for name in ("measurement_kind", "scale_id", "scenario_id", "permutation_seed"))


def _preflight_cache(directory: Path, config: dict[str, Any], envelope: dict[str, Any], subject_commit: str, manifest: list[dict[str, str]], subject_digest: str) -> None:
    if not directory.exists():
        return
    if not directory.is_dir():
        raise ReferencePreflightError("reference cache is not a directory")
    try:
        children = list(directory.iterdir())
    except OSError as exc:
        raise ReferencePreflightError("reference cache is unreadable") from exc
    if not children:
        return
    if any(not child.is_file() or child.suffix != ".json" or child.name.endswith(".tmp") for child in children):
        raise ReferencePreflightError("reference cache contains a non-final entry")
    try:
        cached = scan_cache(directory, config, envelope, subject_digest=subject_digest)
    except (BenchmarkContractError, BenchmarkRunnerError, OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReferencePreflightError("reference cache is invalid") from exc
    if len(cached) != len(children) or len({_result_key(item) for item in cached.values()}) != len(cached):
        raise ReferencePreflightError("reference cache has unknown or duplicate logical evidence")
    expected_status = {"state": "current", "observed_subject_digest": subject_digest, "revalidation_required": False}
    for result in cached.values():
        if result.get("benchmark_subject_commit") != subject_commit or result.get("subject_manifest") != manifest or result.get("benchmark_subject_digest") != subject_digest or result.get("subject_digest_status") != expected_status:
            raise ReferencePreflightError("reference cache subject identity differs")


def run_reference_campaign(benchmark_subject_commit: str) -> dict[str, Any]:
    """Run the fixed reference campaign only after all H preflight gates pass."""
    subject_commit = _require_commit(benchmark_subject_commit)
    cache_directory = _fixed_results_directory()
    config, envelope = _load_json(CONFIG_PATH), _load_json(ENVELOPE_PATH)
    try:
        validate_campaign_inputs(config, envelope)
    except (BenchmarkContractError, BenchmarkRunnerError) as exc:
        raise ReferencePreflightError("formal inputs are invalid") from exc
    _require_clean_subject_worktree(subject_commit, cache_directory)
    manifest, digest = _require_subject_manifest(subject_commit)
    _reference_environment()
    _preflight_cache(cache_directory, config, envelope, subject_commit, manifest, digest)
    return run_benchmark_campaign(
        config, envelope, benchmark_subject_commit=subject_commit,
        cache_directory=cache_directory, timeout_seconds=REFERENCE_TIMEOUT_SECONDS,
        os_build_class="frozen_reference",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="internal C2C reference benchmark driver")
    parser.add_argument("--benchmark-subject-commit", required=True)
    args = parser.parse_args(argv)
    try:
        closure = run_reference_campaign(args.benchmark_subject_commit)
    except (ReferencePreflightError, BenchmarkContractError, BenchmarkRunnerError):
        print("reference benchmark preflight failed", file=sys.stderr)
        return 2
    sys.stdout.buffer.write(canonical_json_bytes(closure) + b"\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
