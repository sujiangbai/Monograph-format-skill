#!/usr/bin/env python3
"""Portable, resumable orchestration for whole-book formatting runs."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
STATE_NAME = "run-state.json"
STATE_SCHEMA_VERSION = "1.0"
DELIVERY_STATES = {
    "analysis_only",
    "prepared",
    "blocked_qa",
    "candidate_ready",
    "final_ready",
    "failed",
}
RESOLVED_QA_STATES = {"accepted", "resolved", "closed"}
REFRESHED_FIELD_STATES = {
    "absent",
    "refreshed",
    "refreshed_external",
    "refreshed_target_word",
}


class RunError(RuntimeError):
    """A safe, user-facing orchestration failure."""


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_sha256(value: Any) -> str:
    payload = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RunError(f"Invalid JSON file: {path.name}") from exc
    if not isinstance(value, dict):
        raise RunError(f"JSON root must be an object: {path.name}")
    return value


def write_json_atomic(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temp.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    temp.replace(path)


def state_path(work_dir: Path) -> Path:
    return work_dir / STATE_NAME


def load_state(work_dir: Path) -> dict[str, Any]:
    path = state_path(work_dir)
    if not path.is_file():
        raise RunError("Run state was not found. Run prepare first.")
    value = read_json(path)
    if value.get("schema_version") != STATE_SCHEMA_VERSION:
        raise RunError("Unsupported run-state schema version.")
    if value.get("status") not in DELIVERY_STATES:
        raise RunError("Run state contains an invalid delivery status.")
    return value


def save_state(work_dir: Path, state: dict[str, Any]) -> None:
    state["updated_at"] = utc_now()
    write_json_atomic(state_path(work_dir), state)


def artifact_path(work_dir: Path, value: str | None) -> Path | None:
    if not value:
        return None
    path = Path(value)
    return path if path.is_absolute() else work_dir / path


def relative_artifact(work_dir: Path, path: Path) -> str:
    try:
        return str(path.resolve().relative_to(work_dir.resolve()))
    except ValueError:
        return str(path.resolve())


def command_json(script: str, *arguments: object) -> dict[str, Any]:
    completed = run_script(script, *arguments)
    if completed.returncode != 0:
        raise RunError(f"{script} failed with exit code {completed.returncode}.")
    try:
        value = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RunError(f"{script} did not return valid JSON.") from exc
    if not isinstance(value, dict):
        raise RunError(f"{script} returned a non-object JSON value.")
    return value


def run_script(script: str, *arguments: object) -> subprocess.CompletedProcess[str]:
    command = [sys.executable, str(SCRIPT_DIR / script)]
    command.extend(str(argument) for argument in arguments if argument is not None)
    return subprocess.run(
        command,
        cwd=SCRIPT_DIR.parent,
        capture_output=True,
        text=True,
        check=False,
    )


def begin_stage(state: dict[str, Any], name: str, input_key: str) -> float:
    state.setdefault("stages", {})[name] = {
        "status": "running",
        "input_key_sha256": input_key,
        "started_at": utc_now(),
        "cache_hit": False,
    }
    return time.monotonic()


def finish_stage(
    state: dict[str, Any], name: str, started: float, *, status: str = "complete"
) -> None:
    stage = state["stages"][name]
    stage.update(
        {
            "status": status,
            "completed_at": utc_now(),
            "duration_seconds": round(time.monotonic() - started, 3),
        }
    )


def cached_stage(
    state: dict[str, Any], name: str, input_key: str, artifacts: list[Path]
) -> bool:
    stage = state.get("stages", {}).get(name, {})
    return bool(
        stage.get("status") == "complete"
        and stage.get("input_key_sha256") == input_key
        and all(path.is_file() for path in artifacts)
    )


def mark_cache_hit(state: dict[str, Any], name: str) -> None:
    stage = state["stages"][name]
    stage["cache_hit"] = True
    stage["cache_hits"] = int(stage.get("cache_hits", 0)) + 1
    stage["last_reused_at"] = utc_now()


def safe_failure(
    work_dir: Path, state: dict[str, Any] | None, stage: str, message: str
) -> None:
    if state is None:
        return
    state["status"] = "failed"
    state.setdefault("stages", {}).setdefault(stage, {}).update(
        {"status": "failed", "completed_at": utc_now(), "error": message}
    )
    save_state(work_dir, state)


def source_descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.suffix.lower() != ".docx":
        raise RunError("Input must be an existing DOCX file.")
    return {
        "path": str(path.resolve()),
        "name": path.name,
        "sha256": file_sha256(path),
        "size_bytes": path.stat().st_size,
    }


def profile_descriptor(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RunError("Profile file was not found.")
    return {
        "path": str(path.resolve()),
        "name": path.name,
        "sha256": file_sha256(path),
    }


def new_state(source: dict[str, Any], profile: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": STATE_SCHEMA_VERSION,
        "run_id": str(uuid.uuid4()),
        "status": "prepared",
        "created_at": utc_now(),
        "updated_at": utc_now(),
        "source": source,
        "profile": profile,
        "structure_map": None,
        "capabilities": {},
        "stages": {},
        "artifacts": {},
        "blockers": [],
        "qa_groups": [],
        "frozen_scopes": [],
        "metrics": {"cache_hits": 0},
    }


def prepare(args: argparse.Namespace) -> int:
    work_dir = args.work_dir.resolve()
    work_dir.mkdir(parents=True, exist_ok=True)
    source = source_descriptor(args.input.resolve())
    profile = profile_descriptor(args.profile.resolve())
    state_file = state_path(work_dir)
    state = read_json(state_file) if state_file.is_file() else new_state(source, profile)
    if state.get("schema_version") != STATE_SCHEMA_VERSION:
        state = new_state(source, profile)

    same_inputs = (
        state.get("source", {}).get("sha256") == source["sha256"]
        and state.get("profile", {}).get("sha256") == profile["sha256"]
    )
    if not same_inputs:
        state = new_state(source, profile)
    else:
        state["source"] = source
        state["profile"] = profile

    inventory_path = work_dir / "inventory.json"
    map_path = work_dir / "candidate-structure-map.json"
    input_key = json_sha256({"source": source["sha256"], "profile": profile["sha256"]})
    if args.resume and cached_stage(
        state, "prepare", input_key, [inventory_path, map_path]
    ):
        mark_cache_hit(state, "prepare")
        state["metrics"]["cache_hits"] = int(
            state["metrics"].get("cache_hits", 0)
        ) + 1
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    started = begin_stage(state, "prepare", input_key)
    save_state(work_dir, state)
    environment = command_json("check_environment.py", "--json")
    state["capabilities"] = environment
    if not environment.get("capabilities", {}).get("inspection"):
        state["status"] = "analysis_only"
        state["blockers"] = [
            {
                "id": "capability:inspection",
                "kind": "capability",
                "status": "open",
            }
        ]
        finish_stage(state, "prepare", started, status="blocked")
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2

    validation = run_script("validate_profile.py", args.profile.resolve())
    if validation.returncode != 0:
        raise RunError("validate_profile.py failed with exit code 1.")
    inspection = run_script(
        "inspect_docx.py",
        args.input.resolve(),
        "--output",
        inventory_path,
        "--structure-map-output",
        map_path,
    )
    if inspection.returncode != 0:
        raise RunError("inspect_docx.py failed with exit code 1.")

    structure_map = read_json(map_path)
    state["status"] = "prepared"
    state["structure_map"] = {
        "path": relative_artifact(work_dir, map_path),
        "sha256": file_sha256(map_path),
        "schema_version": structure_map.get("schema_version"),
        "status": structure_map.get("status"),
    }
    state["qa_groups"] = list(structure_map.get("qa_groups", []))
    state["frozen_scopes"] = list(structure_map.get("frozen_scopes", []))
    state["blockers"] = []
    state["artifacts"].update(
        {
            "inventory": relative_artifact(work_dir, inventory_path),
            "candidate_structure_map": relative_artifact(work_dir, map_path),
        }
    )
    finish_stage(state, "prepare", started)
    save_state(work_dir, state)
    print(json.dumps(state, ensure_ascii=False, indent=2))
    return 0


def open_qa_items(structure_map: dict[str, Any]) -> tuple[list[dict], list[dict]]:
    qa_groups = [
        item
        for item in structure_map.get("qa_groups", [])
        if item.get("status", "open") not in RESOLVED_QA_STATES
    ]
    frozen = [
        item
        for item in structure_map.get("frozen_scopes", [])
        if item.get("status", "open") not in RESOLVED_QA_STATES
    ]
    return qa_groups, frozen


def current_inputs(state: dict[str, Any]) -> tuple[Path, Path]:
    source = Path(state["source"]["path"])
    profile = Path(state["profile"]["path"])
    if not source.is_file() or file_sha256(source) != state["source"]["sha256"]:
        raise RunError("Source changed after prepare; run prepare again.")
    if not profile.is_file() or file_sha256(profile) != state["profile"]["sha256"]:
        raise RunError("Profile changed after prepare; run prepare again.")
    return source, profile


def apply(args: argparse.Namespace) -> int:
    work_dir = args.work_dir.resolve()
    state = load_state(work_dir)
    source, profile = current_inputs(state)
    structure_path = args.structure_map.resolve()
    structure_map = read_json(structure_path)
    qa_groups, frozen = open_qa_items(structure_map)
    map_descriptor = {
        "path": str(structure_path),
        "sha256": file_sha256(structure_path),
        "schema_version": structure_map.get("schema_version"),
        "status": structure_map.get("status"),
    }
    state["structure_map"] = map_descriptor
    state["qa_groups"] = qa_groups
    state["frozen_scopes"] = frozen
    input_key = json_sha256(
        {
            "source": state["source"]["sha256"],
            "profile": state["profile"]["sha256"],
            "map": map_descriptor["sha256"],
        }
    )
    output_dir = work_dir / "applied"
    formatted = output_dir / f"{source.stem}-formatted.docx"
    review = output_dir / f"{source.stem}-review.docx"
    report = output_dir / f"{source.stem}-format-report.md"
    audit = output_dir / "audit.json"
    if args.resume and cached_stage(
        state, "apply", input_key, [formatted, review, report, audit]
    ):
        mark_cache_hit(state, "apply")
        state["metrics"]["cache_hits"] = int(
            state["metrics"].get("cache_hits", 0)
        ) + 1
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    started = begin_stage(state, "apply", input_key)
    save_state(work_dir, state)
    validation = run_script(
        "validate_structure_map.py", structure_path, "--source", source
    )
    if validation.returncode != 0:
        state["status"] = "blocked_qa"
        state["blockers"] = [
            {
                "id": "structure-map:approval",
                "kind": "structure_map",
                "status": "open",
            }
        ]
        finish_stage(state, "apply", started, status="blocked")
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2

    command: list[object] = [
        source,
        "--profile",
        profile,
        "--structure-map",
        structure_path,
        "--output-dir",
        output_dir,
    ]
    if output_dir.exists():
        command.append("--force")
    if args.allow_missing_fonts:
        command.append("--allow-missing-fonts")
    applied = run_script("apply_profile.py", *command)
    if applied.returncode != 0:
        raise RunError("apply_profile.py failed with exit code 1.")
    audited = run_script(
        "audit_docx.py",
        source,
        formatted,
        "--profile",
        profile,
        "--structure-map",
        structure_path,
        "--output",
        audit,
    )
    if audited.returncode != 0:
        raise RunError("Post-application content audit failed.")

    state["artifacts"].update(
        {
            "formatted": relative_artifact(work_dir, formatted),
            "review": relative_artifact(work_dir, review),
            "format_report": relative_artifact(work_dir, report),
            "apply_audit": relative_artifact(work_dir, audit),
        }
    )
    state["blockers"] = [
        {"id": str(item.get("id", "qa")), "kind": "qa_group", "status": "open"}
        for item in qa_groups
    ] + [
        {
            "id": str(item.get("id", "frozen_scope")),
            "kind": "frozen_scope",
            "status": "open",
        }
        for item in frozen
    ]
    state["status"] = "blocked_qa" if state["blockers"] else "candidate_ready"
    finish_stage(state, "apply", started)
    save_state(work_dir, state)
    print(json.dumps(state, ensure_ascii=False, indent=2))
    return 0


def finalize(args: argparse.Namespace) -> int:
    work_dir = args.work_dir.resolve()
    state = load_state(work_dir)
    source, profile = current_inputs(state)
    if state.get("blockers") or state.get("qa_groups") or state.get("frozen_scopes"):
        state["status"] = "blocked_qa"
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2
    structure = state.get("structure_map") or {}
    structure_path = Path(structure.get("path", ""))
    if not structure_path.is_file() or file_sha256(structure_path) != structure.get(
        "sha256"
    ):
        raise RunError("Approved structure map changed; run apply again.")
    formatted = artifact_path(work_dir, state.get("artifacts", {}).get("formatted"))
    if formatted is None or not formatted.is_file():
        raise RunError("Formatted candidate was not found. Run apply first.")
    final_dir = work_dir / "final"
    final_docx = final_dir / f"{source.stem}-finalized.docx"
    final_status = final_dir / "finalization.json"
    target_pdf = final_dir / f"{source.stem}-target.pdf"
    input_key = json_sha256(
        {
            "formatted": file_sha256(formatted),
            "map": structure["sha256"],
            "field_updater": args.field_updater,
            "field_command": args.field_updater_command or "",
            "target": args.target_software or "",
        }
    )
    if args.resume and cached_stage(
        state, "finalize", input_key, [final_docx, final_status]
    ):
        mark_cache_hit(state, "finalize")
        state["metrics"]["cache_hits"] = int(
            state["metrics"].get("cache_hits", 0)
        ) + 1
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    started = begin_stage(state, "finalize", input_key)
    final_dir.mkdir(parents=True, exist_ok=True)
    command: list[object] = [
        formatted,
        "--source",
        source,
        "--profile",
        profile,
        "--structure-map",
        structure_path,
        "--output",
        final_docx,
        "--status-output",
        final_status,
        "--field-updater",
        args.field_updater,
        "--pdf-output",
        target_pdf,
        "--force",
    ]
    if args.field_updater_command:
        command.extend(["--field-updater-command", args.field_updater_command])
    if args.target_software:
        command.extend(["--target-software", args.target_software])
    if args.renderer:
        command.extend(["--renderer", args.renderer])
    if args.approve_deferred:
        command.append("--approve-deferred")
    completed = run_script("finalize_docx.py", *command)
    if completed.returncode != 0:
        raise RunError("finalize_docx.py failed with exit code 1.")
    evidence = read_json(final_status)
    field_status = str(evidence.get("delivery_field_status", "stale"))
    state["artifacts"].update(
        {
            "finalized": relative_artifact(work_dir, final_docx),
            "finalization_status": relative_artifact(work_dir, final_status),
        }
    )
    if target_pdf.is_file():
        state["artifacts"]["target_pdf"] = relative_artifact(work_dir, target_pdf)
    if field_status not in REFRESHED_FIELD_STATES:
        state["status"] = "candidate_ready"
        state["blockers"] = [
            {"id": "field-update:deferred", "kind": "field_update", "status": "open"}
        ]
    else:
        state["status"] = "candidate_ready"
        state["blockers"] = []
    finish_stage(state, "finalize", started)
    save_state(work_dir, state)
    print(json.dumps(state, ensure_ascii=False, indent=2))
    return 0


def validate_visual_manifest(path: Path, page_count: int) -> dict[str, Any]:
    value = read_json(path)
    required = {
        "all_pages_inspected": True,
        "target_layout_verified": True,
        "page_count": page_count,
    }
    for key, expected in required.items():
        if value.get(key) != expected:
            raise RunError(f"Visual QA manifest has invalid {key}.")
    if value.get("issues") not in ([], None):
        raise RunError("Visual QA manifest contains unresolved issues.")
    return value


def has_target_layout_evidence(render_manifest: dict[str, Any]) -> bool:
    if render_manifest.get("target_pdf_source"):
        return True
    target = str(render_manifest.get("target_software") or "").casefold()
    return bool(target and not render_manifest.get("target_layout_unverified"))


def verify(args: argparse.Namespace) -> int:
    work_dir = args.work_dir.resolve()
    state = load_state(work_dir)
    source, profile = current_inputs(state)
    non_visual_blockers = [
        item
        for item in state.get("blockers", [])
        if item.get("kind") != "visual_qa"
    ]
    if non_visual_blockers:
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2
    structure = state.get("structure_map") or {}
    structure_path = Path(structure.get("path", ""))
    finalized = artifact_path(work_dir, state.get("artifacts", {}).get("finalized"))
    if finalized is None or not finalized.is_file():
        raise RunError("Finalized DOCX was not found. Run finalize first.")
    target_pdf = artifact_path(work_dir, state.get("artifacts", {}).get("target_pdf"))
    render_dir = work_dir / "rendered"
    audit_path = work_dir / "final" / "audit.json"
    visual_hash = file_sha256(args.visual_qa_manifest) if args.visual_qa_manifest else None
    input_key = json_sha256(
        {
            "finalized": file_sha256(finalized),
            "map": structure.get("sha256"),
            "target_pdf": file_sha256(target_pdf) if target_pdf and target_pdf.is_file() else None,
            "visual_manifest": visual_hash,
        }
    )
    required_artifacts = [audit_path, render_dir / "render-manifest.json"]
    if args.resume and cached_stage(state, "verify", input_key, required_artifacts):
        mark_cache_hit(state, "verify")
        state["metrics"]["cache_hits"] = int(
            state["metrics"].get("cache_hits", 0)
        ) + 1
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    started = begin_stage(state, "verify", input_key)
    audited = run_script(
        "audit_docx.py",
        source,
        finalized,
        "--profile",
        profile,
        "--structure-map",
        structure_path,
        "--output",
        audit_path,
    )
    if audited.returncode != 0:
        raise RunError("Final content audit failed.")
    render_command: list[object] = [
        finalized,
        "--output-dir",
        render_dir,
        "--keep-pdf",
        "--force",
    ]
    if target_pdf and target_pdf.is_file():
        render_command.extend(["--target-pdf", target_pdf])
    elif args.renderer:
        render_command.extend(["--renderer", args.renderer])
    if args.target_software:
        render_command.extend(["--target-software", args.target_software])
    rendered = run_script("render_docx.py", *render_command)
    if rendered.returncode != 0:
        raise RunError("render_docx.py failed with exit code 1.")
    render_manifest = read_json(render_dir / "render-manifest.json")
    page_count = int(render_manifest.get("page_count", 0))
    state["artifacts"].update(
        {
            "final_audit": relative_artifact(work_dir, audit_path),
            "render_manifest": relative_artifact(
                work_dir, render_dir / "render-manifest.json"
            ),
        }
    )
    state["metrics"]["rendered_pages"] = page_count
    if args.visual_qa_manifest:
        visual = validate_visual_manifest(args.visual_qa_manifest.resolve(), page_count)
        state["artifacts"]["visual_qa_manifest"] = str(
            args.visual_qa_manifest.resolve()
        )
        state["visual_qa"] = {
            "all_pages_inspected": True,
            "target_layout_verified": True,
            "page_count": page_count,
            "verified_at": visual.get("verified_at") or utc_now(),
        }
        if has_target_layout_evidence(render_manifest):
            state["status"] = "final_ready"
            state["blockers"] = []
        else:
            state["status"] = "candidate_ready"
            state["blockers"] = [
                {
                    "id": "visual-qa:target-layout",
                    "kind": "visual_qa",
                    "status": "open",
                }
            ]
    else:
        state["status"] = "candidate_ready"
        state["blockers"] = [
            {"id": "visual-qa:all-pages", "kind": "visual_qa", "status": "open"}
        ]
    finish_stage(state, "verify", started)
    save_state(work_dir, state)
    print(json.dumps(state, ensure_ascii=False, indent=2))
    return 0


def status(args: argparse.Namespace) -> int:
    state = load_state(args.work_dir.resolve())
    if args.json:
        print(json.dumps(state, ensure_ascii=False, indent=2))
    else:
        print(f"Run status: {state['status']}")
        print(f"Run id: {state['run_id']}")
        print(f"Blockers: {len(state.get('blockers', []))}")
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(
        description="Run format-monograph through portable resumable stages."
    )
    subparsers = result.add_subparsers(dest="command", required=True)

    prepare_parser = subparsers.add_parser("prepare")
    prepare_parser.add_argument("input", type=Path)
    prepare_parser.add_argument("--profile", required=True, type=Path)
    prepare_parser.add_argument("--work-dir", required=True, type=Path)
    prepare_parser.add_argument("--resume", action="store_true")
    prepare_parser.set_defaults(handler=prepare)

    apply_parser = subparsers.add_parser("apply")
    apply_parser.add_argument("--work-dir", required=True, type=Path)
    apply_parser.add_argument("--structure-map", required=True, type=Path)
    apply_parser.add_argument("--resume", action="store_true")
    apply_parser.add_argument("--allow-missing-fonts", action="store_true")
    apply_parser.set_defaults(handler=apply)

    finalize_parser = subparsers.add_parser("finalize")
    finalize_parser.add_argument("--work-dir", required=True, type=Path)
    finalize_parser.add_argument("--resume", action="store_true")
    finalize_parser.add_argument(
        "--field-updater",
        choices=("auto", "external", "libreoffice", "deferred"),
        default="auto",
    )
    finalize_parser.add_argument("--field-updater-command")
    finalize_parser.add_argument("--target-software")
    finalize_parser.add_argument("--renderer")
    finalize_parser.add_argument("--approve-deferred", action="store_true")
    finalize_parser.set_defaults(handler=finalize)

    verify_parser = subparsers.add_parser("verify")
    verify_parser.add_argument("--work-dir", required=True, type=Path)
    verify_parser.add_argument("--resume", action="store_true")
    verify_parser.add_argument("--renderer")
    verify_parser.add_argument("--target-software")
    verify_parser.add_argument("--visual-qa-manifest", type=Path)
    verify_parser.set_defaults(handler=verify)

    status_parser = subparsers.add_parser("status")
    status_parser.add_argument("--work-dir", required=True, type=Path)
    status_parser.add_argument("--resume", action="store_true")
    status_parser.add_argument("--json", action="store_true")
    status_parser.set_defaults(handler=status)
    return result


def main() -> int:
    args = parser().parse_args()
    state: dict[str, Any] | None = None
    work_dir = getattr(args, "work_dir", None)
    try:
        if work_dir and state_path(work_dir.resolve()).is_file():
            state = read_json(state_path(work_dir.resolve()))
        return int(args.handler(args))
    except Exception as exc:
        if work_dir:
            resolved_work_dir = work_dir.resolve()
            current_state_path = state_path(resolved_work_dir)
            if current_state_path.is_file():
                state = read_json(current_state_path)
            safe_failure(resolved_work_dir, state, args.command, str(exc))
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
