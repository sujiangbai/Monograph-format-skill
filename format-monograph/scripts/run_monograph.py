#!/usr/bin/env python3
"""Portable, resumable orchestration for whole-book formatting runs."""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from field_completion import (
    completion_evidence,
    final_ready_evidence_errors,
    final_ready_evidence_valid,
    finalization_evidence_shape_errors,
)
from backend_evidence import (
    BackendEvidenceError,
    backend_audit_path,
    read_bound_backend_audit,
)
from _common import field_cache_inventory
from external_command import (
    external_command_cache_reusable,
    external_command_identity,
    external_command_identity_errors,
)
from target_software import LIBREOFFICE, MICROSOFT_WORD, UNSUPPORTED, resolve_target_id


SCRIPT_DIR = Path(__file__).resolve().parent
STATE_NAME = "run-state.json"
STATE_SCHEMA_VERSION = "1.0"
FINALIZATION_GATE_VERSION = 2
REQUEST_IDENTITY_VERSION = 2
VERIFICATION_OUTPUT_VERSION = 3
DELIVERY_STATES = {
    "analysis_only",
    "prepared",
    "blocked_qa",
    "candidate_ready",
    "final_ready",
    "failed",
}
RESOLVED_QA_STATES = {"accepted", "resolved", "closed"}
class RunError(RuntimeError):
    """A safe, user-facing orchestration failure."""


def final_ready_field_evidence(evidence: dict[str, Any]) -> bool:
    """Compatibility wrapper around the canonical full-evidence validator."""
    return final_ready_evidence_valid(evidence)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise RunError(f"Cannot read artifact for SHA-256: {path}") from exc
    return digest.hexdigest()


def json_sha256(value: Any) -> str:
    payload = json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (
        OSError,
        RuntimeError,
        TypeError,
        ValueError,
        UnicodeDecodeError,
        json.JSONDecodeError,
    ) as exc:
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
    if (
        not isinstance(value, str)
        or not value
        or any(ord(character) < 32 or ord(character) == 127 for character in value)
    ):
        return None
    try:
        path = Path(value)
        return path if path.is_absolute() else work_dir / path
    except (OSError, RuntimeError, TypeError, ValueError):
        return None


def relative_artifact(work_dir: Path, path: Path) -> str:
    try:
        return str(
            path.resolve(strict=False).relative_to(work_dir.resolve(strict=False))
        )
    except ValueError:
        return str(path.resolve(strict=False))
    except (OSError, RuntimeError, TypeError) as exc:
        raise RunError("Artifact path cannot be resolved") from exc


def local_artifact_identity(path: Path) -> dict[str, Any]:
    try:
        return {
            "path": str(path.resolve(strict=False)),
            "sha256": file_sha256(path),
            "size_bytes": path.stat().st_size,
        }
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise RunError("Artifact identity path is invalid") from exc


def canonical_target_software(value: object) -> str:
    """Compatibility name for the single allowlisted target-ID resolver."""
    return resolve_target_id(value)


def _version_is(value: object, expected: int) -> bool:
    return type(value) is int and value == expected


def renderer_request_identity(requested: str | None) -> dict[str, Any]:
    """Resolve a stable request identity without claiming application authenticity."""
    source = "argument"
    candidate = requested
    if not candidate:
        source = "environment"
        candidate = os.environ.get("FORMAT_MONOGRAPH_RENDERER")
    if not candidate:
        source = "application"
        application = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        candidate = str(application) if application.is_file() else None
    if not candidate:
        source = "path"
        candidate = shutil.which("soffice") or shutil.which("libreoffice")
    identity: dict[str, Any] = {
        "version": REQUEST_IDENTITY_VERSION,
        "requested": str(requested) if requested is not None else None,
        "source": source if candidate else "unavailable",
        "resolved_path": None,
        "sha256": None,
        "size_bytes": None,
        "executable": False,
    }
    if not candidate:
        return identity
    try:
        if any(
            ord(character) < 32 or ord(character) == 127
            for character in str(candidate)
        ):
            raise ValueError("control character")
        path = Path(candidate).expanduser()
        if not path.is_file() and os.sep not in candidate:
            located = shutil.which(candidate)
            if located:
                path = Path(located)
        if path.is_file():
            resolved = path.resolve(strict=False)
            identity.update(
                {
                    "resolved_path": str(resolved),
                    "sha256": file_sha256(resolved),
                    "size_bytes": resolved.stat().st_size,
                    "executable": os.access(resolved, os.X_OK),
                }
            )
        else:
            identity["resolved_path"] = str(path.resolve(strict=False))
    except (OSError, RuntimeError, TypeError, ValueError, RunError):
        identity["source"] = "invalid_path"
    return identity


def renderer_identity_errors(label: str, identity: object) -> list[str]:
    if not isinstance(identity, dict) or not _version_is(
        identity.get("version"), REQUEST_IDENTITY_VERSION
    ):
        return [f"{label} renderer identity is missing or unsupported"]
    resolved = identity.get("resolved_path")
    if resolved is None:
        return [] if identity.get("source") == "unavailable" else [
            f"{label} renderer path is missing"
        ]
    errors = []
    try:
        resolved_text = str(resolved)
        if any(
            ord(character) < 32 or ord(character) == 127
            for character in resolved_text
        ):
            raise ValueError("control character")
        path = Path(resolved_text)
        if not path.is_file():
            return [f"{label} renderer executable is missing"]
        if identity.get("executable") is not True or not os.access(path, os.X_OK):
            errors.append(f"{label} renderer is not executable")
        if file_sha256(path) != identity.get("sha256"):
            errors.append(f"{label} renderer executable SHA-256 changed")
        if path.stat().st_size != identity.get("size_bytes"):
            errors.append(f"{label} renderer executable size changed")
    except (OSError, RuntimeError, TypeError, ValueError, RunError):
        errors.append(f"{label} renderer path cannot be inspected")
    return errors


def _profile_target_id(profile_path: Path) -> str:
    profile = read_json(profile_path)
    targets = profile.get("target_applications")
    if not isinstance(targets, list) or not targets:
        return UNSUPPORTED
    return resolve_target_id(targets[0])


def _finalize_uses_renderer(args: argparse.Namespace, formatted: Path | None) -> bool:
    potential = bool(
        args.field_updater == "libreoffice"
        or (args.field_updater == "auto" and not args.field_updater_command)
    )
    if not potential:
        return False
    if formatted is None or not formatted.is_file():
        return False
    if field_cache_inventory(formatted).get("status") in {"absent", "refreshed"}:
        return False
    return True


def finalize_request_identity(
    args: argparse.Namespace,
    *,
    target_id: str | None = None,
    renderer_used: bool | None = None,
) -> dict[str, Any]:
    effective_target = target_id or resolve_target_id(args.target_software)
    external_used = bool(
        args.field_updater == "external"
        or (args.field_updater == "auto" and args.field_updater_command)
    )
    if renderer_used is None:
        renderer_used = bool(
            args.field_updater == "libreoffice"
            or (args.field_updater == "auto" and not args.field_updater_command)
        )
    return {
        "version": REQUEST_IDENTITY_VERSION,
        "field_updater": args.field_updater,
        "external_updater": (
            external_command_identity(args.field_updater_command, SCRIPT_DIR.parent)
            if external_used
            else None
        ),
        "target_software": effective_target,
        "renderer": renderer_request_identity(args.renderer) if renderer_used else None,
        "approve_deferred": bool(args.approve_deferred),
        "persistent_pdf_requested": True,
        "force_output": True,
    }


def verify_request_identity(
    args: argparse.Namespace,
    *,
    target_id: str | None = None,
    renderer_used: bool | None = None,
) -> dict[str, Any]:
    effective_target = target_id or resolve_target_id(args.target_software)
    if renderer_used is None:
        renderer_used = effective_target != MICROSOFT_WORD
    return {
        "version": REQUEST_IDENTITY_VERSION,
        "target_software": effective_target,
        "renderer": renderer_request_identity(args.renderer) if renderer_used else None,
        "dpi": 150,
        "keep_pdf": True,
        "force_output": True,
    }


def finalize_stage_input_key(
    formatted_sha256: str, structure_sha256: str, request: dict[str, Any]
) -> str:
    return json_sha256(
        {
            "formatted": formatted_sha256,
            "map": structure_sha256,
            "behavior": request,
        }
    )


def finalize_request_cache_reusable(request: dict[str, Any]) -> bool:
    """Return whether a completed finalize stage may be reused safely."""
    external = request.get("external_updater")
    return external is None or external_command_cache_reusable(external)


def verify_stage_input_key(
    finalized_sha256: str,
    structure_sha256: object,
    target_pdf_sha256: str | None,
    visual_manifest: Path | None,
    request: dict[str, Any],
) -> str:
    visual_input = (
        {
            "path": str(visual_manifest.resolve()),
            "sha256": file_sha256(visual_manifest),
        }
        if visual_manifest is not None
        else None
    )
    return json_sha256(
        {
            "finalized": finalized_sha256,
            "map": structure_sha256,
            "target_pdf": target_pdf_sha256,
            "visual_manifest": visual_input,
            "behavior": request,
        }
    )


def _normalized_evidence_path(work_dir: Path, value: object) -> str | None:
    if (
        not isinstance(value, str)
        or not value
        or any(ord(character) < 32 or ord(character) == 127 for character in value)
    ):
        return None
    try:
        path = Path(value)
        if not path.is_absolute():
            path = work_dir / path
        return str(path.resolve(strict=False))
    except (OSError, RuntimeError, TypeError, ValueError):
        return None


def _gate_artifact_identity(work_dir: Path, value: object) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        return {"path": None, "sha256": None, "size_bytes": None}
    result = {
        "path": _normalized_evidence_path(work_dir, value.get("path")),
        "sha256": value.get("sha256"),
        "size_bytes": value.get("size_bytes"),
    }
    if "page_count" in value:
        result["page_count"] = value.get("page_count")
    return result


def canonical_finalization_gate_summary(
    work_dir: Path, finalization: dict[str, Any]
) -> dict[str, Any]:
    """Project only explicit final-ready gate fields from finalization evidence."""
    workflow = finalization.get("workflow_state")
    if not isinstance(workflow, dict):
        workflow = {}
    binding = finalization.get("artifact_binding")
    if not isinstance(binding, dict):
        binding = {}
    backend = finalization.get("field_backend")
    if not isinstance(backend, dict):
        backend = {}
    verification = backend.get("read_only_verification")
    if not isinstance(verification, dict):
        verification = {}
    field_completion = finalization.get("field_completion")
    if not isinstance(field_completion, dict):
        field_completion = {}
    validation = field_completion.get("evidence_validation")
    if not isinstance(validation, dict):
        validation = {}
    return {
        "version": FINALIZATION_GATE_VERSION,
        "status": finalization.get("status"),
        "integrity": {
            "content": finalization.get("content_integrity"),
            "protected_objects": finalization.get("protected_object_integrity"),
            "effective_fonts": finalization.get("effective_font_integrity"),
        },
        "workflow": {
            "stage": workflow.get("stage"),
            "source_sha256": workflow.get("source_sha256"),
            "input_sha256": workflow.get("input_sha256"),
            "profile_sha256": workflow.get("profile_sha256"),
            "structure_map_sha256": workflow.get("structure_map_sha256"),
            "output_sha256": workflow.get("output_sha256"),
        },
        "output_path": _normalized_evidence_path(
            work_dir, finalization.get("output")
        ),
        "target": {
            "software": canonical_target_software(
                finalization.get("target_software")
            ),
            "pdf_path": _normalized_evidence_path(
                work_dir, finalization.get("target_pdf")
            ),
            "layout_status": finalization.get("target_layout_status"),
            "backend_software": resolve_target_id(
                backend.get("target_id") or backend.get("software")
            ),
            "verification_software": resolve_target_id(
                verification.get("target_id") or verification.get("software")
            ),
        },
        "artifacts": {
            "version": binding.get("version"),
            "finalized_docx": _gate_artifact_identity(
                work_dir, binding.get("finalized_docx")
            ),
            "word_verification_pdf": _gate_artifact_identity(
                work_dir, binding.get("word_verification_pdf")
            ),
        },
        "backend_audit": {
            "version": (finalization.get("backend_audit") or {}).get("version"),
            "status": (finalization.get("backend_audit") or {}).get("status"),
            "artifact": _gate_artifact_identity(
                work_dir,
                (finalization.get("backend_audit") or {}).get("artifact"),
            ),
        },
        "field_completion": completion_evidence(finalization),
        "evidence_validation": {
            "status": validation.get("status"),
            "errors": validation.get("errors"),
        },
    }


def artifact_identity_errors(
    label: str, identity: Any, expected_path: Path | None
) -> list[str]:
    if not isinstance(identity, dict) or not identity.get("path"):
        return [f"{label} output identity is missing"]
    errors = []
    path_text = identity.get("path")
    if (
        not isinstance(path_text, str)
        or not path_text
        or any(
            ord(character) < 32 or ord(character) == 127
            for character in path_text
        )
    ):
        return [f"{label} output path is invalid"]
    try:
        path = Path(path_text).resolve(strict=False)
        if expected_path is None or expected_path.resolve(strict=False) != path:
            errors.append(f"{label} output path differs from state")
        if not path.is_file():
            errors.append(f"{label} output is missing")
            return errors
        if file_sha256(path) != identity.get("sha256"):
            errors.append(f"{label} output SHA-256 differs from state")
        if path.stat().st_size != identity.get("size_bytes"):
            errors.append(f"{label} output size differs from state")
    except (OSError, RuntimeError, TypeError, ValueError, RunError):
        errors.append(f"{label} output path cannot be inspected")
    return errors


def artifact_binding_errors(
    work_dir: Path, finalization: dict[str, Any], state: dict[str, Any]
) -> list[str]:
    """Revalidate bound artifacts against both finalization and state evidence."""
    errors: list[str] = []
    binding = finalization.get("artifact_binding")
    if not isinstance(binding, dict):
        return ["finalization artifact_binding is missing"]
    stored = state.get("field_writeback", {}).get("artifact_binding")
    if stored != binding:
        errors.append("stored artifact binding differs from finalization evidence")

    for key, state_artifact in (
        ("finalized_docx", "finalized"),
        ("word_verification_pdf", "target_pdf"),
    ):
        identity = binding.get(key)
        if identity is None:
            if artifact_path(
                work_dir, state.get("artifacts", {}).get(state_artifact)
            ) is not None:
                errors.append(
                    f"{key} is absent from finalization but present in state artifacts"
                )
            continue
        if not isinstance(identity, dict) or not identity.get("path"):
            errors.append(f"{key} identity is incomplete")
            continue
        state_path_value = artifact_path(
            work_dir, state.get("artifacts", {}).get(state_artifact)
        )
        path_text = identity.get("path")
        if (
            not isinstance(path_text, str)
            or not path_text
            or any(
                ord(character) < 32 or ord(character) == 127
                for character in path_text
            )
        ):
            errors.append(f"{key} path is invalid")
            continue
        try:
            bound_path = Path(path_text).resolve(strict=False)
            if (
                state_path_value is None
                or state_path_value.resolve(strict=False) != bound_path
            ):
                errors.append(f"{key} path differs from the state artifact path")
            if not bound_path.is_file():
                errors.append(f"{key} artifact is missing")
                continue
            if file_sha256(bound_path) != identity.get("sha256"):
                errors.append(f"{key} SHA-256 differs from finalization evidence")
            if bound_path.stat().st_size != identity.get("size_bytes"):
                errors.append(f"{key} size differs from finalization evidence")
        except (OSError, RuntimeError, TypeError, ValueError, RunError):
            errors.append(f"{key} path cannot be inspected")

    finalized_identity = binding.get("finalized_docx") or {}
    workflow_hash = finalization.get("workflow_state", {}).get("output_sha256")
    if workflow_hash != finalized_identity.get("sha256"):
        errors.append("workflow output SHA-256 differs from finalized DOCX binding")
    return errors


def backend_audit_binding_errors(
    work_dir: Path, finalization: dict[str, Any], state: dict[str, Any]
) -> list[str]:
    binding = finalization.get("backend_audit")
    if not isinstance(binding, dict) or binding.get("status") != "persisted":
        return ["backend audit is not persistently bound"]
    identity = binding.get("artifact")
    expected_path = artifact_path(
        work_dir, state.get("artifacts", {}).get("backend_audit")
    )
    errors = artifact_identity_errors("backend_audit", identity, expected_path)
    if errors:
        return errors
    try:
        read_bound_backend_audit(finalization)
    except BackendEvidenceError as exc:
        errors.append(str(exc))
    return errors


def fresh_backend_audit_errors(
    finalization: dict[str, Any], final_status: Path
) -> list[str]:
    """Validate a new sidecar against the orchestrator-derived output path."""
    try:
        read_bound_backend_audit(
            finalization, expected_path=backend_audit_path(final_status)
        )
    except BackendEvidenceError as exc:
        return [str(exc)]
    return []


def _current_file_hash_errors(
    label: str, path: Path | None, state_sha256: object, gate_sha256: object
) -> list[str]:
    errors: list[str] = []
    try:
        if path is None or not path.is_file():
            return [f"current {label} artifact is missing"]
        current_hash = file_sha256(path)
    except (OSError, RuntimeError, TypeError, ValueError, RunError):
        return [f"current {label} artifact path cannot be inspected"]
    if current_hash != state_sha256:
        errors.append(f"current {label} SHA-256 differs from state")
    if current_hash != gate_sha256:
        errors.append(f"current {label} SHA-256 differs from finalization workflow")
    return errors


def finalization_gate_errors(
    work_dir: Path,
    gate: dict[str, Any],
    state: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    if not _version_is(gate.get("version"), FINALIZATION_GATE_VERSION):
        return ["finalization gate summary version is unsupported"]
    if gate.get("status") != "pass":
        errors.append("finalization status is not pass")
    integrity = gate.get("integrity") or {}
    for key in ("content", "protected_objects", "effective_fonts"):
        if integrity.get(key) != "pass":
            errors.append(f"finalization {key} integrity is not pass")
    workflow = gate.get("workflow") or {}
    if workflow.get("stage") != "finalized":
        errors.append("finalization workflow stage is not finalized")

    source_value = state.get("source") or {}
    profile_value = state.get("profile") or {}
    structure_value = state.get("structure_map") or {}
    artifacts = state.get("artifacts") or {}
    source_path = Path(source_value["path"]) if source_value.get("path") else None
    profile_path = Path(profile_value["path"]) if profile_value.get("path") else None
    structure_path = (
        Path(structure_value["path"]) if structure_value.get("path") else None
    )
    formatted_path = artifact_path(work_dir, artifacts.get("formatted"))
    finalized_path = artifact_path(work_dir, artifacts.get("finalized"))
    errors.extend(
        _current_file_hash_errors(
            "source",
            source_path,
            source_value.get("sha256"),
            workflow.get("source_sha256"),
        )
    )
    errors.extend(
        _current_file_hash_errors(
            "profile",
            profile_path,
            profile_value.get("sha256"),
            workflow.get("profile_sha256"),
        )
    )
    errors.extend(
        _current_file_hash_errors(
            "structure map",
            structure_path,
            structure_value.get("sha256"),
            workflow.get("structure_map_sha256"),
        )
    )
    errors.extend(
        _current_file_hash_errors(
            "formatted input",
            formatted_path,
            workflow.get("input_sha256"),
            workflow.get("input_sha256"),
        )
    )
    finalized_binding = (gate.get("artifacts") or {}).get("finalized_docx") or {}
    errors.extend(
        _current_file_hash_errors(
            "finalized DOCX",
            finalized_path,
            finalized_binding.get("sha256"),
            workflow.get("output_sha256"),
        )
    )
    try:
        expected_output_path = (
            str(finalized_path.resolve(strict=False))
            if finalized_path is not None
            else None
        )
    except (OSError, RuntimeError, TypeError, ValueError):
        expected_output_path = None
        errors.append("state finalized artifact path cannot be resolved")
    if gate.get("output_path") != expected_output_path:
        errors.append("finalization output path differs from state finalized artifact")

    completion = gate.get("field_completion") or {}
    recalculated_evidence_errors = final_ready_evidence_errors(completion)
    validation = gate.get("evidence_validation") or {}
    if validation.get("status") != "pass":
        errors.append("field completion evidence validation is not pass")
    if validation.get("errors") != []:
        errors.append("field completion evidence validation errors are not empty")
    expected_validation = {
        "status": "pass" if not recalculated_evidence_errors else "incomplete",
        "errors": recalculated_evidence_errors,
    }
    if validation != expected_validation:
        errors.append(
            "field completion evidence validation differs from recalculated evidence"
        )
    target = gate.get("target") or {}
    pdf_binding = (gate.get("artifacts") or {}).get("word_verification_pdf")
    state_pdf = artifact_path(work_dir, artifacts.get("target_pdf"))
    delivery = completion.get("delivery_status")
    if delivery == "absent":
        if pdf_binding is not None or target.get("pdf_path") is not None:
            errors.append("no-fields finalization unexpectedly has a target PDF")
        if target.get("layout_status") != "not_verified":
            errors.append("no-fields finalization target layout status is inconsistent")
    elif delivery == "selective_verified":
        if not isinstance(pdf_binding, dict):
            errors.append("Word finalization lacks a bound verification PDF")
        else:
            if target.get("pdf_path") != pdf_binding.get("path"):
                errors.append("target PDF path differs from Word artifact binding")
            try:
                resolved_state_pdf = (
                    str(state_pdf.resolve(strict=False))
                    if state_pdf is not None
                    else None
                )
            except (OSError, RuntimeError, TypeError, ValueError):
                resolved_state_pdf = None
            if target.get("pdf_path") != resolved_state_pdf:
                errors.append("target PDF path differs from state target artifact")
        if target.get("layout_status") != "target_pdf_ready_for_visual_qa":
            errors.append("Word target layout status is inconsistent")
        if target.get("software") != MICROSOFT_WORD:
            errors.append("Word completion target software is inconsistent")
        if target.get("backend_software") != MICROSOFT_WORD:
            errors.append("Word backend target software is inconsistent")
        if target.get("verification_software") != MICROSOFT_WORD:
            errors.append("Word verification target software is inconsistent")
    if target.get("software") == UNSUPPORTED:
        errors.append("finalization target software is unsupported")
    return errors


def finalization_request_gate_errors(
    work_dir: Path, gate: dict[str, Any], state: dict[str, Any]
) -> list[str]:
    request = state.get("finalization_request")
    if not isinstance(request, dict) or not _version_is(
        request.get("version"), REQUEST_IDENTITY_VERSION
    ):
        return ["finalization request binding is missing or unsupported"]
    requested_target = request.get("target_software") or ""
    gate_target = (gate.get("target") or {}).get("software") or ""
    if requested_target and requested_target != gate_target:
        errors = ["finalization target software differs from request binding"]
    else:
        errors = []
    if request.get("renderer") is not None:
        errors.extend(
            renderer_identity_errors("finalization request", request.get("renderer"))
        )
    if request.get("external_updater") is not None:
        errors.extend(
            external_command_identity_errors(
                "finalization request", request.get("external_updater")
            )
        )
    formatted = artifact_path(
        work_dir, state.get("artifacts", {}).get("formatted")
    )
    structure = state.get("structure_map") or {}
    if formatted is None or not formatted.is_file():
        errors.append("finalization request cannot bind a missing formatted input")
    elif not isinstance(structure.get("sha256"), str):
        errors.append("finalization request cannot bind a missing structure hash")
    else:
        expected_key = finalize_stage_input_key(
            file_sha256(formatted), structure["sha256"], request
        )
        if (
            state.get("stages", {})
            .get("finalize", {})
            .get("input_key_sha256")
            != expected_key
        ):
            errors.append("finalization request differs from finalized stage input key")
    return errors


def finalization_consistency_errors(
    work_dir: Path, finalization: dict[str, Any], state: dict[str, Any]
) -> tuple[dict[str, Any], list[str]]:
    shape_errors = finalization_evidence_shape_errors(finalization)
    if shape_errors:
        return {}, shape_errors
    canonical = completion_evidence(finalization)
    errors = final_ready_evidence_errors(canonical)
    stored = state.get("field_writeback", {}).get("completion_evidence")
    if stored != canonical:
        errors.append(
            "stored field completion evidence differs from finalization evidence"
        )
    gate = canonical_finalization_gate_summary(work_dir, finalization)
    if state.get("finalization_gate") != gate:
        errors.append("stored finalization gate differs from finalization evidence")
    errors.extend(finalization_gate_errors(work_dir, gate, state))
    errors.extend(finalization_request_gate_errors(work_dir, gate, state))
    errors.extend(artifact_binding_errors(work_dir, finalization, state))
    errors.extend(backend_audit_binding_errors(work_dir, finalization, state))
    return canonical, errors


def load_finalization_consistency(
    work_dir: Path, state: dict[str, Any]
) -> tuple[dict[str, Any] | None, dict[str, Any] | None, list[str]]:
    status_path = artifact_path(
        work_dir, state.get("artifacts", {}).get("finalization_status")
    )
    if status_path is None or not status_path.is_file():
        return None, None, ["finalization status evidence is missing"]
    try:
        finalization = read_json(status_path)
    except RunError as exc:
        return None, None, [str(exc)]
    completion, errors = finalization_consistency_errors(
        work_dir, finalization, state
    )
    return finalization, completion, errors


def render_page_count_errors(
    completion: dict[str, Any], render_page_count: int
) -> list[str]:
    if (
        completion.get("delivery_status") == "selective_verified"
        and render_page_count != completion.get("verification_page_count")
    ):
        return [
            "render manifest page_count differs from Word verification page_count"
        ]
    return []


def verification_target_evidence_errors(
    work_dir: Path,
    state: dict[str, Any],
    completion: dict[str, Any],
    render_manifest: dict[str, Any],
    request_identity: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    requested_target = request_identity.get("target_software") or ""
    manifest_target = resolve_target_id(render_manifest.get("target_software"))
    if manifest_target != requested_target:
        errors.append("render target software differs from verification request")
    target_pdf = artifact_path(
        work_dir, state.get("artifacts", {}).get("target_pdf")
    )
    manifest_pdf = _normalized_evidence_path(
        work_dir, render_manifest.get("target_pdf_source")
    )
    if completion.get("delivery_status") == "selective_verified":
        if requested_target != MICROSOFT_WORD:
            errors.append("Word completion verification target is inconsistent")
        try:
            expected_pdf = (
                str(target_pdf.resolve(strict=False))
                if target_pdf is not None
                else None
            )
        except (OSError, RuntimeError, TypeError, ValueError):
            expected_pdf = None
            errors.append("state target PDF path cannot be resolved")
        if manifest_pdf != expected_pdf:
            errors.append("render target PDF source differs from Word verification PDF")
        if render_manifest.get("renderer") is not None:
            errors.append("target-PDF render unexpectedly reports a renderer")
        if render_manifest.get("renderer_source") != "target_pdf":
            errors.append("Word verification render did not report target_pdf source")
    else:
        if manifest_pdf is not None:
            errors.append("non-Word verification unexpectedly uses a target PDF")
        renderer = request_identity.get("renderer") or {}
        expected_renderer = renderer.get("resolved_path")
        manifest_renderer = _normalized_evidence_path(
            work_dir, render_manifest.get("renderer")
        )
        if not renderer:
            errors.append("rendered verification lacks a renderer request identity")
        elif manifest_renderer != expected_renderer:
            errors.append("rendered application differs from renderer request identity")
        if render_manifest.get("renderer_source") != renderer.get("source"):
            errors.append("render source differs from renderer request identity")
    return errors


def verification_output_consistency_errors(
    work_dir: Path,
    state: dict[str, Any],
    completion: dict[str, Any],
    *,
    requested_visual_manifest: Path | None = None,
    requested_identity: dict[str, Any] | None = None,
    require_final_ready: bool = False,
) -> list[str]:
    errors: list[str] = []
    outputs = state.get("verification_outputs")
    if (
        not isinstance(outputs, dict)
        or not _version_is(outputs.get("version"), VERIFICATION_OUTPUT_VERSION)
    ):
        return ["verification output bindings are missing or unsupported"]
    recorded_request = outputs.get("request")
    if not isinstance(recorded_request, dict):
        errors.append("verification request binding is missing")
        recorded_request = {}
    if requested_identity is not None and recorded_request != requested_identity:
        errors.append("verification request differs from cached request binding")
    if recorded_request.get("renderer") is not None:
        errors.extend(
            renderer_identity_errors(
                "verification request", recorded_request.get("renderer")
            )
        )
    artifacts = state.get("artifacts", {})
    audit_path = artifact_path(work_dir, artifacts.get("final_audit"))
    render_path = artifact_path(work_dir, artifacts.get("render_manifest"))
    visual_path = artifact_path(work_dir, artifacts.get("visual_qa_manifest"))
    errors.extend(
        artifact_identity_errors("final_audit", outputs.get("final_audit"), audit_path)
    )
    errors.extend(
        artifact_identity_errors(
            "render_manifest", outputs.get("render_manifest"), render_path
        )
    )
    recorded_visual = outputs.get("visual_manifest")
    if recorded_visual is None:
        if visual_path is not None or requested_visual_manifest is not None:
            errors.append("visual manifest binding is missing")
    else:
        errors.extend(
            artifact_identity_errors(
                "visual_manifest", recorded_visual, visual_path
            )
        )
        if requested_visual_manifest is not None:
            try:
                visual_differs = (
                    visual_path is None
                    or requested_visual_manifest.resolve(strict=False)
                    != visual_path.resolve(strict=False)
                )
            except (OSError, RuntimeError, TypeError, ValueError):
                visual_differs = True
            if visual_differs:
                errors.append(
                    "requested visual manifest differs from cached manifest"
                )

    finalized_path = artifact_path(work_dir, artifacts.get("finalized"))
    target_pdf_path = artifact_path(work_dir, artifacts.get("target_pdf"))
    uses_target_pdf = bool(
        completion.get("delivery_status") == "selective_verified"
        and recorded_request.get("target_software") == MICROSOFT_WORD
    )
    structure = state.get("structure_map") or {}
    try:
        finalized_exists = bool(finalized_path and finalized_path.is_file())
        target_pdf_exists = bool(target_pdf_path and target_pdf_path.is_file())
        visual_exists = bool(visual_path and visual_path.is_file())
    except (OSError, RuntimeError, TypeError, ValueError):
        finalized_exists = target_pdf_exists = visual_exists = False
        errors.append("verification artifact path cannot be inspected")
    if not finalized_exists:
        errors.append("verification request cannot bind a missing finalized DOCX")
    elif not isinstance(structure.get("sha256"), str):
        errors.append("verification request cannot bind a missing structure hash")
    elif uses_target_pdf and not target_pdf_exists:
        errors.append("verification request cannot bind a missing target PDF")
    elif visual_path is not None and not visual_exists:
        errors.append("verification request cannot bind a missing visual manifest")
    else:
        expected_key = verify_stage_input_key(
            file_sha256(finalized_path),
            structure["sha256"],
            file_sha256(target_pdf_path)
            if uses_target_pdf and target_pdf_path is not None
            else None,
            visual_path,
            recorded_request,
        )
        if (
            state.get("stages", {})
            .get("verify", {})
            .get("input_key_sha256")
            != expected_key
        ):
            errors.append("verification request differs from verify stage input key")

    render_manifest = None
    try:
        render_exists = bool(render_path and render_path.is_file())
    except (OSError, RuntimeError, TypeError, ValueError):
        render_exists = False
        errors.append("render manifest path cannot be inspected")
    if render_exists:
        try:
            render_manifest = read_json(render_path)
        except RunError as exc:
            errors.append(str(exc))
    if render_manifest is not None:
        errors.extend(
            verification_target_evidence_errors(
                work_dir,
                state,
                completion,
                render_manifest,
                recorded_request,
            )
        )
        page_count = render_manifest.get("page_count")
        if (
            not isinstance(page_count, int)
            or isinstance(page_count, bool)
            or page_count < 1
        ):
            errors.append("render manifest page_count is invalid")
        else:
            errors.extend(render_page_count_errors(completion, page_count))
            if visual_path is not None and visual_exists:
                try:
                    validate_visual_manifest(visual_path, page_count)
                except RunError as exc:
                    errors.append(str(exc))
            visual_state = state.get("visual_qa", {})
            if visual_path is not None and (
                visual_state.get("all_pages_inspected") is not True
                or visual_state.get("target_layout_verified") is not True
                or visual_state.get("page_count") != page_count
            ):
                errors.append("stored visual QA state differs from current manifests")
        if require_final_ready and not has_target_layout_evidence(render_manifest):
            errors.append("render manifest lacks target-layout evidence")
    if require_final_ready:
        if state.get("status") != "final_ready":
            errors.append("run state is not final_ready")
        if state.get("blockers"):
            errors.append("final_ready state contains blockers")
        if visual_path is None or recorded_visual is None:
            errors.append("final_ready state lacks a bound visual manifest")
        errors.extend(final_ready_evidence_errors(completion))
    return errors


def downgrade_invalid_final_ready(
    work_dir: Path, state: dict[str, Any], errors: list[str], source: str
) -> None:
    state["status"] = "candidate_ready"
    state["blockers"] = [
        {
            "id": f"local-evidence:{source}",
            "kind": "field_update",
            "status": "open",
        }
    ]
    state["local_evidence_validation"] = {
        "status": "invalid",
        "source": source,
        "errors": list(dict.fromkeys(errors)),
        "checked_at": utc_now(),
    }
    save_state(work_dir, state)


def invalidate_verification_state(state: dict[str, Any]) -> None:
    """Discard downstream evidence after a real finalization execution."""
    stages = state.get("stages")
    if isinstance(stages, dict):
        stages.pop("verify", None)
    state.pop("verification_outputs", None)
    artifacts = state.get("artifacts")
    if isinstance(artifacts, dict):
        for name in ("final_audit", "render_manifest", "visual_qa_manifest"):
            artifacts.pop(name, None)
    state.pop("visual_qa", None)
    state.pop("local_evidence_validation", None)
    metrics = state.get("metrics")
    if isinstance(metrics, dict):
        metrics.pop("rendered_pages", None)


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
    try:
        source_text = state["source"]["path"]
        profile_text = state["profile"]["path"]
        if any(
            ord(character) < 32 or ord(character) == 127
            for character in str(source_text) + str(profile_text)
        ):
            raise ValueError("control character")
        source = Path(source_text)
        profile = Path(profile_text)
        if not source.is_file() or file_sha256(source) != state["source"]["sha256"]:
            raise RunError("Source changed after prepare; run prepare again.")
        if not profile.is_file() or file_sha256(profile) != state["profile"]["sha256"]:
            raise RunError("Profile changed after prepare; run prepare again.")
    except RunError:
        raise
    except (KeyError, OSError, RuntimeError, TypeError, ValueError) as exc:
        raise RunError("Stored source or profile path is invalid.") from exc
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
    structure_path = artifact_path(work_dir, structure.get("path"))
    try:
        structure_valid = bool(
            structure_path
            and structure_path.is_file()
            and file_sha256(structure_path) == structure.get("sha256")
        )
    except (OSError, RuntimeError, TypeError, ValueError, RunError):
        structure_valid = False
    if not structure_valid:
        raise RunError("Approved structure map changed; run apply again.")
    formatted = artifact_path(work_dir, state.get("artifacts", {}).get("formatted"))
    try:
        formatted_valid = bool(formatted and formatted.is_file())
    except (OSError, RuntimeError, TypeError, ValueError):
        formatted_valid = False
    if not formatted_valid:
        raise RunError("Formatted candidate was not found. Run apply first.")
    final_dir = work_dir / "final"
    final_docx = final_dir / f"{source.stem}-finalized.docx"
    final_status = final_dir / "finalization.json"
    target_pdf = final_dir / f"{source.stem}-target.pdf"
    target_id = resolve_target_id(
        args.target_software or _profile_target_id(profile)
    )
    renderer_used = _finalize_uses_renderer(args, formatted)
    request_identity = finalize_request_identity(
        args, target_id=target_id, renderer_used=renderer_used
    )
    input_key = finalize_stage_input_key(
        file_sha256(formatted), structure["sha256"], request_identity
    )
    if args.resume and finalize_request_cache_reusable(request_identity) and cached_stage(
        state, "finalize", input_key, [final_docx, final_status]
    ):
        _, completion, cache_errors = load_finalization_consistency(
            work_dir, state
        )
        if state.get("finalization_request") != request_identity:
            cache_errors.append(
                "finalization request differs from cached request binding"
            )
        if state.get("status") == "final_ready" and completion is not None:
            cache_errors.extend(
                verification_output_consistency_errors(
                    work_dir,
                    state,
                    completion,
                    require_final_ready=True,
                )
            )
        if cache_errors:
            downgrade_invalid_final_ready(
                work_dir, state, cache_errors, "finalize_resume"
            )
            print(json.dumps(state, ensure_ascii=False, indent=2))
            return 2
        mark_cache_hit(state, "finalize")
        state["metrics"]["cache_hits"] = int(
            state["metrics"].get("cache_hits", 0)
        ) + 1
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0

    state_before_execution = copy.deepcopy(state)
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
    command.extend(["--target-software", target_id])
    if renderer_used and args.renderer:
        command.extend(["--renderer", args.renderer])
    if args.approve_deferred:
        command.append("--approve-deferred")
    completed = run_script("finalize_docx.py", *command)
    if completed.returncode != 0:
        state.clear()
        state.update(state_before_execution)
        print("finalize_docx.py failed with a nonzero exit code.", file=sys.stderr)
        return 2
    try:
        evidence = read_json(final_status)
    except RunError as exc:
        state.clear()
        state.update(state_before_execution)
        print(str(exc), file=sys.stderr)
        return 2
    shape_errors = finalization_evidence_shape_errors(evidence)
    if shape_errors:
        state.clear()
        state.update(state_before_execution)
        print(
            "Finalization evidence shape is invalid: " + "; ".join(shape_errors),
            file=sys.stderr,
        )
        return 2
    fresh_audit_errors = fresh_backend_audit_errors(evidence, final_status)
    if fresh_audit_errors:
        state.clear()
        state.update(state_before_execution)
        print(
            "Fresh backend audit evidence is invalid: "
            + "; ".join(fresh_audit_errors),
            file=sys.stderr,
        )
        return 2
    invalidate_verification_state(state)
    field_status = str(evidence.get("delivery_field_status", "stale"))
    field_completion = evidence.get("field_completion", {})
    canonical_completion = completion_evidence(evidence)
    state["artifacts"].update(
        {
            "finalized": relative_artifact(work_dir, final_docx),
            "finalization_status": relative_artifact(work_dir, final_status),
        }
    )
    backend_audit_artifact = (evidence.get("backend_audit") or {}).get("artifact")
    if isinstance(backend_audit_artifact, dict) and backend_audit_artifact.get(
        "path"
    ):
        state["artifacts"]["backend_audit"] = relative_artifact(
            work_dir, Path(str(backend_audit_artifact["path"]))
        )
    else:
        state["artifacts"].pop("backend_audit", None)
    if target_pdf.is_file():
        state["artifacts"]["target_pdf"] = relative_artifact(work_dir, target_pdf)
    else:
        state["artifacts"].pop("target_pdf", None)
    state["finalization_request"] = request_identity
    state["finalization_gate"] = canonical_finalization_gate_summary(
        work_dir, evidence
    )
    state["field_writeback"] = {
        "status": evidence.get("field_writeback_status", "not_verified"),
        "delivery_status": field_status,
        "backend": evidence.get("field_backend", {}).get("backend"),
        "matched_fields": (
            evidence.get("field_backend", {}).get("selective_writeback") or {}
        ).get("matched_fields", 0),
        "updated_fields": (
            evidence.get("field_backend", {}).get("selective_writeback") or {}
        ).get("updated_fields", 0),
        "read_only_verified": bool(
            (
                evidence.get("field_backend", {}).get("read_only_verification")
                or {}
            ).get("read_only_verified")
        ),
        "field_gate_completed": False,
        "final_ready_eligible": False,
        "word_verification_required": bool(
            field_completion.get("word_verification_required")
        ),
        "word_verification_completed": bool(
            field_completion.get("word_verification_completed")
        ),
        "completion_scope": field_completion.get("completion_scope", "incomplete"),
        "completion_evidence": canonical_completion,
        "artifact_binding": evidence.get("artifact_binding"),
        "completion_evidence_errors": [],
    }
    _, completion_errors = finalization_consistency_errors(
        work_dir, evidence, state
    )
    final_ready_eligible = not completion_errors
    state["field_writeback"]["field_gate_completed"] = bool(
        final_ready_eligible and field_completion.get("field_gate_completed")
    )
    state["field_writeback"]["final_ready_eligible"] = final_ready_eligible
    state["field_writeback"]["completion_evidence_errors"] = completion_errors
    if not final_ready_eligible:
        state["status"] = "candidate_ready"
        blocker_id = (
            "field-update:word-verification-required"
            if field_status == "libreoffice_refreshed"
            else "field-update:incomplete"
        )
        state["blockers"] = [
            {"id": blocker_id, "kind": "field_update", "status": "open"}
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
    target = resolve_target_id(render_manifest.get("target_software"))
    if render_manifest.get("target_pdf_source"):
        return target == MICROSOFT_WORD
    return bool(
        target == LIBREOFFICE
        and not render_manifest.get("target_layout_unverified")
    )


def verify(args: argparse.Namespace) -> int:
    work_dir = args.work_dir.resolve()
    state = load_state(work_dir)
    source, profile = current_inputs(state)
    finalization, canonical_completion, completion_errors = load_finalization_consistency(
        work_dir, state
    )
    if canonical_completion is None:
        canonical_completion = {}
    if completion_errors:
        state["status"] = "candidate_ready"
        state.setdefault("field_writeback", {})[
            "completion_evidence_errors"
        ] = completion_errors
        if not any(
            item.get("kind") == "field_update"
            for item in state.get("blockers", [])
        ):
            state.setdefault("blockers", []).append(
                {
                    "id": "field-update:word-verification-required",
                    "kind": "field_update",
                    "status": "open",
                }
            )
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2
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
    persisted_target = (
        resolve_target_id(finalization.get("target_software"))
        if isinstance(finalization, dict)
        else UNSUPPORTED
    )
    target_id = resolve_target_id(args.target_software) if args.target_software else persisted_target
    uses_target_pdf = bool(
        target_id == MICROSOFT_WORD
        and persisted_target == target_id
        and canonical_completion.get("delivery_status") == "selective_verified"
        and target_pdf is not None
        and target_pdf.is_file()
    )
    request_identity = verify_request_identity(
        args, target_id=target_id, renderer_used=not uses_target_pdf
    )
    input_key = verify_stage_input_key(
        file_sha256(finalized),
        structure.get("sha256"),
        file_sha256(target_pdf) if uses_target_pdf else None,
        args.visual_qa_manifest,
        request_identity,
    )
    required_artifacts = [audit_path, render_dir / "render-manifest.json"]
    if args.resume and cached_stage(
        state, "verify", input_key, required_artifacts
    ):
        cache_errors = verification_output_consistency_errors(
            work_dir,
            state,
            canonical_completion,
            requested_visual_manifest=(
                args.visual_qa_manifest.resolve()
                if args.visual_qa_manifest is not None
                else None
            ),
            requested_identity=request_identity,
            require_final_ready=state.get("status") == "final_ready",
        )
        if cache_errors:
            downgrade_invalid_final_ready(
                work_dir, state, cache_errors, "verify_resume"
            )
            print(json.dumps(state, ensure_ascii=False, indent=2))
            return 2
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
    if uses_target_pdf:
        render_command.extend(["--target-pdf", target_pdf])
    elif args.renderer:
        render_command.extend(["--renderer", args.renderer])
    render_command.extend(["--target-software", target_id])
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
    state["verification_outputs"] = {
        "version": VERIFICATION_OUTPUT_VERSION,
        "request": request_identity,
        "final_audit": local_artifact_identity(audit_path),
        "render_manifest": local_artifact_identity(
            render_dir / "render-manifest.json"
        ),
        "visual_manifest": None,
    }
    render_evidence_errors = render_page_count_errors(
        canonical_completion, page_count
    )
    render_evidence_errors.extend(
        verification_target_evidence_errors(
            work_dir,
            state,
            canonical_completion,
            render_manifest,
            request_identity,
        )
    )
    if render_evidence_errors:
        state["status"] = "candidate_ready"
        state["blockers"] = [
            {
                "id": "verification-output:target-evidence",
                "kind": "field_update",
                "status": "open",
            }
        ]
        state.setdefault("field_writeback", {})[
            "completion_evidence_errors"
        ] = render_evidence_errors
        finish_stage(state, "verify", started, status="blocked")
        save_state(work_dir, state)
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 2
    if args.visual_qa_manifest:
        visual = validate_visual_manifest(args.visual_qa_manifest.resolve(), page_count)
        state["artifacts"]["visual_qa_manifest"] = str(
            args.visual_qa_manifest.resolve()
        )
        state["verification_outputs"]["visual_manifest"] = (
            local_artifact_identity(args.visual_qa_manifest.resolve())
        )
        state["visual_qa"] = {
            "all_pages_inspected": True,
            "target_layout_verified": True,
            "page_count": page_count,
            "verified_at": visual.get("verified_at") or utc_now(),
        }
        if (
            not final_ready_evidence_errors(canonical_completion)
            and has_target_layout_evidence(render_manifest)
        ):
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
    work_dir = args.work_dir.resolve()
    state = load_state(work_dir)
    if state.get("status") == "final_ready":
        _, completion, errors = load_finalization_consistency(work_dir, state)
        if completion is None:
            completion = {}
        errors.extend(
            verification_output_consistency_errors(
                work_dir,
                state,
                completion,
                require_final_ready=True,
            )
        )
        if errors:
            downgrade_invalid_final_ready(work_dir, state, errors, "status")
            if args.json:
                print(json.dumps(state, ensure_ascii=False, indent=2))
            else:
                print("Run status: candidate_ready (local evidence invalid)")
                print(f"Blockers: {len(state.get('blockers', []))}")
            return 2
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
