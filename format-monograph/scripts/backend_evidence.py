#!/usr/bin/env python3
"""Closed backend evidence projection and bounded diagnostic audit artifacts."""

from __future__ import annotations

import hashlib
import json
import math
import os
import stat
from pathlib import Path
from typing import Any

from target_software import LIBREOFFICE, MICROSOFT_WORD, UNSUPPORTED


CANONICAL_BACKEND_VERSION = 1
CANONICAL_NESTED_VERSION = 1
BACKEND_AUDIT_VERSION = 1
BACKEND_AUDIT_BINDING_VERSION = 1
BACKEND_AUDIT_MAX_BYTES = 1024 * 1024
BACKEND_AUDIT_MAX_DEPTH = 32
BACKEND_AUDIT_MAX_NODES = 100_000
BACKEND_AUDIT_MAX_STRING_LENGTH = 65_536

BACKENDS = frozenset(
    {"not_needed", "external", "deferred_on_open", "libreoffice_uno"}
)
TARGET_IDS = frozenset({MICROSOFT_WORD, LIBREOFFICE, UNSUPPORTED})
COMPLETION_SCOPES = frozenset(
    {"no_fields", "target_word_verified", "libreoffice_non_final", "incomplete"}
)
SELECTIVE_STATUSES = frozenset(
    {"selective_verified", "libreoffice_selective", "rejected", "error"}
)
FALLBACK_REASONS = frozenset(
    {"external_error", "libreoffice_error", "libreoffice_contract_or_integrity"}
)
FAILURE_STAGES = frozenset(
    {
        "external_field_workflow",
        "libreoffice_refresh",
        "selective_writeback",
        "post_writeback_integrity",
    }
)
FAILURE_CHECKS = frozenset(
    {
        "external_field_workflow",
        "libreoffice_refresh",
        "selective_writeback",
        "field_contract",
        "field_refresh",
        "content_integrity",
        "protected_object_integrity",
        "effective_font_integrity",
    }
)


class BackendEvidenceError(ValueError):
    """Backend evidence was not representable by the approved contract."""


def _strict_version(value: Any, expected: int) -> bool:
    return type(value) is int and value == expected


def _safe_path_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise BackendEvidenceError(f"{label} path is missing")
    if any(ord(character) < 32 or ord(character) == 127 for character in value):
        raise BackendEvidenceError(f"{label} path contains a control character")
    return value


def _safe_resolve(value: Any, label: str) -> Path:
    text = _safe_path_text(value, label)
    try:
        return Path(text).resolve(strict=False)
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise BackendEvidenceError(f"{label} path cannot be resolved") from exc


def _positive_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _nonnegative_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


def _optional_bool(value: dict[str, Any], name: str) -> bool | None:
    item = value.get(name)
    if item is not None and not isinstance(item, bool):
        raise BackendEvidenceError(f"backend {name} must be boolean or null")
    return item


def _optional_positive_int(value: dict[str, Any], name: str) -> int | None:
    item = value.get(name)
    if item is not None and not _positive_int(item):
        raise BackendEvidenceError(f"backend {name} must be a positive integer or null")
    return item


def _optional_nonnegative_int(value: dict[str, Any], name: str) -> int | None:
    item = value.get(name)
    if item is not None and not _nonnegative_int(item):
        raise BackendEvidenceError(
            f"backend {name} must be a non-negative integer or null"
        )
    return item


def _target_id(value: Any, label: str) -> str | None:
    if value is None:
        return None
    if value not in TARGET_IDS:
        raise BackendEvidenceError(f"{label} target ID is unsupported")
    return value


def _selective_projection(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        raise BackendEvidenceError("selective_writeback must be an object")
    status = value.get("status")
    if status not in SELECTIVE_STATUSES:
        raise BackendEvidenceError("selective_writeback status is unsupported")
    return {
        "version": CANONICAL_NESTED_VERSION,
        "status": status,
        "matched_fields": _optional_nonnegative_int(value, "matched_fields"),
        "updated_fields": _optional_nonnegative_int(value, "updated_fields"),
    }


def _verification_projection(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        raise BackendEvidenceError("read_only_verification must be an object")
    operation = value.get("operation")
    if operation is not None and operation != "verify_only":
        raise BackendEvidenceError("read_only_verification operation is unsupported")
    return {
        "version": CANONICAL_NESTED_VERSION,
        "operation": operation,
        "target_id": _target_id(value.get("target_id"), "verification"),
        "read_only_verified": _optional_bool(value, "read_only_verified"),
        "repaginated": _optional_bool(value, "repaginated"),
        "saved": _optional_bool(value, "saved"),
        "pdf_exported": _optional_bool(value, "pdf_exported"),
        "page_count": _optional_positive_int(value, "page_count"),
    }


def _failure_projection(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        raise BackendEvidenceError("attempt failure must be an object")
    status = value.get("status")
    stage = value.get("stage")
    failed_checks = value.get("failed_checks")
    if status != "rejected":
        raise BackendEvidenceError("attempt failure status is unsupported")
    if stage not in FAILURE_STAGES:
        raise BackendEvidenceError("attempt failure stage is unsupported")
    if not isinstance(failed_checks, list) or any(
        item not in FAILURE_CHECKS for item in failed_checks
    ):
        raise BackendEvidenceError("attempt failure checks are unsupported")
    return {
        "version": CANONICAL_NESTED_VERSION,
        "status": status,
        "stage": stage,
        "failed_checks": list(failed_checks),
    }


def _attempt_projection(raw: dict[str, Any]) -> dict[str, Any] | None:
    attempted = raw.get("attempted_backend")
    fallback = raw.get("fallback_from")
    if attempted is None and fallback is None:
        return None
    if not isinstance(attempted, dict):
        raise BackendEvidenceError("attempted_backend must be an object")
    backend = attempted.get("backend")
    if backend not in BACKENDS:
        raise BackendEvidenceError("attempted backend is unsupported")
    if fallback not in FALLBACK_REASONS:
        raise BackendEvidenceError("backend fallback reason is unsupported")
    selective = attempted.get("selective_writeback")
    selective_status = None
    if selective is not None:
        if not isinstance(selective, dict):
            raise BackendEvidenceError("attempted selective_writeback must be an object")
        selective_status = selective.get("status")
        if selective_status not in SELECTIVE_STATUSES:
            raise BackendEvidenceError(
                "attempted selective_writeback status is unsupported"
            )
    return {
        "version": CANONICAL_NESTED_VERSION,
        "backend": backend,
        "fallback_from": fallback,
        "selective_writeback_status": selective_status,
        "failure": _failure_projection(attempted.get("failure")),
    }


def canonical_backend_projection(raw: Any) -> dict[str, Any]:
    """Project raw backend diagnostics into the finite business-gate schema."""
    if not isinstance(raw, dict):
        raise BackendEvidenceError("backend diagnostics must be an object")
    backend = raw.get("backend")
    if backend not in BACKENDS:
        raise BackendEvidenceError("backend is unsupported")
    completion_scope = raw.get("completion_scope")
    if completion_scope is not None and completion_scope not in COMPLETION_SCOPES:
        raise BackendEvidenceError("backend completion_scope is unsupported")
    return {
        "version": CANONICAL_BACKEND_VERSION,
        "backend": backend,
        "target_id": _target_id(raw.get("target_id"), "backend"),
        "field_cache_verified": _optional_bool(raw, "field_cache_verified"),
        "page_count": _optional_positive_int(raw, "page_count"),
        "completion_scope": completion_scope,
        "delivery_field_contract_identical": _optional_bool(
            raw, "delivery_field_contract_identical"
        ),
        "selective_writeback": _selective_projection(
            raw.get("selective_writeback")
        ),
        "read_only_verification": _verification_projection(
            raw.get("read_only_verification")
        ),
        "attempt": _attempt_projection(raw),
    }


def _exact_keys(value: Any, expected: set[str], label: str) -> list[str]:
    if not isinstance(value, dict):
        return [f"{label} must be an object"]
    errors = []
    for name in sorted(expected - value.keys()):
        errors.append(f"{label} is missing {name}")
    for name in sorted(str(item) for item in value.keys() - expected):
        errors.append(f"{label} contains unsupported field {name}")
    return errors


def canonical_backend_shape_errors(value: Any) -> list[str]:
    keys = {
        "version",
        "backend",
        "target_id",
        "field_cache_verified",
        "page_count",
        "completion_scope",
        "delivery_field_contract_identical",
        "selective_writeback",
        "read_only_verification",
        "attempt",
    }
    errors = _exact_keys(value, keys, "field_backend")
    if not isinstance(value, dict):
        return errors
    if not _strict_version(value.get("version"), CANONICAL_BACKEND_VERSION):
        errors.append("field_backend version is unsupported")
    if value.get("backend") not in BACKENDS:
        errors.append("field_backend backend is unsupported")
    if (
        value.get("target_id") is not None
        and value.get("target_id") not in TARGET_IDS
    ):
        errors.append("field_backend target ID is unsupported")
    for name in ("field_cache_verified", "delivery_field_contract_identical"):
        if value.get(name) is not None and not isinstance(value.get(name), bool):
            errors.append(f"field_backend {name} must be boolean or null")
    if value.get("page_count") is not None and not _positive_int(
        value.get("page_count")
    ):
        errors.append("field_backend page_count must be a positive integer or null")
    if value.get("completion_scope") is not None and value.get(
        "completion_scope"
    ) not in COMPLETION_SCOPES:
        errors.append("field_backend completion_scope is unsupported")

    selective = value.get("selective_writeback")
    if selective is not None:
        selective_keys = {"version", "status", "matched_fields", "updated_fields"}
        errors.extend(
            _exact_keys(selective, selective_keys, "field_backend.selective_writeback")
        )
        if isinstance(selective, dict):
            if not _strict_version(
                selective.get("version"), CANONICAL_NESTED_VERSION
            ):
                errors.append(
                    "field_backend.selective_writeback version is unsupported"
                )
            if selective.get("status") not in SELECTIVE_STATUSES:
                errors.append("field_backend.selective_writeback status is unsupported")
            for name in ("matched_fields", "updated_fields"):
                if selective.get(name) is not None and not _nonnegative_int(
                    selective.get(name)
                ):
                    errors.append(
                        f"field_backend.selective_writeback {name} must be a non-negative integer or null"
                    )

    verification = value.get("read_only_verification")
    if verification is not None:
        verification_keys = {
            "version",
            "operation",
            "target_id",
            "read_only_verified",
            "repaginated",
            "saved",
            "pdf_exported",
            "page_count",
        }
        errors.extend(
            _exact_keys(
                verification,
                verification_keys,
                "field_backend.read_only_verification",
            )
        )
        if isinstance(verification, dict):
            if not _strict_version(
                verification.get("version"), CANONICAL_NESTED_VERSION
            ):
                errors.append(
                    "field_backend.read_only_verification version is unsupported"
                )
            if verification.get("operation") not in {None, "verify_only"}:
                errors.append(
                    "field_backend.read_only_verification operation is unsupported"
                )
            if verification.get("target_id") is not None and verification.get(
                "target_id"
            ) not in TARGET_IDS:
                errors.append(
                    "field_backend.read_only_verification target ID is unsupported"
                )
            for name in (
                "read_only_verified",
                "repaginated",
                "saved",
                "pdf_exported",
            ):
                if verification.get(name) is not None and not isinstance(
                    verification.get(name), bool
                ):
                    errors.append(
                        f"field_backend.read_only_verification {name} must be boolean or null"
                    )
            if verification.get("page_count") is not None and not _positive_int(
                verification.get("page_count")
            ):
                errors.append(
                    "field_backend.read_only_verification page_count must be a positive integer or null"
                )

    attempt = value.get("attempt")
    if attempt is not None:
        attempt_keys = {
            "version",
            "backend",
            "fallback_from",
            "selective_writeback_status",
            "failure",
        }
        errors.extend(_exact_keys(attempt, attempt_keys, "field_backend.attempt"))
        if isinstance(attempt, dict):
            if not _strict_version(
                attempt.get("version"), CANONICAL_NESTED_VERSION
            ):
                errors.append("field_backend.attempt version is unsupported")
            if attempt.get("backend") not in BACKENDS:
                errors.append("field_backend.attempt backend is unsupported")
            if attempt.get("fallback_from") not in FALLBACK_REASONS:
                errors.append("field_backend.attempt fallback is unsupported")
            if attempt.get("selective_writeback_status") is not None and attempt.get(
                "selective_writeback_status"
            ) not in SELECTIVE_STATUSES:
                errors.append(
                    "field_backend.attempt selective status is unsupported"
                )
            failure = attempt.get("failure")
            if failure is not None:
                failure_keys = {"version", "status", "stage", "failed_checks"}
                errors.extend(
                    _exact_keys(
                        failure, failure_keys, "field_backend.attempt.failure"
                    )
                )
                if isinstance(failure, dict):
                    if not _strict_version(
                        failure.get("version"), CANONICAL_NESTED_VERSION
                    ):
                        errors.append(
                            "field_backend.attempt.failure version is unsupported"
                        )
                    if failure.get("status") != "rejected":
                        errors.append(
                            "field_backend.attempt.failure status is unsupported"
                        )
                    if failure.get("stage") not in FAILURE_STAGES:
                        errors.append(
                            "field_backend.attempt.failure stage is unsupported"
                        )
                    failed_checks = failure.get("failed_checks")
                    if not isinstance(failed_checks, list) or any(
                        item not in FAILURE_CHECKS for item in failed_checks
                    ):
                        errors.append(
                            "field_backend.attempt.failure checks are unsupported"
                        )
    return list(dict.fromkeys(errors))


def _validate_json_value(value: Any, *, depth: int, counter: list[int]) -> None:
    if depth > BACKEND_AUDIT_MAX_DEPTH:
        raise BackendEvidenceError("backend audit exceeds maximum depth")
    counter[0] += 1
    if counter[0] > BACKEND_AUDIT_MAX_NODES:
        raise BackendEvidenceError("backend audit exceeds maximum node count")
    if value is None or isinstance(value, bool):
        return
    if isinstance(value, str):
        if len(value) > BACKEND_AUDIT_MAX_STRING_LENGTH:
            raise BackendEvidenceError("backend audit string exceeds maximum length")
        return
    if isinstance(value, int):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise BackendEvidenceError("backend audit contains a non-finite number")
        return
    if isinstance(value, list):
        for item in value:
            _validate_json_value(item, depth=depth + 1, counter=counter)
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise BackendEvidenceError("backend audit object key is not a string")
            if len(key) > BACKEND_AUDIT_MAX_STRING_LENGTH:
                raise BackendEvidenceError("backend audit key exceeds maximum length")
            _validate_json_value(item, depth=depth + 1, counter=counter)
        return
    raise BackendEvidenceError("backend audit contains a non-JSON value")


def backend_audit_bytes(raw_backend: Any) -> bytes:
    document = {"backend_audit_version": BACKEND_AUDIT_VERSION, "backend": raw_backend}
    _validate_json_value(document, depth=0, counter=[0])
    try:
        encoded = (
            json.dumps(
                document,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
                allow_nan=False,
            )
            + "\n"
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise BackendEvidenceError("backend audit is not standard JSON") from exc
    if len(encoded) > BACKEND_AUDIT_MAX_BYTES:
        raise BackendEvidenceError("backend audit exceeds maximum byte size")
    return encoded


def backend_audit_path(status_output: Path) -> Path:
    return status_output.with_name(f"{status_output.stem}-backend-audit.json")


def backend_audit_binding(path: Path | None, encoded: bytes) -> dict[str, Any]:
    if path is None:
        return {
            "version": BACKEND_AUDIT_BINDING_VERSION,
            "status": "not_persisted",
            "artifact": None,
        }
    return {
        "version": BACKEND_AUDIT_BINDING_VERSION,
        "status": "persisted",
        "artifact": {
            "path": str(path.resolve(strict=False)),
            "sha256": hashlib.sha256(encoded).hexdigest(),
            "size_bytes": len(encoded),
        },
    }


def atomic_write_bytes(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        temp.write_bytes(payload)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def atomic_write_json(path: Path, value: Any) -> None:
    try:
        payload = (
            json.dumps(
                value,
                ensure_ascii=False,
                indent=2,
                sort_keys=True,
                allow_nan=False,
            )
            + "\n"
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise BackendEvidenceError("finalization evidence is not standard JSON") from exc
    atomic_write_bytes(path, payload)


def read_bound_backend_audit(
    finalization: dict[str, Any], *, expected_path: Path | None = None
) -> dict[str, Any]:
    binding = finalization.get("backend_audit")
    if not isinstance(binding, dict) or set(binding) != {
        "version",
        "status",
        "artifact",
    }:
        raise BackendEvidenceError("backend audit binding schema is invalid")
    if not _strict_version(
        binding.get("version"), BACKEND_AUDIT_BINDING_VERSION
    ):
        raise BackendEvidenceError("backend audit binding version is unsupported")
    if binding.get("status") != "persisted":
        raise BackendEvidenceError("backend audit is not persisted")
    artifact = binding.get("artifact")
    if not isinstance(artifact, dict) or set(artifact) != {
        "path",
        "sha256",
        "size_bytes",
    }:
        raise BackendEvidenceError("backend audit identity is incomplete")
    sha256 = artifact.get("sha256")
    size_bytes = artifact.get("size_bytes")
    if not (
        isinstance(sha256, str)
        and len(sha256) == 64
        and all(character in "0123456789abcdef" for character in sha256)
    ):
        raise BackendEvidenceError("backend audit SHA-256 identity is invalid")
    if type(size_bytes) is not int or size_bytes < 1:
        raise BackendEvidenceError("backend audit size identity is invalid")
    if size_bytes > BACKEND_AUDIT_MAX_BYTES:
        raise BackendEvidenceError("backend audit exceeds maximum byte size")
    path_text = _safe_path_text(artifact.get("path"), "backend audit")
    lexical_path = Path(path_text)
    try:
        if lexical_path.is_symlink():
            raise BackendEvidenceError("backend audit artifact must not be a symlink")
    except BackendEvidenceError:
        raise
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise BackendEvidenceError("backend audit artifact cannot be inspected") from exc
    path = _safe_resolve(path_text, "backend audit")
    if expected_path is not None:
        try:
            if expected_path.is_symlink():
                raise BackendEvidenceError(
                    "expected backend audit output must not be a symlink"
                )
            resolved_expected = expected_path.resolve(strict=False)
        except BackendEvidenceError:
            raise
        except (OSError, RuntimeError, TypeError, ValueError) as exc:
            raise BackendEvidenceError(
                "expected backend audit path cannot be resolved"
            ) from exc
        if path != resolved_expected:
            raise BackendEvidenceError("backend audit path differs from expected output")
    try:
        details = path.lstat()
        mode = details.st_mode
        if not stat.S_ISREG(mode):
            raise BackendEvidenceError("backend audit artifact is not a regular file")
        if details.st_size != size_bytes:
            raise BackendEvidenceError("backend audit size differs from binding")
        payload = path.read_bytes()
    except BackendEvidenceError:
        raise
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise BackendEvidenceError("backend audit artifact is missing") from exc
    if len(payload) != size_bytes:
        raise BackendEvidenceError("backend audit size differs from binding")
    if hashlib.sha256(payload).hexdigest() != sha256:
        raise BackendEvidenceError("backend audit SHA-256 differs from binding")
    if len(payload) > BACKEND_AUDIT_MAX_BYTES:
        raise BackendEvidenceError("backend audit exceeds maximum byte size")
    try:
        def reject_constant(value: str) -> None:
            raise ValueError(value)

        def strict_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
            value: dict[str, Any] = {}
            for key, item in pairs:
                if key in value:
                    raise ValueError(f"duplicate key: {key}")
                value[key] = item
            return value

        document = json.loads(
            payload,
            parse_constant=reject_constant,
            object_pairs_hook=strict_object,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
        raise BackendEvidenceError("backend audit is invalid standard JSON") from exc
    _validate_json_value(document, depth=0, counter=[0])
    if not isinstance(document, dict) or set(document) != {
        "backend_audit_version",
        "backend",
    }:
        raise BackendEvidenceError("backend audit root schema is invalid")
    if not _strict_version(
        document.get("backend_audit_version"), BACKEND_AUDIT_VERSION
    ):
        raise BackendEvidenceError("backend audit version is unsupported")
    if not isinstance(document.get("backend"), dict):
        raise BackendEvidenceError("backend audit payload is not an object")
    return document["backend"]
