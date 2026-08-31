#!/usr/bin/env python3
"""Canonical final-ready field evidence shapes and consistency validation."""

from __future__ import annotations

from typing import Any

from backend_evidence import (
    BACKEND_AUDIT_BINDING_VERSION,
    canonical_backend_shape_errors,
)
from target_software import LIBREOFFICE, MICROSOFT_WORD, UNSUPPORTED


FINALIZATION_EVIDENCE_VERSION = 1
ARTIFACT_BINDING_VERSION = 1
FIELD_CACHE_STATUSES = frozenset({"absent", "code_only", "stale", "refreshed"})
DELIVERY_FIELD_STATUSES = frozenset(
    FIELD_CACHE_STATUSES
    | {"deferred", "selective_verified", "libreoffice_refreshed"}
)
FIELD_BACKENDS = frozenset(
    {"not_needed", "external", "deferred_on_open", "libreoffice_uno"}
)
FIELD_WRITEBACK_STATUSES = frozenset(
    {"not_needed", "deferred", "selective_verified", "libreoffice_selective"}
)
COMPLETION_SCOPES = frozenset(
    {"no_fields", "target_word_verified", "libreoffice_non_final", "incomplete"}
)
TARGET_SOFTWARE_IDS = frozenset({MICROSOFT_WORD, LIBREOFFICE, UNSUPPORTED})
PUBLICATION_RECORD_VERSION = 1


def _valid_sha256(value: Any) -> bool:
    return bool(
        isinstance(value, str)
        and len(value) == 64
        and all(character in "0123456789abcdef" for character in value)
    )


def _valid_nonempty_path(value: Any) -> bool:
    return bool(
        isinstance(value, str)
        and value.strip()
        and not any(
            ord(character) < 32 or ord(character) == 127 for character in value
        )
    )


def _strict_version(value: Any, expected: int) -> bool:
    return type(value) is int and value == expected


def _valid_nonnegative_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


def _valid_positive_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _missing_keys(value: dict[str, Any], required: set[str]) -> list[str]:
    return sorted(required - value.keys())


def _unexpected_keys(value: dict[str, Any], allowed: set[str]) -> list[str]:
    return sorted(str(name) for name in value.keys() - allowed)


def _field_cache_shape_errors(label: str, value: Any) -> list[str]:
    if not isinstance(value, dict):
        return [f"{label} must be an object"]
    required = {
        "status",
        "main_toc_fields",
        "toc_entries",
        "dirty_fields",
        "update_on_open",
        "field_types",
    }
    errors = [f"{label} is missing {name}" for name in _missing_keys(value, required)]
    errors.extend(
        f"{label} contains unsupported field {name}"
        for name in _unexpected_keys(value, required)
    )
    if value.get("status") not in FIELD_CACHE_STATUSES:
        errors.append(f"{label}.status is unsupported")
    for name in ("main_toc_fields", "toc_entries", "dirty_fields"):
        if not _valid_nonnegative_int(value.get(name)):
            errors.append(f"{label}.{name} must be a non-negative integer")
    if not isinstance(value.get("update_on_open"), bool):
        errors.append(f"{label}.update_on_open must be a boolean")
    field_types = value.get("field_types")
    if not isinstance(field_types, dict):
        errors.append(f"{label}.field_types must be an object")
    else:
        for name, count in field_types.items():
            if not isinstance(name, str) or not name.strip():
                errors.append(f"{label}.field_types contains an invalid field name")
            if not _valid_nonnegative_int(count):
                errors.append(
                    f"{label}.field_types[{name!r}] must be a non-negative integer"
                )
    return errors


def artifact_identity_shape_errors(
    label: str, value: Any, *, allow_page_count: bool
) -> list[str]:
    if not isinstance(value, dict):
        return [f"{label} must be an object"]
    errors = []
    required = {"path", "sha256", "size_bytes"}
    allowed = required | ({"page_count"} if allow_page_count else set())
    for name in required:
        if name not in value:
            errors.append(f"{label} is missing {name}")
    errors.extend(
        f"{label} contains unsupported field {name}"
        for name in _unexpected_keys(value, allowed)
    )
    if not _valid_nonempty_path(value.get("path")):
        errors.append(f"{label}.path must be a non-empty string")
    if not _valid_sha256(value.get("sha256")):
        errors.append(f"{label}.sha256 must be a lowercase SHA-256")
    if not _valid_positive_int(value.get("size_bytes")):
        errors.append(f"{label}.size_bytes must be a positive integer")
    if "page_count" in value:
        if not allow_page_count:
            errors.append(f"{label}.page_count is not allowed")
        elif not _valid_positive_int(value.get("page_count")):
            errors.append(f"{label}.page_count must be a positive integer")
    return errors


def _backend_shape_errors(value: Any) -> list[str]:
    return canonical_backend_shape_errors(value)


def finalization_evidence_shape_errors(value: Any) -> list[str]:
    """Validate the versioned producer schema without granting final readiness."""
    if not isinstance(value, dict):
        return ["finalization evidence must be an object"]
    required = {
        "finalization_evidence_version",
        "status",
        "delivery_field_status",
        "input_field_cache",
        "output_field_cache",
        "field_backend",
        "backend_audit",
        "artifact_binding",
        "field_writeback_status",
        "field_completion",
        "content_integrity",
        "protected_object_integrity",
        "effective_font_integrity",
        "workflow_state",
        "target_pdf",
        "target_layout_status",
        "target_software",
        "output",
    }
    errors = [f"finalization evidence is missing {name}" for name in _missing_keys(value, required)]
    allowed = required | {"publication"}
    errors.extend(
        f"finalization evidence contains unsupported field {name}"
        for name in _unexpected_keys(value, allowed)
    )
    publication = value.get("publication")
    if publication is not None:
        publication_keys = {
            "version",
            "transaction_id",
            "retained_staging_directory",
            "cleanup_policy",
            "business_gate",
        }
        if not isinstance(publication, dict):
            errors.append("publication must be an object")
        else:
            errors.extend(
                f"publication is missing {name}"
                for name in _missing_keys(publication, publication_keys)
            )
            errors.extend(
                f"publication contains unsupported field {name}"
                for name in _unexpected_keys(publication, publication_keys)
            )
            if not _strict_version(
                publication.get("version"), PUBLICATION_RECORD_VERSION
            ):
                errors.append("publication.version is unsupported")
            transaction_id = publication.get("transaction_id")
            if not (
                isinstance(transaction_id, str)
                and len(transaction_id) == 32
                and all(character in "0123456789abcdef" for character in transaction_id)
            ):
                errors.append("publication.transaction_id is invalid")
            if not _valid_nonempty_path(
                publication.get("retained_staging_directory")
            ):
                errors.append(
                    "publication.retained_staging_directory must be a non-empty path"
                )
            if publication.get("cleanup_policy") != "manual_only":
                errors.append("publication.cleanup_policy is unsupported")
            if publication.get("business_gate") is not False:
                errors.append("publication.business_gate must be false")
    version = value.get("finalization_evidence_version")
    if not _strict_version(version, FINALIZATION_EVIDENCE_VERSION):
        errors.append("finalization_evidence_version is unsupported")
    if value.get("status") != "pass":
        errors.append("finalization status is unsupported")
    if value.get("delivery_field_status") not in DELIVERY_FIELD_STATUSES:
        errors.append("delivery_field_status is unsupported")
    errors.extend(
        _field_cache_shape_errors("input_field_cache", value.get("input_field_cache"))
    )
    errors.extend(
        _field_cache_shape_errors(
            "output_field_cache", value.get("output_field_cache")
        )
    )
    errors.extend(_backend_shape_errors(value.get("field_backend")))
    backend_audit = value.get("backend_audit")
    if not isinstance(backend_audit, dict):
        errors.append("backend_audit must be an object")
    else:
        backend_audit_keys = {"version", "status", "artifact"}
        errors.extend(
            f"backend_audit contains unsupported field {name}"
            for name in _unexpected_keys(backend_audit, backend_audit_keys)
        )
        errors.extend(
            f"backend_audit is missing {name}"
            for name in _missing_keys(backend_audit, backend_audit_keys)
        )
        if not _strict_version(
            backend_audit.get("version"), BACKEND_AUDIT_BINDING_VERSION
        ):
            errors.append("backend_audit.version is unsupported")
        audit_status = backend_audit.get("status")
        if audit_status not in {"persisted", "not_persisted"}:
            errors.append("backend_audit.status is unsupported")
        audit_artifact = backend_audit.get("artifact")
        if audit_status == "persisted":
            errors.extend(
                artifact_identity_shape_errors(
                    "backend_audit.artifact",
                    audit_artifact,
                    allow_page_count=False,
                )
            )
        elif audit_artifact is not None:
            errors.append("non-persisted backend_audit unexpectedly has an artifact")
    if value.get("field_writeback_status") not in FIELD_WRITEBACK_STATUSES:
        errors.append("field_writeback_status is unsupported")

    binding = value.get("artifact_binding")
    if not isinstance(binding, dict):
        errors.append("artifact_binding must be an object")
    else:
        binding_allowed = {"version", "finalized_docx", "word_verification_pdf"}
        errors.extend(
            f"artifact_binding contains unsupported field {name}"
            for name in _unexpected_keys(binding, binding_allowed)
        )
        if not _strict_version(
            binding.get("version"), ARTIFACT_BINDING_VERSION
        ):
            errors.append("artifact_binding.version is unsupported")
        if "finalized_docx" not in binding:
            errors.append("artifact_binding is missing finalized_docx")
        errors.extend(
            artifact_identity_shape_errors(
                "artifact_binding.finalized_docx",
                binding.get("finalized_docx"),
                allow_page_count=False,
            )
        )
        if "word_verification_pdf" not in binding:
            errors.append("artifact_binding is missing word_verification_pdf")
        elif binding.get("word_verification_pdf") is not None:
            errors.extend(
                artifact_identity_shape_errors(
                    "artifact_binding.word_verification_pdf",
                    binding.get("word_verification_pdf"),
                    allow_page_count=True,
                )
            )

    completion = value.get("field_completion")
    completion_required = {
        "field_gate_completed",
        "final_ready_eligible",
        "word_verification_required",
        "word_verification_completed",
        "completion_scope",
        "evidence_validation",
    }
    if not isinstance(completion, dict):
        errors.append("field_completion must be an object")
    else:
        errors.extend(
            f"field_completion is missing {name}"
            for name in _missing_keys(completion, completion_required)
        )
        errors.extend(
            f"field_completion contains unsupported field {name}"
            for name in _unexpected_keys(completion, completion_required)
        )
        for name in (
            "field_gate_completed",
            "final_ready_eligible",
            "word_verification_required",
            "word_verification_completed",
        ):
            if not isinstance(completion.get(name), bool):
                errors.append(f"field_completion.{name} must be a boolean")
        if completion.get("completion_scope") not in COMPLETION_SCOPES:
            errors.append("field_completion.completion_scope is unsupported")
        validation = completion.get("evidence_validation")
        if not isinstance(validation, dict):
            errors.append("field_completion.evidence_validation must be an object")
        else:
            if set(validation) != {"status", "errors"}:
                errors.append(
                    "field_completion.evidence_validation must contain status and errors"
                )
            validation_status = validation.get("status")
            validation_errors = validation.get("errors")
            if validation_status not in {"pass", "incomplete"}:
                errors.append(
                    "field_completion.evidence_validation.status is unsupported"
                )
            if not isinstance(validation_errors, list) or any(
                not isinstance(item, str) or not item
                for item in validation_errors or []
            ):
                errors.append(
                    "field_completion.evidence_validation.errors must be a list of non-empty strings"
                )
            elif (validation_status == "pass") != (validation_errors == []):
                errors.append(
                    "field_completion.evidence_validation status and errors disagree"
                )

    for name in (
        "content_integrity",
        "protected_object_integrity",
        "effective_font_integrity",
    ):
        if value.get(name) != "pass":
            errors.append(f"{name} is unsupported")
    workflow = value.get("workflow_state")
    workflow_hashes = {
        "source_sha256",
        "input_sha256",
        "profile_sha256",
        "structure_map_sha256",
        "output_sha256",
    }
    if not isinstance(workflow, dict):
        errors.append("workflow_state must be an object")
    else:
        workflow_allowed = workflow_hashes | {"stage"}
        for name in _missing_keys(workflow, workflow_allowed):
            errors.append(f"workflow_state is missing {name}")
        errors.extend(
            f"workflow_state contains unsupported field {name}"
            for name in _unexpected_keys(workflow, workflow_allowed)
        )
        if workflow.get("stage") != "finalized":
            errors.append("workflow_state.stage is unsupported")
        for name in workflow_hashes:
            if not _valid_sha256(workflow.get(name)):
                errors.append(f"workflow_state.{name} must be a lowercase SHA-256")
    if value.get("target_software") not in TARGET_SOFTWARE_IDS:
        errors.append("target_software is unsupported")
    target_pdf = value.get("target_pdf")
    if target_pdf is not None and not _valid_nonempty_path(target_pdf):
        errors.append("target_pdf must be null or a non-empty path")
    if value.get("target_layout_status") not in {
        "not_verified",
        "target_pdf_ready_for_visual_qa",
    }:
        errors.append("target_layout_status is unsupported")
    if not _valid_nonempty_path(value.get("output")):
        errors.append("output must be a non-empty path")
    return list(dict.fromkeys(errors))


def completion_evidence(finalization: dict[str, Any]) -> dict[str, Any]:
    """Extract every field-gate fact needed for later revalidation."""
    backend = finalization.get("field_backend", {})
    selective = backend.get("selective_writeback") or {}
    verification = backend.get("read_only_verification") or {}
    completion = finalization.get("field_completion", {})
    return {
        "delivery_status": finalization.get("delivery_field_status"),
        "input_cache_status": finalization.get("input_field_cache", {}).get(
            "status"
        ),
        "output_cache_status": finalization.get("output_field_cache", {}).get(
            "status"
        ),
        "backend": backend.get("backend"),
        "field_cache_verified": backend.get("field_cache_verified"),
        "calculation_page_count": backend.get("page_count"),
        "writeback_status": finalization.get("field_writeback_status"),
        "selective_writeback_status": selective.get("status"),
        "read_only_operation": verification.get("operation"),
        "read_only_verified": verification.get("read_only_verified"),
        "read_only_repaginated": verification.get("repaginated"),
        "read_only_saved": verification.get("saved"),
        "verification_pdf_exported": verification.get("pdf_exported"),
        "verification_page_count": verification.get("page_count"),
        "artifact_binding": finalization.get("artifact_binding"),
        "word_verification_required": completion.get(
            "word_verification_required"
        ),
        "word_verification_completed": completion.get(
            "word_verification_completed"
        ),
        "completion_scope": completion.get("completion_scope"),
        "field_gate_completed": completion.get("field_gate_completed"),
        "final_ready_eligible": completion.get("final_ready_eligible"),
    }


def final_ready_evidence_errors(evidence: dict[str, Any]) -> list[str]:
    """Accept only the complete no-fields or target-Word evidence shape."""
    no_fields = evidence.get("delivery_status") == "absent"
    word_verified = evidence.get("delivery_status") == "selective_verified"
    if not no_fields and not word_verified:
        return ["delivery_status is neither absent nor selective_verified"]

    expected = (
        {
            "input_cache_status": "absent",
            "output_cache_status": "absent",
            "backend": "not_needed",
            "writeback_status": "not_needed",
            "selective_writeback_status": None,
            "field_cache_verified": None,
            "calculation_page_count": None,
            "read_only_operation": None,
            "read_only_verified": None,
            "read_only_repaginated": None,
            "read_only_saved": None,
            "verification_pdf_exported": None,
            "verification_page_count": None,
            "word_verification_required": False,
            "word_verification_completed": False,
            "completion_scope": "no_fields",
            "field_gate_completed": True,
            "final_ready_eligible": True,
        }
        if no_fields
        else {
            "input_cache_status": "stale",
            "output_cache_status": "refreshed",
            "backend": "external",
            "field_cache_verified": True,
            "writeback_status": "selective_verified",
            "selective_writeback_status": "selective_verified",
            "read_only_operation": "verify_only",
            "read_only_verified": True,
            "read_only_repaginated": True,
            "read_only_saved": False,
            "verification_pdf_exported": True,
            "word_verification_required": True,
            "word_verification_completed": True,
            "completion_scope": "target_word_verified",
            "field_gate_completed": True,
            "final_ready_eligible": True,
        }
    )
    errors = [
        f"{name}={evidence.get(name)!r}, expected {value!r}"
        for name, value in expected.items()
        if evidence.get(name) != value
    ]
    if word_verified:
        calculation_pages = evidence.get("calculation_page_count")
        verification_pages = evidence.get("verification_page_count")
        if (
            not isinstance(calculation_pages, int)
            or isinstance(calculation_pages, bool)
            or calculation_pages < 1
        ):
            errors.append("calculation_page_count is not a positive integer")
        if (
            not isinstance(verification_pages, int)
            or isinstance(verification_pages, bool)
            or verification_pages < 1
        ):
            errors.append("verification_page_count is not a positive integer")
        if calculation_pages != verification_pages:
            errors.append("verification_page_count does not match calculation_page_count")
    binding = evidence.get("artifact_binding")
    if not isinstance(binding, dict) or not _strict_version(
        binding.get("version"), ARTIFACT_BINDING_VERSION
    ):
        errors.append("artifact_binding is missing or unsupported")
        binding = {}
    finalized = binding.get("finalized_docx")
    if not _valid_artifact_identity(finalized):
        errors.append("finalized_docx artifact identity is incomplete")
    verification_pdf = binding.get("word_verification_pdf")
    if word_verified:
        if not _valid_artifact_identity(verification_pdf):
            errors.append("word_verification_pdf artifact identity is incomplete")
        elif verification_pdf.get("page_count") != evidence.get(
            "verification_page_count"
        ):
            errors.append(
                "word_verification_pdf page_count does not match verification_page_count"
            )
    elif verification_pdf is not None:
        errors.append("no-fields evidence unexpectedly contains a Word verification PDF")
    return errors


def _valid_artifact_identity(value: Any) -> bool:
    return not artifact_identity_shape_errors(
        "artifact", value, allow_page_count=True
    )


def final_ready_evidence_valid(evidence: dict[str, Any]) -> bool:
    return not final_ready_evidence_errors(evidence)
