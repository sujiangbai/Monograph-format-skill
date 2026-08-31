#!/usr/bin/env python3
"""External-updater parsing and explicit-entity audit identities."""

from __future__ import annotations

import ast
import hashlib
import json
import os
import re
import shlex
import shutil
from pathlib import Path
from typing import Any


COMMAND_IDENTITY_VERSION = 3
DEPENDENCY_CLOSURE_VERSION = 2
PYTHON_EXECUTABLE = re.compile(r"^python(?:\d+(?:\.\d+)*)?(?:\.exe)?$", re.I)
PYTHON_MODULE = re.compile(r"^[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*$")
PATH_SUFFIXES = frozenset(
    {
        ".cfg",
        ".conf",
        ".ini",
        ".json",
        ".py",
        ".rsp",
        ".toml",
        ".txt",
        ".yaml",
        ".yml",
    }
)
PYTHON_OPTIONS_WITH_VALUE = frozenset({"-W", "-X", "--check-hash-based-pycs"})
UNBOUNDED_PYTHON_CALLS = frozenset(
    {
        "__import__",
        "compile",
        "eval",
        "exec",
        "importlib.import_module",
        "open",
        "runpy.run_module",
        "runpy.run_path",
    }
)


class ExternalCommandError(ValueError):
    """The configured external command cannot be parsed."""


def parse_external_command(value: str) -> list[str]:
    stripped = value.strip()
    if stripped.startswith("["):
        try:
            parsed = json.loads(stripped)
        except json.JSONDecodeError as exc:
            raise ExternalCommandError(
                "--field-updater-command JSON is invalid."
            ) from exc
        if not isinstance(parsed, list) or not parsed or not all(
            isinstance(item, str) and item for item in parsed
        ):
            raise ExternalCommandError(
                "--field-updater-command JSON must be a non-empty string array."
            )
        return parsed
    try:
        parts = shlex.split(stripped, posix=os.name != "nt")
    except ValueError as exc:
        raise ExternalCommandError(
            "--field-updater-command shell-like value is invalid."
        ) from exc
    if not parts:
        raise ExternalCommandError("--field-updater-command cannot be empty.")
    return parts


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _absolute_path(value: str, base_dir: Path) -> Path:
    candidate = Path(value).expanduser()
    if not candidate.is_absolute():
        candidate = base_dir / candidate
    return Path(os.path.abspath(candidate))


def _symlink_metadata(path: Path) -> dict[str, Any]:
    return {
        "is_symlink": path.is_symlink(),
        "link_target": os.readlink(path) if path.is_symlink() else None,
    }


def _file_identity(requested: str, path: Path) -> dict[str, Any]:
    lexical = Path(os.path.abspath(path))
    resolved = lexical.resolve(strict=True)
    return {
        "requested": requested,
        "lexical_path": str(lexical),
        "resolved_path": str(resolved),
        "type": "file",
        "sha256": _sha256(resolved),
        "size_bytes": resolved.stat().st_size,
        **_symlink_metadata(lexical),
    }


def _resolve_executable(value: str, base_dir: Path) -> tuple[Path, Path] | None:
    candidate = Path(value).expanduser()
    if candidate.is_absolute() or len(candidate.parts) > 1:
        lexical = _absolute_path(value, base_dir)
    else:
        search_path = os.environ.get("PATH", os.defpath)
        normalized_path = os.pathsep.join(
            str(_absolute_path(entry or ".", base_dir))
            for entry in search_path.split(os.pathsep)
        )
        located = shutil.which(value, path=normalized_path)
        if not located:
            return None
        lexical = Path(os.path.abspath(located))
    try:
        resolved = lexical.resolve(strict=True)
    except (OSError, RuntimeError):
        return None
    if not resolved.is_file() or not os.access(resolved, os.X_OK):
        return None
    return lexical, resolved


def _executable_identity(requested: str, lexical: Path) -> dict[str, Any]:
    item = _file_identity(requested, lexical)
    item["located_path"] = str(lexical)
    return item


def _looks_path_like(value: str) -> bool:
    if not value or "://" in value:
        return False
    candidate = Path(value).expanduser()
    return bool(
        candidate.is_absolute()
        or value.startswith((".", "~"))
        or "/" in value
        or "\\" in value
        or candidate.suffix.lower() in PATH_SUFFIXES
    )


def _directory_identity(
    requested: str, path: Path
) -> tuple[dict[str, Any], list[str]]:
    lexical = Path(os.path.abspath(path))
    reasons: list[str] = []
    root: dict[str, Any] = {
        "requested": requested,
        "lexical_path": str(lexical),
        "resolved_path": None,
        "type": "directory",
        "entries": [],
        **_symlink_metadata(lexical),
    }
    if lexical.is_symlink():
        try:
            root["resolved_path"] = str(lexical.resolve(strict=True))
        except (OSError, RuntimeError):
            reasons.append("directory_symlink_unresolved")
        reasons.append("directory_symlink_not_followed")
        return root, reasons
    try:
        resolved = lexical.resolve(strict=True)
    except (OSError, RuntimeError):
        reasons.append("directory_unresolved")
        return root, reasons
    root["resolved_path"] = str(resolved)
    entries: list[dict[str, Any]] = [{"path": ".", "type": "directory"}]
    try:
        pending = [(resolved, Path("."))]
        while pending:
            current, relative_root = pending.pop()
            with os.scandir(current) as iterator:
                children = sorted(iterator, key=lambda item: item.name)
            directories: list[tuple[Path, Path]] = []
            for child_entry in children:
                child = Path(child_entry.path)
                relative_path = relative_root / child_entry.name
                relative = relative_path.as_posix().removeprefix("./")
                if child_entry.is_symlink():
                    entries.append(
                        {
                            "path": relative,
                            "type": "symlink",
                            "link_target": os.readlink(child),
                        }
                    )
                    reasons.append(f"directory_symlink_not_followed:{relative}")
                elif child_entry.is_dir(follow_symlinks=False):
                    entries.append({"path": relative, "type": "directory"})
                    directories.append((child, relative_path))
                elif child_entry.is_file(follow_symlinks=False):
                    entries.append(
                        {
                            "path": relative,
                            "type": "file",
                            "sha256": _sha256(child),
                            "size_bytes": child.stat().st_size,
                        }
                    )
                else:
                    entries.append({"path": relative, "type": "other"})
                    reasons.append(f"directory_special_entry:{relative}")
            pending.extend(reversed(directories))
        entries[1:] = sorted(entries[1:], key=lambda item: item["path"])
    except OSError as exc:
        reasons.append(f"directory_scan_failed:{type(exc).__name__}")
    root["entries"] = entries
    return root, reasons


def _call_name(node: ast.Call) -> str | None:
    value: ast.AST = node.func
    parts: list[str] = []
    while isinstance(value, ast.Attribute):
        parts.append(value.attr)
        value = value.value
    if isinstance(value, ast.Name):
        parts.append(value.id)
        return ".".join(reversed(parts))
    return None


def _python_source_reasons(path: Path, allowed_package: str | None) -> list[str]:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, SyntaxError, UnicodeError) as exc:
        return [f"python_source_not_statically_readable:{path.name}:{type(exc).__name__}"]
    reasons: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                root = alias.name.split(".", 1)[0]
                if allowed_package is None or root != allowed_package:
                    reasons.append(f"python_external_import:{alias.name}")
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                if allowed_package is None:
                    reasons.append("python_relative_import_outside_bound_package")
            else:
                module = node.module or ""
                root = module.split(".", 1)[0]
                if allowed_package is None or root != allowed_package:
                    reasons.append(f"python_external_import:{module or '<unknown>'}")
        elif isinstance(node, ast.Call) and _call_name(node) in UNBOUNDED_PYTHON_CALLS:
            reasons.append(f"python_unbounded_call:{_call_name(node)}")
    return reasons


def _bind_path(
    value: str,
    base_dir: Path,
    *,
    source: str,
    ordinal: int,
) -> tuple[dict[str, Any] | None, list[str]]:
    lexical = _absolute_path(value, base_dir)
    if lexical.is_symlink() and not lexical.exists():
        return (
            {
                "requested": value,
                "lexical_path": str(lexical),
                "resolved_path": None,
                "type": "symlink",
                "source": source,
                "ordinal": ordinal,
                **_symlink_metadata(lexical),
            },
            [f"broken_symlink:{ordinal}"],
        )
    if lexical.is_file():
        item = _file_identity(value, lexical)
        item.update({"source": source, "ordinal": ordinal})
        return item, []
    if lexical.is_dir():
        item, reasons = _directory_identity(value, lexical)
        item.update({"source": source, "ordinal": ordinal})
        return item, reasons
    return None, [f"unresolved_path:{ordinal}:{value}"]


def _python_entry(
    argv: list[str],
    executable: Path,
    base_dir: Path,
) -> tuple[dict[str, Any] | None, set[int], list[dict[str, Any]], list[str]]:
    if not PYTHON_EXECUTABLE.fullmatch(executable.name):
        return None, set(), [], []
    index = 1
    while index < len(argv):
        argument = argv[index]
        if argument == "-m":
            if index + 1 >= len(argv):
                return {"kind": "module", "name": None}, {index}, [], ["python_module_missing"]
            module_name = argv[index + 1]
            entry = {"kind": "module", "name": module_name, "resolution": None}
            consumed = {index, index + 1}
            if not PYTHON_MODULE.fullmatch(module_name):
                return entry, consumed, [], ["python_module_name_invalid"]
            module_path = Path(*module_name.split("."))
            file_candidate = base_dir / module_path.with_suffix(".py")
            package_candidate = base_dir / module_path
            has_file = file_candidate.is_file()
            has_package = (package_candidate / "__init__.py").is_file() and (
                package_candidate / "__main__.py"
            ).is_file()
            if has_file == has_package:
                reason = "python_module_ambiguous" if has_file else "python_module_not_locally_resolved"
                return entry, consumed, [], [reason]
            if has_file:
                item = _file_identity(module_name, file_candidate)
                item.update({"source": "python_module", "ordinal": index + 1})
                entry["resolution"] = "local_module_file"
                reasons = _python_source_reasons(file_candidate, None)
                return entry, consumed, [item], reasons
            top_package = module_name.split(".", 1)[0]
            package_root = base_dir / top_package
            item, reasons = _directory_identity(top_package, package_root)
            item.update({"source": "python_module_package", "ordinal": index + 1})
            entry["resolution"] = "local_package_tree"
            for package_entry in item.get("entries", []):
                if (
                    package_entry.get("type") == "file"
                    and str(package_entry.get("path", "")).endswith(".py")
                ):
                    reasons.extend(
                        _python_source_reasons(
                            package_root / package_entry["path"], top_package
                        )
                    )
            return entry, consumed, [item], reasons
        if argument in {"-c", "-"}:
            return {"kind": "dynamic", "argument": argument}, {index}, [], [
                "python_dynamic_entry"
            ]
        if argument in PYTHON_OPTIONS_WITH_VALUE:
            index += 2
            continue
        if argument.startswith("-"):
            index += 1
            continue
        item, reasons = _bind_path(
            argument, base_dir, source="python_script", ordinal=index
        )
        entry = {"kind": "script", "argument": argument, "resolution": None}
        if item is None or item.get("type") != "file":
            return entry, {index}, [item] if item else [], reasons or [
                "python_script_not_resolved"
            ]
        entry["resolution"] = "local_script_file"
        source_path = Path(str(item["resolved_path"]))
        reasons.extend(_python_source_reasons(source_path, None))
        return entry, {index}, [item], reasons
    return {"kind": "stdin", "argument": None}, set(), [], [
        "python_stdin_entry_not_cacheable"
    ]


def _dependency_closure(
    argv: list[str], executable: Path, base_dir: Path
) -> dict[str, Any]:
    entities: list[dict[str, Any]] = []
    reasons: list[str] = []
    python_entry, consumed, python_entities, python_reasons = _python_entry(
        argv, executable, base_dir
    )
    entities.extend(item for item in python_entities if item is not None)
    reasons.extend(python_reasons)
    index = 1
    while index < len(argv):
        if index in consumed:
            index += 1
            continue
        argument = argv[index]
        candidate_value: str | None = None
        source = "direct_argument"
        if argument.startswith("@") and len(argument) > 1:
            candidate_value = argument[1:]
            source = "response_file"
            reasons.append("response_file_contents_not_expanded")
        elif argument.startswith("--") and "=" in argument:
            _, candidate_value = argument.split("=", 1)
            source = "option_assignment"
        elif (
            argument.startswith("--")
            and index + 1 < len(argv)
            and index + 1 not in consumed
            and not argv[index + 1].startswith("-")
        ):
            value = argv[index + 1]
            lexical = _absolute_path(value, base_dir)
            if lexical.exists() or lexical.is_symlink() or _looks_path_like(value):
                candidate_value = value
                source = "option_value"
                index += 1
        elif not argument.startswith("-"):
            candidate_value = argument
        if candidate_value:
            lexical = _absolute_path(candidate_value, base_dir)
            if lexical.exists() or lexical.is_symlink() or _looks_path_like(candidate_value):
                item, item_reasons = _bind_path(
                    candidate_value, base_dir, source=source, ordinal=index
                )
                if item is not None:
                    entities.append(item)
                reasons.extend(item_reasons)
        index += 1
    reasons = sorted(
        {
            "external_program_not_hermetic",
            "runtime_dependencies_unproven",
            *reasons,
        }
    )
    return {
        "version": DEPENDENCY_CLOSURE_VERSION,
        "status": "unproven",
        "cache_reusable": False,
        "execution_cwd": str(base_dir.resolve()),
        "symlink_policy": (
            "explicit file symlinks bind lexical and resolved targets; "
            "directory symlinks are recorded but never followed"
        ),
        "python_entry": python_entry,
        "entities": entities,
        "incomplete_reasons": reasons,
    }


def external_command_identity(value: str | None, base_dir: Path) -> dict[str, Any]:
    base_dir = base_dir.resolve()
    identity: dict[str, Any] = {
        "version": COMMAND_IDENTITY_VERSION,
        "base_dir": str(base_dir),
        "raw": value or "",
        "status": "unavailable" if not value else "invalid",
        "argv": None,
        "executable": None,
        "file_arguments": [],
        "dependency_closure": None,
        "cache_reusable": False,
        "error": None,
    }
    if not value:
        identity["error"] = "command_not_configured"
        return identity
    try:
        argv = parse_external_command(value)
    except ExternalCommandError as exc:
        identity["error"] = str(exc)
        return identity
    identity["argv"] = argv
    executable_result = _resolve_executable(argv[0], base_dir)
    if executable_result is None:
        identity["status"] = "unavailable"
        identity["executable"] = {
            "requested": argv[0],
            "located_path": None,
            "resolved_path": None,
            "sha256": None,
            "size_bytes": None,
        }
        identity["error"] = "executable_not_found"
        return identity
    executable_lexical, executable = executable_result
    identity["executable"] = _executable_identity(argv[0], executable_lexical)
    closure = _dependency_closure(argv, executable, base_dir)
    identity["dependency_closure"] = closure
    identity["file_arguments"] = [
        item
        for item in closure["entities"]
        if item.get("type") == "file"
    ]
    identity["cache_reusable"] = closure["cache_reusable"]
    identity["status"] = "available"
    return identity


def external_command_cache_reusable(value: object) -> bool:
    """External programs are never eligible for finalize-stage cache reuse."""
    return False


def external_command_identity_errors(label: str, value: object) -> list[str]:
    if not isinstance(value, dict) or not (
        type(value.get("version")) is int
        and value.get("version") == COMMAND_IDENTITY_VERSION
    ):
        return [f"{label} command identity is missing or unsupported"]
    status = value.get("status")
    if status != "available":
        return [f"{label} command identity is {status or 'invalid'}"]
    errors: list[str] = []
    raw = value.get("raw")
    base_dir = value.get("base_dir")
    closure = value.get("dependency_closure")
    if not isinstance(raw, str) or not isinstance(base_dir, str):
        errors.append(f"{label} command request is incomplete")
    elif not isinstance(closure, dict) or not (
        type(closure.get("version")) is int
        and closure.get("version") == DEPENDENCY_CLOSURE_VERSION
    ):
        errors.append(f"{label} dependency closure is missing or unsupported")
    else:
        try:
            if any(
                ord(character) < 32 or ord(character) == 127
                for character in raw + base_dir
            ):
                raise ValueError("control character")
            current = external_command_identity(raw, Path(base_dir))
            if current != value:
                errors.append(
                    f"{label} command resolution or dependency content changed"
                )
        except (OSError, RuntimeError, TypeError, ValueError):
            errors.append(f"{label} command path cannot be inspected")
    if value.get("cache_reusable") is not False:
        errors.append(f"{label} cache-reusable state conflicts with dependency closure")
    if isinstance(closure, dict):
        reasons = closure.get("incomplete_reasons")
        if closure.get("status") != "unproven" or closure.get("cache_reusable") is not False:
            errors.append(f"{label} external dependency closure claims unsupported reuse")
        if not isinstance(reasons, list) or not {
            "external_program_not_hermetic",
            "runtime_dependencies_unproven",
        }.issubset(reasons):
            errors.append(f"{label} external dependency closure lacks hermeticity reasons")
    executable = value.get("executable")
    if not isinstance(executable, dict) or not executable.get("resolved_path"):
        errors.append(f"{label} executable identity is missing")
    return errors
