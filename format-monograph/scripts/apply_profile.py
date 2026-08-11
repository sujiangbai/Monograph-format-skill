#!/usr/bin/env python3
"""Apply an approved format profile and produce the three review artifacts."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from _common import (
    FormatMonographError,
    apply_rule,
    content_fingerprint,
    equation_inventory,
    field_inventory,
    first_anchor_paragraph,
    load_document,
    profile_font_resolutions,
    protected_object_manifest,
    summarize_rule,
)
from validate_profile import validate
from docx_pagination import finalize_pagination_sections
from structure_map import (
    approved_data_tables,
    approved_role_paragraphs,
    apply_structure_map,
    has_semantic_structure_map,
    load_structure_map,
    prime_structure_map_locators,
    resolve_paragraph_locator,
    structure_content_fingerprint,
    validate_structure_map_source,
)


def output_paths(input_path: Path, output_dir: Path) -> tuple[Path, Path, Path]:
    stem = input_path.stem
    return (
        output_dir / f"{stem}-formatted.docx",
        output_dir / f"{stem}-review.docx",
        output_dir / f"{stem}-format-report.md",
    )


def assert_outputs_available(paths: tuple[Path, ...], force: bool) -> None:
    existing = [str(path) for path in paths if path.exists()]
    if existing and not force:
        raise FormatMonographError(
            "Output files already exist; use --force to replace them: " + ", ".join(existing)
        )


def uses_derived_normalization(profile: dict) -> bool:
    keys = {
        "convert_explicit_markers",
        "rebuild_heading_numbering",
        "strip_manual_heading_prefixes",
    }
    return any(
        rule.get("status") == "approved"
        and rule.get("application") == "automatic"
        and rule.get("selector", {}).get("kind") == "field_role"
        and any(rule.get("properties", {}).get(key) for key in keys)
        for rule in profile.get("rules", [])
    )


def assert_caption_actions_authorized(profile: dict, structure_map: dict) -> None:
    requests_seq = any(
        entry.get("approved") and entry.get("action") == "convert_to_seq"
        for entry in structure_map.get("captions", [])
    )
    if not requests_seq:
        return
    allows_seq = any(
        rule.get("status") == "approved"
        and rule.get("application") == "automatic"
        and rule.get("selector", {}).get("kind") == "caption_role"
        and rule.get("properties", {}).get("numbering_mode") == "seq_field"
        and rule.get("properties", {}).get("allow_automatic_renumbering") is True
        for rule in profile.get("rules", [])
    )
    if not allows_seq:
        raise FormatMonographError(
            "SEQ caption conversion requires an approved caption profile rule with "
            "numbering_mode=seq_field and allow_automatic_renumbering=true."
        )


def preflight_fields(input_path: Path, profile: dict) -> dict:
    inventory = field_inventory(input_path)
    rebuilds_fields = any(
        rule.get("status") == "approved"
        and rule.get("application") == "automatic"
        and rule.get("selector", {}).get("kind") == "field_role"
        for rule in profile.get("rules", [])
    )
    if rebuilds_fields and inventory["unresolved_references"]:
        raise FormatMonographError(
            "Field policy blocked the document because REF/PAGEREF targets are missing: "
            + ", ".join(inventory["unresolved_references"])
        )
    return inventory


def preflight_equations(input_path: Path, profile: dict) -> dict:
    inventory = equation_inventory(input_path)
    policy = profile.get("runtime_policy", {})
    if not policy.get("editable_equations_required"):
        return inventory
    if inventory["formula_image_candidates"]:
        raise FormatMonographError(
            "Editable-equation policy blocked the document: "
            f"{inventory['formula_image_candidates']} formula image candidate(s) found."
        )
    if (
        inventory["legacy_equation_ole"]
        and policy.get("legacy_equation_policy", "qa") == "qa"
    ):
        raise FormatMonographError(
            "Legacy Equation Editor objects require QA before formatting: "
            f"{inventory['legacy_equation_ole']} object(s)."
        )
    return inventory


def report_markdown(
    input_path: Path,
    profile: dict,
    original_fp: str,
    formatted_fp: str,
    changes: list[dict],
    manual: list[dict],
    unanchored: list[str],
    derived_changes: list[dict],
    equation_summary: dict,
    protected_objects_ok: bool,
    missing_fonts: list[str],
    font_resolutions: list[dict],
    missing_fonts_approved: bool,
) -> str:
    integrity = "PASS" if original_fp == formatted_fp and protected_objects_ok else "FAIL"
    lines = [
        "# 专著格式修改报告",
        "",
        "## 基本信息",
        "",
        f"- 输入文件：`{input_path.name}`",
        f"- 格式配置：`{profile['profile_id']}` / {profile['name']}",
        f"- 配置版本：`{profile['schema_version']}`",
        f"- 批准人：{profile['approval'].get('approved_by', '')}",
        f"- 批准时间：{profile['approval'].get('approved_at', '')}",
        f"- 生成时间：{dt.datetime.now(dt.timezone.utc).isoformat()}",
        "",
        "## 运行时依据",
        "",
        f"- 调用者明确要求最高：{profile.get('runtime_policy', {}).get('caller_requirements_highest', False)}",
        "- 来源优先级：" + " > ".join(profile.get("source_precedence", [])),
        "",
        "## 字体预检",
        "",
        "- 缺失字体：" + (", ".join(missing_fonts) if missing_fonts else "无"),
        "- 字体解析："
        + (
            "; ".join(
                f"{item['requested']} -> {item['matched_name'] or 'missing'} ({item['match']})"
                for item in font_resolutions
            )
            if font_resolutions
            else "无"
        ),
        f"- 用户批准缺失字体降级：{missing_fonts_approved}",
        "",
        "## 内容一致性",
        "",
        f"- 结果：**{integrity}**",
        f"- 原稿指纹：`{original_fp}`",
        f"- 格式稿指纹：`{formatted_fp}`",
        "- 指纹排除字段显示结果；获批字段重建还会规范化明确识别的派生编号。",
        f"- OMML、嵌入对象和媒体哈希：{'PASS' if protected_objects_ok else 'FAIL'}",
        "",
        "## 公式对象",
        "",
        f"- OMML：{equation_summary['omml']}",
        f"- MathType OLE：{equation_summary['mathtype_ole']}",
        f"- 旧版 Equation Editor：{equation_summary['legacy_equation_ole']}",
        f"- 公式图片候选：{equation_summary['formula_image_candidates']}",
        "",
        "## 已应用规则",
        "",
        "| 规则 | 选择器 | 命中数量 | 属性 |",
        "| --- | --- | ---: | --- |",
    ]
    for change in changes:
        props = ", ".join(
            f"{key}={value}" for key, value in sorted(change["properties"].items())
        )
        lines.append(
            f"| {change['id']} | {change['selector']} | {change['targets']} | {props} |"
        )
    if not changes:
        lines.append("| - | - | 0 | 没有自动规则 |")

    lines.extend(["", "## 派生字段变更", ""])
    if derived_changes:
        for change in derived_changes:
            lines.append(f"- `{change['kind']}`：{json.dumps(change, ensure_ascii=False)}")
    else:
        lines.append("- 无。")

    lines.extend(["", "## 人工复核", ""])
    if manual:
        for rule in manual:
            lines.append(
                f"- `{rule['id']}`：{rule['evidence_summary']} "
                f"（选择器：`{rule['selector']['kind']}:{rule['selector']['value']}`）"
            )
    else:
        lines.append("- 无。")

    lines.extend(["", "## 审阅标注", ""])
    if unanchored:
        lines.append("- 以下规则无法精确锚定，已仅在本报告记录：" + ", ".join(unanchored))
    else:
        lines.append("- 所有自动规则均已在审阅稿中建立批注锚点。")

    lines.extend(
        [
            "",
            "## 渲染与视觉 QA",
            "",
            "- 状态：待运行 `render_docx.py` 并逐页人工检查。",
            "- 在完成逐页检查前，不得将本报告结论改为“通过”。",
            "",
            "## 结论",
            "",
            "- 当前结论："
            + (
                "结构检查通过，等待视觉 QA。"
                if integrity == "PASS"
                else "失败：内容或可编辑对象一致性检查未通过。"
            ),
            "",
        ]
    )
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--profile", required=True, type=Path)
    parser.add_argument(
        "--structure-map",
        type=Path,
        help="Caller-approved structure map generated during inspection.",
    )
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument(
        "--allow-missing-fonts",
        action="store_true",
        help="Continue only after the caller explicitly approves missing-font visual QA.",
    )
    args = parser.parse_args()

    try:
        errors, profile = validate(args.profile)
        if errors:
            raise FormatMonographError("Profile validation failed: " + "; ".join(errors))
        if profile["approval"]["status"] != "approved":
            raise FormatMonographError("Profile approval.status must be approved.")
        structure_map = None
        if args.structure_map:
            structure_map = load_structure_map(args.structure_map)
            validate_structure_map_source(args.input, structure_map)
            assert_caption_actions_authorized(profile, structure_map)

        formatted_path, review_path, report_path = output_paths(args.input, args.output_dir)
        args.output_dir.mkdir(parents=True, exist_ok=True)
        assert_outputs_available((formatted_path, review_path, report_path), args.force)
        if args.input.resolve() in {formatted_path.resolve(), review_path.resolve()}:
            raise FormatMonographError("Output path must not overwrite the input DOCX.")

        normalize_derived = uses_derived_normalization(profile)
        preflight_fields(args.input, profile)
        equation_summary = preflight_equations(args.input, profile)
        font_resolutions = profile_font_resolutions(profile)
        missing_fonts = [
            item["requested"] for item in font_resolutions if not item["available"]
        ]
        if missing_fonts and not args.allow_missing_fonts:
            raise FormatMonographError(
                "Required fonts are unavailable: " + ", ".join(missing_fonts)
                + ". Obtain caller QA approval before using --allow-missing-fonts."
            )
        original_objects = protected_object_manifest(args.input)
        original_fp = (
            structure_content_fingerprint(args.input, structure_map)
            if structure_map
            else content_fingerprint(args.input, normalize_derived=normalize_derived)
        )
        document = load_document(args.input)
        changes: list[dict] = []
        manual: list[dict] = []

        if structure_map:
            prime_structure_map_locators(document, structure_map)
            for change in apply_structure_map(document, structure_map):
                _changes = getattr(document, "_format_monograph_derived_changes", [])
                _changes.append(change)
                setattr(document, "_format_monograph_derived_changes", _changes)

        for rule in profile["rules"]:
            if rule["status"] != "approved":
                continue
            if rule["application"] == "manual_review":
                manual.append(rule)
                continue
            kind = rule["selector"]["kind"]
            paragraph_targets = None
            table_targets = None
            chapter_start = None
            if structure_map and has_semantic_structure_map(structure_map):
                if kind in {"paragraph_role", "caption_role", "bibliography_role"}:
                    paragraph_targets = approved_role_paragraphs(
                        document, structure_map, rule["selector"]
                    )
                elif kind == "table_role":
                    table_targets = approved_data_tables(document, structure_map)
                numbering = structure_map.get("numbering", {})
                if kind == "field_role" and numbering.get("approved"):
                    chapter_start = int(numbering["chapter_start"])
            targets = apply_rule(
                document,
                rule,
                paragraph_targets=paragraph_targets,
                table_targets=table_targets,
                chapter_start=chapter_start,
            )
            changes.append(
                {
                    "id": rule["id"],
                    "selector": f"{rule['selector']['kind']}:{rule['selector']['value']}",
                    "targets": targets,
                    "properties": rule["properties"],
                }
            )

        if structure_map and finalize_pagination_sections(
            document,
            structure_map.get("pagination_sections", {}),
            resolve_paragraph_locator,
            front_matter=structure_map.get("front_matter", {}),
        ):
            _changes = getattr(document, "_format_monograph_derived_changes", [])
            _changes.append(
                {
                    "kind": "structure_pagination_page_break_deduplicated",
                    "reason": "body_section_already_starts_on_new_page",
                }
            )
            setattr(document, "_format_monograph_derived_changes", _changes)

        derived_changes = list(
            getattr(document, "_format_monograph_derived_changes", [])
        )
        document.save(str(formatted_path))
        formatted_fp = (
            structure_content_fingerprint(formatted_path, structure_map)
            if structure_map
            else content_fingerprint(formatted_path, normalize_derived=normalize_derived)
        )
        formatted_objects = protected_object_manifest(formatted_path)
        protected_objects_ok = original_objects == formatted_objects
        if original_fp != formatted_fp or not protected_objects_ok:
            formatted_path.unlink(missing_ok=True)
            raise FormatMonographError(
                "Integrity failed after formatting "
                f"(content={'pass' if original_fp == formatted_fp else 'fail'}, "
                f"protected_objects={'pass' if protected_objects_ok else 'fail'}, "
                f"original_fingerprint={original_fp}, "
                f"formatted_fingerprint={formatted_fp}). "
                "The generated formatted copy was removed."
            )

        review = load_document(formatted_path)
        unanchored: list[str] = []
        for rule in profile["rules"]:
            if rule["status"] != "approved" or rule["application"] != "automatic":
                continue
            anchor = None
            semantic_targeted = (
                structure_map
                and has_semantic_structure_map(structure_map)
                and rule["selector"]["kind"]
                in {"paragraph_role", "caption_role", "bibliography_role"}
            )
            if semantic_targeted:
                anchor = next(
                    (
                        paragraph
                        for paragraph in approved_role_paragraphs(
                            review, structure_map, rule["selector"]
                        )
                        if paragraph.runs
                    ),
                    None,
                )
            if anchor is None and not semantic_targeted:
                anchor = first_anchor_paragraph(review, rule)
            if anchor is None:
                unanchored.append(rule["id"])
                continue
            review.add_comment(
                runs=anchor.runs,
                text=summarize_rule(rule),
                author="format-monograph",
                initials="FM",
            )
        review.save(str(review_path))
        review_fp = (
            structure_content_fingerprint(review_path, structure_map)
            if structure_map
            else content_fingerprint(review_path, normalize_derived=normalize_derived)
        )
        if (
            review_fp != original_fp
            or protected_object_manifest(review_path) != original_objects
        ):
            review_path.unlink(missing_ok=True)
            formatted_path.unlink(missing_ok=True)
            raise FormatMonographError(
                "Content or protected-object integrity failed after adding review comments. "
                "Generated DOCX files were removed."
            )

        report_path.write_text(
            report_markdown(
                args.input,
                profile,
                original_fp,
                formatted_fp,
                changes,
                manual,
                unanchored,
                derived_changes,
                equation_summary,
                protected_objects_ok,
                missing_fonts,
                font_resolutions,
                bool(missing_fonts and args.allow_missing_fonts),
            ),
            encoding="utf-8",
        )
        print(
            json.dumps(
                {
                    "formatted": str(formatted_path),
                    "review": str(review_path),
                    "report": str(report_path),
                    "content_integrity": "pass",
                    "protected_object_integrity": "pass",
                    "derived_changes": len(derived_changes),
                    "missing_fonts": missing_fonts,
                    "font_resolutions": font_resolutions,
                    "missing_fonts_approved": bool(
                        missing_fonts and args.allow_missing_fonts
                    ),
                    "render_status": "pending",
                },
                ensure_ascii=False,
            )
        )
        return 0
    except (FormatMonographError, OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
