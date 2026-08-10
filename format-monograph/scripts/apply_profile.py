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
        "# 涓撹憲鏍煎紡淇敼鎶ュ憡",
        "",
        "## 鍩烘湰淇℃伅",
        "",
        f"- 杈撳叆鏂囦欢锛歚{input_path.name}`",
        f"- 鏍煎紡閰嶇疆锛歚{profile['profile_id']}` / {profile['name']}",
        f"- 閰嶇疆鐗堟湰锛歚{profile['schema_version']}`",
        f"- 鎵瑰噯浜猴細{profile['approval'].get('approved_by', '')}",
        f"- 鎵瑰噯鏃堕棿锛歿profile['approval'].get('approved_at', '')}",
        f"- 鐢熸垚鏃堕棿锛歿dt.datetime.now(dt.timezone.utc).isoformat()}",
        "",
        "## 杩愯鏃朵緷鎹?,
        "",
        f"- 璋冪敤鑰呮槑纭姹傛渶楂橈細{profile.get('runtime_policy', {}).get('caller_requirements_highest', False)}",
        "- 鏉ユ簮浼樺厛绾э細" + " > ".join(profile.get("source_precedence", [])),
        "",
        "## 瀛椾綋棰勬",
        "",
        "- 缂哄け瀛椾綋锛? + (", ".join(missing_fonts) if missing_fonts else "鏃?),
        "- 瀛椾綋瑙ｆ瀽锛?
        + (
            "; ".join(
                f"{item['requested']} -> {item['matched_name'] or 'missing'} ({item['match']})"
                for item in font_resolutions
            )
            if font_resolutions
            else "鏃?
        ),
        f"- 鐢ㄦ埛鎵瑰噯缂哄け瀛椾綋闄嶇骇锛歿missing_fonts_approved}",
        "",
        "## 鍐呭涓€鑷存€?,
        "",
        f"- 缁撴灉锛?*{integrity}**",
        f"- 鍘熺鎸囩汗锛歚{original_fp}`",
        f"- 鏍煎紡绋挎寚绾癸細`{formatted_fp}`",
        "- 鎸囩汗鎺掗櫎瀛楁鏄剧ず缁撴灉锛涜幏鎵瑰瓧娈甸噸寤鸿繕浼氳鑼冨寲鏄庣‘璇嗗埆鐨勬淳鐢熺紪鍙枫€?,
        f"- OMML銆佸祵鍏ュ璞″拰濯掍綋鍝堝笇锛歿'PASS' if protected_objects_ok else 'FAIL'}",
        "",
        "## 鍏紡瀵硅薄",
        "",
        f"- OMML锛歿equation_summary['omml']}",
        f"- MathType OLE锛歿equation_summary['mathtype_ole']}",
        f"- 鏃х増 Equation Editor锛歿equation_summary['legacy_equation_ole']}",
        f"- 鍏紡鍥剧墖鍊欓€夛細{equation_summary['formula_image_candidates']}",
        "",
        "## 宸插簲鐢ㄨ鍒?,
        "",
        "| 瑙勫垯 | 閫夋嫨鍣?| 鍛戒腑鏁伴噺 | 灞炴€?|",
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
        lines.append("| - | - | 0 | 娌℃湁鑷姩瑙勫垯 |")

    lines.extend(["", "## 娲剧敓瀛楁鍙樻洿", ""])
    if derived_changes:
        for change in derived_changes:
            lines.append(f"- `{change['kind']}`锛歿json.dumps(change, ensure_ascii=False)}")
    else:
        lines.append("- 鏃犮€?)

    lines.extend(["", "## 浜哄伐澶嶆牳", ""])
    if manual:
        for rule in manual:
            lines.append(
                f"- `{rule['id']}`锛歿rule['evidence_summary']} "
                f"锛堥€夋嫨鍣細`{rule['selector']['kind']}:{rule['selector']['value']}`锛?
            )
    else:
        lines.append("- 鏃犮€?)

    lines.extend(["", "## 瀹￠槄鏍囨敞", ""])
    if unanchored:
        lines.append("- 浠ヤ笅瑙勫垯鏃犳硶绮剧‘閿氬畾锛屽凡浠呭湪鏈姤鍛婅褰曪細" + ", ".join(unanchored))
    else:
        lines.append("- 鎵€鏈夎嚜鍔ㄨ鍒欏潎宸插湪瀹￠槄绋夸腑寤虹珛鎵规敞閿氱偣銆?)

    lines.extend(
        [
            "",
            "## 娓叉煋涓庤瑙?QA",
            "",
            "- 鐘舵€侊細寰呰繍琛?`render_docx.py` 骞堕€愰〉浜哄伐妫€鏌ャ€?,
            "- 鍦ㄥ畬鎴愰€愰〉妫€鏌ュ墠锛屼笉寰楀皢鏈姤鍛婄粨璁烘敼涓衡€滈€氳繃鈥濄€?,
            "",
            "## 缁撹",
            "",
            "- 褰撳墠缁撹锛?
            + (
                "缁撴瀯妫€鏌ラ€氳繃锛岀瓑寰呰瑙?QA銆?
                if integrity == "PASS"
                else "澶辫触锛氬唴瀹规垨鍙紪杈戝璞′竴鑷存€ф鏌ユ湭閫氳繃銆?
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

