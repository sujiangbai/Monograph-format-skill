---
name: format-monograph
description: Extract, confirm, apply, and audit formatting requirements for Chinese or multilingual monograph DOCX files without changing prose, facts, data, or citations. Use when an agent must analyze publisher specifications, Word templates, PDF rules, sample-book pages, images, or written formatting instructions; create an approved format profile; reformat a monograph; produce a review-annotated copy; or verify monograph layout and content preservation.
---

# Format Monograph

Apply only approved formatting rules. Preserve the source DOCX and all authored content.

## Start

1. Locate this skill directory and resolve every referenced path relative to it.
2. Read [capability-levels.md](references/capability-levels.md).
3. Run `<python> scripts/check_environment.py --json`.
4. Select the reported capability mode. Never claim a higher mode.
5. Keep source materials and generated documents outside this skill directory.

## Build a format profile

1. Read [rule-extraction.md](references/rule-extraction.md).
2. Inspect every supplied DOCX, PDF, image, and written requirement using available local tools.
3. Apply source precedence: written requirement, DOCX template, PDF specification, sample book, image, text note.
4. Create a JSON profile matching [format-profile.schema.json](references/format-profile.schema.json).
5. Paraphrase evidence. Do not copy long passages or store unpublished source files in the skill.
6. Run `<python> scripts/validate_profile.py <profile.json>`.
7. Present conflicts, low-confidence rules, missing values, and blocking questions to the user.
8. Stop before document modification until `approval.status` is `approved` and every blocking question and conflict is resolved.

Read [monograph-elements.md](references/monograph-elements.md) while building selectors and properties.

## Apply approved rules

1. Run `<python> scripts/inspect_docx.py <input.docx> --output <inventory.json>`.
2. Review unsupported elements, missing fonts, damaged relationships, and ambiguous paragraph roles.
3. Run:
   `<python> scripts/apply_profile.py <input.docx> --profile <profile.json> --output-dir <directory>`
4. Never overwrite the input file.
5. Treat `manual_review` rules as review items. Do not simulate an automatic change.
6. Preserve all authored text. Only field display values generated from approved TOC, numbering, caption, page-number, or cross-reference fields may change.
7. Run:
   `<python> scripts/audit_docx.py <input.docx> <formatted.docx> --profile <profile.json> --output <audit.json>`

## Render and verify

In full mode, run:

`<python> scripts/render_docx.py <formatted.docx> --output-dir <render-directory>`

Open every generated page image at full readable size. Check every page for clipping, overlap, missing glyphs, broken tables, bad page breaks, misplaced figures, and incorrect headers or footers. Fix, re-audit, and re-render after every layout-sensitive change.

In structural mode, ask the user to approve delivery without visual QA. State the limitation in the report. In analysis mode, do not modify a DOCX.

## Deliver

Deliver exactly:

- `<stem>-formatted.docx`: clean formatted copy.
- `<stem>-review.docx`: formatted copy with anchored comments describing applied rule IDs; record unanchorable page-level changes in the report.
- `<stem>-format-report.md`: profile identity, changes, manual items, conflicts, content-integrity result, render status, and limitations.

Read [qa-and-reporting.md](references/qa-and-reporting.md) before asking questions or delivering files.

## Hard stops

Stop and ask the user when:

- the profile is not approved;
- a blocking question or open conflict remains;
- a required font is unavailable;
- a requested automatic property is unsupported;
- the input is not a valid DOCX;
- content-integrity audit fails;
- rendering fails for a reason other than an unavailable renderer.

Never silently substitute fonts, infer uncovered rules, edit prose, or report unperformed validation as passed.
