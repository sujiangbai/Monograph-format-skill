---
name: format-monograph
description: Extract, confirm, apply, and audit formatting requirements for Chinese or multilingual monograph DOCX files without changing prose, facts, data, equations, answers, or citations. Use when an agent must analyze publisher specifications, Word templates, PDF rules, sample-book pages, images, or caller instructions; create an approved format profile; reformat a monograph; produce a review-annotated copy; or verify layout, editable equations, fields, and content preservation.
---

# Format Monograph

Apply only approved formatting rules. Preserve the source DOCX and all authored content.

## Start

1. Locate this skill directory and resolve every referenced path relative to it.
2. Read [capability-levels.md](references/capability-levels.md).
3. Run `<python> scripts/check_environment.py --json`.
4. Select the reported capability mode. Never claim a higher mode.
5. Keep source materials and generated documents outside this skill directory.

## Resolve authority

1. Treat the current caller's explicit formatting requirements and later QA decisions as the highest formatting authority for that task.
2. Then apply any publisher requirement or DOCX template the caller designates for the task.
3. Then apply the caller-selected built-in profile.
4. Use GB/T defaults only for uncovered rules.
5. Use sample books, images, and inferred values only as supporting evidence.

Record every conflict. Follow a clear higher-priority instruction, but stop for QA when the caller's own instructions conflict, a source's applicability is unclear, or a change would cross a safety boundary. Caller priority never permits overwriting the source, changing authored meaning, rasterizing equations, silently substituting fonts, or skipping required audit and visual QA.

## Build a format profile

1. Read [rule-extraction.md](references/rule-extraction.md).
2. Inspect every supplied DOCX, PDF, image, and written requirement using available local tools.
3. Create a JSON profile matching [format-profile.schema.json](references/format-profile.schema.json).
4. Use schema `1.1` for new profiles and include the runtime policy. Continue to accept existing `1.0` profiles.
5. Paraphrase evidence. Do not copy long passages or store unpublished source files in the skill.
6. Run `<python> scripts/validate_profile.py <profile.json>`.
7. Present conflicts, low-confidence rules, missing values, and blocking questions to the caller.
8. Stop before document modification until `approval.status` is `approved` and every blocking question and conflict is resolved.

Read [monograph-elements.md](references/monograph-elements.md) while building selectors and properties. When the caller selects the technical-textbook baseline, load `examples/profiles/technical-textbook-layout.v0.2.draft.json` as a candidate and make a task-specific approved copy outside the skill directory.

## Apply approved rules

1. Run `<python> scripts/inspect_docx.py <input.docx> --output <inventory.json>`.
2. Review unsupported elements, missing fonts, damaged relationships, ambiguous paragraph roles, formula-image candidates, legacy equation objects, and ambiguous static numbering.
3. Run:
   `<python> scripts/apply_profile.py <input.docx> --profile <profile.json> --output-dir <directory>`
4. Never overwrite the input file.
5. Treat `manual_review` rules as review items. Do not simulate an automatic change.
6. Preserve all authored text. Only display values generated from approved TOC, numbering, caption, page-number, or cross-reference fields may change.
7. Run:
   `<python> scripts/audit_docx.py <input.docx> <formatted.docx> --profile <profile.json> --output <audit.json>`

## Equations

- Keep Word OMML and editable MathType/OLE equations editable and unchanged.
- Convert caller-supplied LaTeX to editable OMML when conversion is supported and verified.
- Never replace an equation with PNG, JPEG, SVG, EMF, or another graphic.
- Do not automatically OCR formula images. Report each candidate as blocking and request an editable source or caller-confirmed reconstruction.
- Report legacy Equation Editor objects and obtain a QA decision before processing.
- Compare OMML, embedded-object, and media hashes before and after formatting.

## Fields and numbering

Use real Word fields and linked numbering only when the profile explicitly approves them. Explicit markers are `[[TOC]]`, `[[PAGE]]`, `[[SEQ:name]]`, `[[REF:name]]`, and `[[PAGEREF:name]]`. Remove manual heading prefixes only when they match an approved, unambiguous pattern. Otherwise stop and ask.

## Render and verify

In full mode, run:

`<python> scripts/render_docx.py <formatted.docx> --output-dir <render-directory>`

Open every generated page image at full readable size. Check every page for clipping, overlap, missing glyphs, broken tables, bad page breaks, misplaced figures, incorrect headers or footers, and visibly rasterized or missing equations. Fix, re-audit, and re-render after every layout-sensitive change.

In structural mode, ask the caller to approve delivery without visual QA and state the limitation in the report. In analysis mode, do not modify a DOCX.

## Deliver

Deliver exactly:

- `<stem>-formatted.docx`: clean formatted copy.
- `<stem>-review.docx`: formatted copy with anchored comments describing applied rule IDs; record unanchorable page-level changes in the report.
- `<stem>-format-report.md`: authority decisions, profile identity, changes, approved derived-field changes, manual items, conflicts, equation and object integrity, content integrity, render status, and limitations.

Read [qa-and-reporting.md](references/qa-and-reporting.md) before asking questions or delivering files.

## Hard stops

Stop and ask the caller when:

- the profile is not approved;
- a blocking question or open conflict remains;
- a required font is unavailable;
- a requested automatic property is unsupported;
- the input is not a valid DOCX;
- a formula-image candidate or unresolved legacy equation object exists;
- content or protected-object integrity fails;
- field migration is ambiguous;
- rendering fails for a reason other than an unavailable renderer.

Never silently substitute fonts, infer uncovered rules, edit prose, generate answers, rasterize formulas, or report unperformed validation as passed.
