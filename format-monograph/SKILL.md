---
name: format-monograph
description: Extract, confirm, apply, and audit formatting requirements for Chinese or multilingual monograph DOCX files without changing prose, facts, data, equations, answers, or citations. Use when an agent must analyze publisher specifications, Word templates, PDF rules, sample-book pages, images, or caller instructions; create an approved format profile; reformat a monograph; produce a review-annotated copy; or verify layout, editable equations, fields, and content preservation.
---

# Format Monograph

Apply only approved formatting rules. Preserve the source DOCX and all authored content.

## Start

1. Locate this skill directory and resolve every referenced path relative to it.
2. Read [capability-levels.md](references/capability-levels.md).
3. For a whole book or a resumed run, read [portable-run-checklist.md](references/portable-run-checklist.md) and [whole-book-runtime.md](references/whole-book-runtime.md).
4. Run `<python> scripts/check_environment.py --json`. When auto-discovery fails, use `--renderer <path>` or set `FORMAT_MONOGRAPH_RENDERER`.
5. Select the reported capability mode. Never claim a higher mode.
6. Keep source materials and generated documents outside this skill directory.

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

Read [monograph-elements.md](references/monograph-elements.md) while building selectors and properties. When the caller selects the technical-textbook baseline, load `examples/profiles/technical-textbook-layout.v0.2.5.draft.json` as a candidate and make a task-specific approved copy outside the skill directory. Earlier candidate profiles remain historical inputs and must not be silently upgraded.

## Apply approved rules

For a whole book, use the portable staged interface and `--resume` after an interruption:

```text
<python> scripts/run_monograph.py prepare <input.docx> --profile <profile.json> --work-dir <directory>
<python> scripts/run_monograph.py apply --work-dir <directory> --structure-map <approved.json>
<python> scripts/run_monograph.py finalize --work-dir <directory>
<python> scripts/run_monograph.py verify --work-dir <directory> --visual-qa-manifest <visual-qa.json>
```

The detailed commands below remain the individual building blocks for analysis, focused diagnosis, and adapter integration.

1. Read [structure-map.md](references/structure-map.md).
2. Run:
   `<python> scripts/inspect_docx.py <input.docx> --output <inventory.json> --structure-map-output <candidate-structure-map.json>`
3. Review unsupported elements, missing fonts, damaged relationships, ambiguous paragraph roles, formula-image candidates, legacy equation objects, static TOC ranges, chapter start and heading progression, numbered and unnumbered figure captions, image placement classes and resize safety, table kinds, figure-panel layout tables, exact header and semantic-separator rows, column roles, complex merges, visible control marks, wide tables, TOC/body page-number boundaries, odd/even footers, and trailing-section evidence.
4. Ask the caller to approve each proposed structural operation. Keep uncertain entries unapproved and report them; never infer approval from profile approval.
   For schema 1.5, mark body/title/front-matter roles explicitly, approve `pagination_sections` only after identifying the TOC and body starts, and classify each table as data, three-line, grid, figure-panel layout, pagination-only layout, callout, or unknown. Group repeated decisions and list object-level exceptions. Keep unresolved objects in `frozen_scopes`; approved independent scopes may continue, but finalization remains blocked. Approve data-table visuals only with a known role for every column. Approve figure-panel formatting only when image rows and their short label rows are explicitly mapped. Caption identifiers and all appendix identifiers remain editable manual text by default. Do not infer a missing number, sequence error, or `SEQ` conversion from punctuation alone.
   In architecture, civil-engineering, structural-engineering, or drafting content, analyze whether hyphenated numbers identify a section, elevation, node, detail, or drawing callout. Preserve uncertain identifiers and ask the caller before changing them.
   Approve image-visibility repair separately from resizing. It is automatic only for an image-only inline paragraph with proven fixed-line clipping, or a simple nonmerged table row with proven insufficient exact height. Mixed text, floating images, text boxes, merged rows, and ambiguous containers remain unchanged for QA.
   Approve `toc_source` only after every included entry maps one-to-one to an approved semantic heading or confirmed appendix. Use heading styles only when every source paragraph is object-free and no unapproved outline paragraph can enter the TOC; otherwise use one reserved pure-text `TC` source set for the entire TOC.
5. Set the map status to `approved`, then bind it to the unchanged source:
   `<python> scripts/validate_structure_map.py <approved-structure-map.json> --source <input.docx>`
6. Run:
   `<python> scripts/apply_profile.py <input.docx> --profile <profile.json> --structure-map <approved-structure-map.json> --output-dir <directory>`
7. If required fonts are missing, stop for QA. Use `--allow-missing-fonts` only after the caller explicitly approves structural output without those fonts; the report records the override.
8. Never overwrite the input file.
9. Treat `manual_review` rules as review items. Do not simulate an automatic change.
10. Preserve all authored text except an individually caller-confirmed caption-identifier replacement. Such a replacement must use `action=replace_identifier`, preserve the caption title exactly, and pass the dedicated audit. Approved field display values may also change.
11. Run:
   `<python> scripts/audit_docx.py <input.docx> <formatted.docx> --profile <profile.json> --structure-map <approved-structure-map.json> --output <audit.json>`
12. Finalize editable field caches without overwriting either input:
   `<python> scripts/finalize_docx.py <formatted.docx> --source <input.docx> --profile <profile.json> --structure-map <approved-structure-map.json> --output <finalized.docx> --status-output <finalization.json>`

When an approved target-application backend is available, call the finalizer with `--field-updater external --field-updater-command <command>`. The backend is a disposable field-calculation service, never the delivery parent. The core imports only uniquely matched, approved field results into its audited baseline, discards backend OOXML serialization, then asks the backend to reopen the selective output without saving and export the target PDF. Require `field_writeback_status=selective_verified`; any authored-content, field-instruction, bookmark-target, pagination-structure, OMML, embedded-object, or media change is blocking. Without an external backend, `auto` may use LibreOffice UNO. Fall back to update-on-open only after explicit caller QA, using `--approve-deferred`; record `delivery_field_status=deferred`. Never describe dirty flags as a completed refresh or LibreOffice pagination as Microsoft Word pagination.

For a whole book, use one source-bound structure map that covers the entire DOCX. A chapter trial is evidence for tuning the rules, not the final scope of the skill.

## Fonts and themes

- Treat a font rule as satisfied only when the effective Word font matches after resolving direct formatting, character style, paragraph style, base styles, document defaults, and the theme font scheme.
- For every approved automatic font rule, write explicit `ascii`, `hAnsi`, `eastAsia`, and, when applicable, `cs` values on the controlled style. Remove conflicting `asciiTheme`, `hAnsiTheme`, `eastAsiaTheme`, and `cstheme` references from that style and from controlled direct formatting.
- Do not change theme definitions globally. Unapproved roles, character styles, formula objects, symbols, hyperlinks, superscript/subscript, and other uncontrolled content retain their original formatting.
- Inspect and report both the declared font and the resolved effective font. A theme-resolved mismatch, including DengXian/等线 or DengXian Light/等线 Light inherited from the source theme, is an audit failure rather than an acceptable fallback.
- Apply the same deterministic check after target-software field refresh. Reject the refreshed copy when any approved effective font changes.

## Equations

- Keep Word OMML and editable MathType/OLE equations editable and unchanged.
- Convert caller-supplied LaTeX to editable OMML when conversion is supported and verified.
- Never replace an equation with PNG, JPEG, SVG, EMF, or another graphic.
- Do not automatically OCR formula images. Report each candidate as blocking and request an editable source or caller-confirmed reconstruction.
- Report legacy Equation Editor objects and obtain a QA decision before processing.
- Compare OMML, embedded-object, and media hashes before and after formatting.

## Fields and numbering

Use real Word fields and linked numbering only when the profile and structure map explicitly approve them. Figure and table identifiers default to editable manual text, not `SEQ`; preserve existing `SEQ`, `REF`, and `PAGEREF` fields. A new caption conversion requires both an approved profile rule with `numbering_mode=seq_field` and `allow_automatic_renumbering=true`, plus an individually approved `convert_to_seq` map entry. Explicit markers are `[[TOC]]`, `[[PAGE]]`, `[[SEQ:name]]`, `[[REF:name]]`, and `[[PAGEREF:name]]`. Remove manual heading prefixes only when they match an approved, unambiguous pattern. Use the approved `chapter_start` for a chapter excerpt and validate progression across a whole book. Otherwise stop and ask.

The TOC result is a text-only derived field. Before selective writeback, require exactly one result entry for every approved source in the same order and level, with nonempty approved title text, an internal target, and a page value. Reject the complete refreshed TOC when it contains DrawingML, VML, OLE, a text box, a table, an external relationship, an empty entry, a duplicate/extra entry, or an unapproved title. Never move or delete an image to clean the TOC.

For the technical-textbook baseline, create separate TOC and body sections only through approved `pagination_sections`. Start each at decimal 1, continue numbering through later body and landscape sections, show the first-page number, and require both default/odd and even footers with exactly one editable PAGE field per page-number footer. Repeated application must not add another field. A footer containing publisher text, a logo, another field, or mixed payload is blocking QA; replace static digits only when the structure map explicitly approves them as derived page text. Never treat the number of PAGE fields as the physical page count.

Treat the approved whole-book title as a front-matter object, not a chapter title. Put it alone on an unnumbered page and apply the caller-approved title format; the technical-textbook fallback is centered bold 22 pt Chinese Heiti with Times New Roman for western text, at least 33 pt line spacing, zero paragraph spacing, and vertical centering within the title-page section. Start a separate TOC page with the derived centered bold heading `目    录` (four ASCII spaces) and restart visible decimal TOC pagination at 1; the `TOC` field generates entries but not this heading, so the skill inserts and maintains it automatically. Restart body pagination at 1 independently. Generated multilevel heading numbers must inherit the same effective Chinese/western fonts, size, and weight as their linked Heading style, and Heading 1-4 paragraphs must have zero first-line indent.

For approved technical-textbook data tables, clear inherited table and cell border overrides before rebuilding the approved model. The default major rule is 1.0 pt, minor rules are 0.5 pt, outside left/right borders are absent, and shading is absent. Do not force one horizontal-line topology on every table: map header boundaries and semantic summary/group separator rows explicitly; use internal vertical minor rules unless the caller approves a three-line variant. A multi-row header receives minor separators only where the hierarchy splits and never through a vertically merged cell. Unknown or special tables remain unchanged for QA.

Treat a table whose primary payload is images with short labels beneath them as a figure-panel layout only after its image rows and label rows are approved. Remove all table and cell borders, center the inline table, disable text wrapping, center its cells, and apply the figure-caption format to label text without changing that text or adding a figure number. Likewise, a short centered paragraph immediately following a standalone image may be approved as an unnumbered figure caption; style it as a figure caption but preserve its wording. Remove a leading empty table-cell paragraph only through an exact source-bound cleanup entry; record it as approved whitespace normalization and never remove a nonempty paragraph.

Never relocate, reinsert, reorder, float, or re-anchor an image or table. Record `position_policy=preserve_anchor` for every mapped image and table. Approved resizing may change only an inline image's display extents and alignment within its existing paragraph or cell; preserve its relationship, media payload, crop state, aspect ratio, container, object order, and surrounding-text order. For a standalone image, fit within 90% of text width, or 100% when its aspect ratio is at least 1.6, and within 65% of usable page height. For an approved figure-panel image, fit within 95% of its existing cell width and use one displayed height across images in the same row. Limit raster enlargement to 125% and require at least 220 effective DPI; do not enlarge when source DPI is unknown or insufficient. Vector images may fit the approved bounds. Leave cropped, floating, ambiguous, or other table-embedded images unchanged for QA. Natural repagination caused by an approved in-place resize is allowed, but the image or table must remain between the same authored neighbors.

An approved image-only inline paragraph that is clipped by exact line spacing uses direct automatic single-line spacing so the line box expands to the existing image height. Preserve alignment, paragraph spacing, pagination properties, display extent, crop, relationship, and object order. In a simple table row, change an insufficient exact row height to `atLeast` only when separately approved. Do not apply either repair to mixed-text paragraphs, floating images, merged rows, text boxes, or uncertain layout tables.

Treat mirror margins versus consecutive physical pages as an explicit layout conflict. Microsoft Word can insert parity blank pages when a mirrored document restarts a new section at page 1. Use a no-save target-layout measurement, then resolve the approved section starts, displayed page-number rules, and PAGE/PAGEREF instructions deterministically in the portable core before field calculation. When an immediate page-1 restart begins on an even physical page, use the editable `{ = { PAGE } - 1 }` display field only for that approved page-number footer; do not approve arbitrary formula fields. The Microsoft Word adapter must report zero structural changes and may only measure, repaginate, update approved fields, save its disposable refresh copy, and perform the later no-save verification render. The caller does not manually update the TOC or page fields. Instruct the caller to rerun the skill after manuscript pagination changes. Never silently retain a parity blank page, disable mirror margins, or import a Word-modified section or footer.

After an approved table or complete figure block, insert exactly one real empty paragraph when following content remains on the same page. Do not replace it with paragraph spacing. Mark the paragraph as an approved derived structure, make repeated runs idempotent, and remove it after target-software repagination when it would otherwise become a page-top blank line.

## Render and verify

In full mode, run:

`<python> scripts/render_docx.py <finalized.docx> --output-dir <render-directory> [--renderer <path>] [--target-software <name>] [--target-pdf <target-export.pdf>]`

Open every generated page image at full readable size. Check every page for clipping, overlap, missing glyphs, broken tables, bad page breaks, misplaced figures, incorrect headers or footers, and visibly rasterized or missing equations. Fix, re-audit, and re-render after every layout-sensitive change.

Classify every blank page as `intentional_recto_blank`, `removable_trailing_blank`, or `unexpected_blank`. An intentional recto blank preserves the approved next-odd-page structure and carries no visible page number; do not delete it. Remove a trailing blank only with approved stable cleanup evidence. Treat every other blank as a layout defect or blocking QA.

Rendering with LibreOffice does not prove Microsoft Word layout compatibility. When an approved backend exports a target-software PDF, render that PDF with `--target-pdf`; it becomes target-layout evidence only after every page image is inspected. Otherwise record `target_layout_unverified` and require a final check in the target application.

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
- field finalization changes authored content, protected objects, or the editable-field contract;
- an approved TOC source does not map one-to-one to text-only semantic headings, or a refreshed TOC contains any non-text object, empty/extra entry, wrong level, wrong order, missing internal target, or unverifiable page value;
- an approved automatic font does not match its resolved effective Word font before or after finalization;
- an approved pagination map lacks a TOC/body boundary, an even PAGE footer, or contains an unapproved body restart;
- a page-number footer contains multiple PAGE fields or non-page payload without explicit structural approval;
- a table has unknown column roles, complex merges, floating objects, visible control marks, or landscape layout without individual approval;
- a figure-panel layout, semantic table separator, unnumbered figure caption, or table-cell cleanup is ambiguous or lacks a source-bound approval;
- an image or table cannot preserve its original anchor, container, object order, or surrounding-text order;
- an image visibility defect occurs in mixed text, a floating object, text box, merged table row, or another container without safe structural evidence;
- a structure map does not match the current source fingerprint or is not approved;
- rendering fails for a reason other than an unavailable renderer.

Never silently substitute fonts, infer uncovered rules, edit prose, generate answers, rasterize formulas, or report unperformed validation as passed.
