# Structure map

Use a source-bound structure map for operations that reinterpret document structure. Profile approval authorizes formatting rules; structure-map approval authorizes specific targets in one unchanged DOCX.

## Workflow

1. Generate an inventory and candidate map with `inspect_docx.py --structure-map-output`.
2. Review candidates with the caller. Approve only unambiguous targets and set top-level `status` to `approved`.
3. Run `validate_structure_map.py <map.json> --source <input.docx>` immediately before applying it.
4. Pass the same map to `apply_profile.py` and `audit_docx.py`.
5. Regenerate the map and repeat QA whenever the source fingerprint changes.

New maps use schema `1.4`; readers continue to accept `1.0` through `1.3`. Version 1.0 can authorize its original TOC, heading, caption, first-row table, and trailing-section operations. Version 1.1 retains its explicit legacy `SEQ` conversion behavior. Version 1.2 adds domain-aware manual caption actions and semantic paragraph roles. Version 1.3 adds stable locators and pagination-only groups.

## Schema 1.4

Schema 1.4 adds approved page-number sections, stable trailing-section evidence, per-table visual decisions, and source-bound in-place image resizing.

### Pagination sections

`pagination_sections` requires separate stable locators for `toc_start` and `body_start`. Approval also records decimal numbering, `start_at={"toc":1,"body":1}`, continuation after the body start, odd outer-right/even outer-left placement, and visible first-page numbers.

Application inserts a real next-page section before the body when needed. It starts the TOC and body independently at visible 1, removes later unapproved restarts, disables first-page hiding, enables odd/even footers, and ensures exactly one editable `PAGE` field in each default and even page-number footer. With approved mirror margins, the Microsoft Word adapter uses odd/even physical section starts to avoid parity blanks. When the physical start must remain even, it writes the editable calculated field `{ = { PAGE } - 1 }`; after updating the TOC, it corrects affected `PAGEREF` results and locks the delivered TOC cache. A subsequent skill run unlocks and rebuilds those fields. The caller does not perform a manual field update and must rerun the skill after pagination-changing edits. Reapplication is idempotent: a page-only footer with duplicate fields is canonicalized to one field. A footer containing publisher text, a logo, another field, or mixed payload blocks automatic replacement. Static digits may be converted only when approved trailing-section evidence marks them as derived footer-only content. Audit fails for missing even footers, duplicate PAGE fields, hidden first-page numbers, body restarts, an unapproved page formula, or unreachable header/footer parts. Physical page counts still require target-software repagination or PDF export.

Classify rendered blank pages as `intentional_recto_blank`, `removable_trailing_blank`, or `unexpected_blank`. When mirror margins and consecutive title/TOC/body content are both approved, a parity blank at either approved boundary is unexpected and the Word adapter must normalize the physical section start before delivery. A trailing blank is deleted only through approved stable cleanup. Header/footer integrity compares canonical content rather than physical part names so Word's harmless relationship renumbering does not mask a payload change.

### Table visuals

Each table candidate reports complex merges, floating objects, and visible control-mark candidates. Classify it as a data table, figure-panel layout table, pagination-only layout table, callout, or unknown. A data-table `visual` block remains unapproved until the caller confirms every column role (`numeric`, `unit`, `short_code`, or `narrative`), preferred widths, cell margins, border preset, header and semantic-separator rows, autofit, and orientation. A figure-panel layout requires explicit image and label rows.

- Only `kind=data` may receive visual formatting.
- `border_preset` is `preserve`, `three_line`, `full_grid`, `technical_textbook`, or `borderless` for an approved figure panel.
- `technical_textbook` clears old cell-level border overrides, then uses configurable major/minor widths, optional inside vertical rules, and exact `horizontal_rule_rows` for summary or group boundaries.
- Landscape requires `orientation=landscape` and `landscape_approved=true`; application creates real sections before and after the table and keeps body numbering continuous.
- Complex merges, floating objects, unknown roles, and visible control marks stay report-only until individually resolved.
- Formatting never changes cell text, merge relationships, row/column counts, or media payloads. The sole whitespace exception is an exact approved `table_cell_cleanups` entry that removes leading empty paragraphs and preserves every nonempty paragraph.

### Stable cleanup

Trailing-section candidates add a boundary locator and section-properties hash. Delete approved empty tail sections from the end inward before physical TOC migration, then remove only header/footer relationships that are unreachable from every retained section. Legacy empty TOC-anchor normalization is limited to schema 1.0 through 1.2; a legitimate separator in 1.3 or later remains authored structure.

## Schema 1.3

Schema 1.3 keeps all 1.2 behavior and adds stable locators and pagination-only groups.

- Body-paragraph locators include `text_sha256`, plus hashes of the nearest preceding and following nonempty paragraphs when available. After TOC expansion or another approved structural shift, resolve by hash and context; block when the result is absent or ambiguous.
- Heading candidates include `normalized_text_sha256` so audit can identify the approved title after a verified manual prefix is removed.
- `pagination_groups` bind an image paragraph to its following caption, or a standalone table caption to its following table, without changing text.
- An approved layout table is permitted only with `pagination_only=true`; formatting rules for data tables still cannot target it.
- `repeat_caption_with_header=true` requires caption row 0 and header row 1 and repeats both rows. `keep_rows_together` applies only pagination control.
- Approved non-heading roles clear inherited direct outline levels and accidental Heading 1-4 paragraph styles before their approved role style is applied.

New TOC migration physically removes verified static TOC paragraphs and inserts one complex Word `TOC` field. Readers normalize legacy emptied TOC anchors for audit compatibility, but new output must not leave them behind.

## Schema 1.2

### Paragraph roles

`paragraph_roles` contains no authored text. Each entry records a role, source style, direct-format signature, SHA-256, approval state, and one typed locator:

- `body_paragraph`: body paragraph index.
- `table_cell_paragraph`: table, row, cell, and paragraph indexes.

Supported roles include title, body, Heading 1-4, figure/table/equation captions, long quote, reference entry, answer, and teaching callout. `unknown` entries cannot be approved. A profile's semantic selector affects only approved matching entries; do not treat every `Normal` paragraph as body text.

### Numbering

The `numbering` block records single-chapter or whole-book mode, `chapter_start`, heading depth, expected progression, anomalies, and approval. Parse and approve the existing manual prefix before removing it. A non-first chapter uses a real level-zero `startOverride`; never restart it at 1 by default.

### Captions and tables

Captions may use either locator. Each candidate records `numbering_mode`, `identifier_semantics`, `domain_context`, `domain_confidence`, and an action. The default is `numbering_mode=manual_text` with `action=preserve`; approval never implies renumbering or `SEQ` conversion.

Caption actions are:

- `preserve`: make no structural or text change.
- `style_only`: apply the approved caption style through its paragraph role; preserve text and location.
- `replace_identifier`: replace only the hashed identifier span after individual caller confirmation; preserve the label, separator, and title exactly.
- `convert_to_seq`: require explicit approval in both the profile and map, plus an unambiguous legacy boundary.
- `move_caption`: separately approve moving a single merged caption row; it does not imply renumbering.

In architecture, civil-engineering, structural-engineering, and drafting contexts, a hyphenated number may identify a section, elevation, node, detail, or drawing callout. Analyze nearby semantic evidence and set `identifier_semantics` to `drawing_mark`, `mixed`, or `unknown` as appropriate. Punctuation alone never proves a missing sequence number. Keep an uncertain candidate report-only.

An in-cell table caption remains in its first row by default. Move it only with `action=move_caption` and `migrate_outside_table=true`; block the move when the row contains other authored content.

Each table has `kind`: `data`, `layout`, `callout`, or `unknown`. Only approved `data` tables receive general table rules. Use `caption_row`, `header_rows`, `repeat_header_rows`, `horizontal_rule_rows`, and `prevent_normal_row_split` to authorize exact rows. A layout table normally remains unchanged; the sole visual exception is an explicitly approved `layout_purpose=figure_panel`, with mapped image and label rows, `borderless`, centered cells, inline/no-wrap placement, and caption-styled labels.

Every table and image records `position_policy=preserve_anchor`. Application must not move either object to another paragraph, cell, table, section, or position among body children. Approved in-place image resizing changes only the existing inline drawing extents and paragraph alignment; it preserves the relationship, media payload, crop state, aspect ratio, drawing ordinal, container, table position, and surrounding authored objects.

Each `images` entry records a stable paragraph locator, drawing ordinal, placement class, media hash, source extent, crop/object state, raster metadata, and resize policy. Placement classes are `standalone`, `table_figure_panel`, and `table_embedded_unknown`. Only an uncropped inline standalone image or an image in an approved stationary figure-panel row may be approved automatically. Standalone bounds are 90% text width, 100% for aspect ratio at least 1.6, and 65% usable page height. Figure-panel bounds are 95% of the existing cell width, with a common displayed height within one row. Raster enlargement is capped at 125% and must retain at least 220 effective DPI; unknown or insufficient DPI prevents enlargement. Vector images may fit the approved bounds. Cropped, floating, missing-relationship, or otherwise ambiguous images remain unchanged and report-only.

A short centered paragraph directly following a standalone image may be proposed as `figure_caption_unnumbered`. A short text row directly beneath mapped image cells may be proposed as `figure_panel_label`. Both remain unapproved candidates until the caller confirms the role; approval changes style only and never inserts a number or edits wording.

For schema 1.4, `front_matter` may approve one hashed whole-book title locator, a separate unnumbered and vertically centered title-page section, an optional `book_title_format`, and insertion of the derived `目    录` heading (four ASCII spaces) before the main TOC. The technical-textbook fallback title uses 22 pt bold text with at least 33 pt line spacing and zero paragraph spacing so wrapped glyphs are not clipped. Legacy approved maps using `目录` remain readable. The `TOC` field does not create its own heading; the skill inserts and maintains the separate derived paragraph. `block_spacing` may approve one real empty paragraph after each approved data table and complete approved figure block. These spacer paragraphs are derived structure and must be removed by target-software pagination when they would start a new page.

### Trailing sections

Candidates report visible payload, header/footer references and payload, page-number start, first-page behavior, section type, page geometry, stable boundary context, and section-properties hash. Delete only an approved candidate whose evidence says `safe_to_delete`. Independent or ambiguous settings block deletion.

## Operations

- `toc_ranges`: replace an approved static directory range with a real `TOC` field.
- `headings`: assign approved body paragraphs to Heading 1-4 and remove only verified prefixes.
- `captions`: preserve or style manual identifiers by default; perform only explicitly approved identifier replacement, relocation, or field conversion.
- `tables`: apply exact approved header, row-split, visual, and landscape controls to data tables.
- `images`: resize only individually approved inline images inside their original anchors and containers.
- `pagination_groups`: apply only approved `keepNext`/`cantSplit` relationships to figures, captions, and tables.
- `pagination_sections`: create and audit approved TOC/body numbering sections and odd/even PAGE footers.
- `front_matter`: separate the whole-book title page from the TOC and insert the approved TOC heading.
- `block_spacing`: insert idempotent same-page-only empty paragraphs after approved figure/table blocks.
- `trailing_empty_sections`: remove safe approved final sections from the end inward.

Style rules clear conflicting direct formatting only for the properties they control and only on approved role targets. Approved font rules also clear the corresponding theme-font attributes and are audited after resolving run, character-style, paragraph-style, base-style, document-default, and theme inheritance. They preserve uncontrolled color, language, character styles, superscript/subscript, hyperlinks, fields, bookmarks, comments, revisions, and formula formatting.

## Privacy and integrity

Maps contain indexes, roles, settings, and hashes, never manuscript text. Keep unpublished DOCX files, detailed inventories, rendered pages, and task reports outside the skill repository.

Application verifies the source fingerprint and approved target hashes. Audit compares effective style plus paragraph/run formatting, normalizes only approved derived-field changes, and keeps all authored text logically significant. A manual identifier replacement receives a separate audit entry proving that the exact approved identifier changed and the caption title did not.
