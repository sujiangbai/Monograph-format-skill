# Structure map

Use a source-bound structure map for operations that reinterpret document structure. Profile approval authorizes formatting rules; structure-map approval authorizes specific targets in one unchanged DOCX.

## Workflow

1. Generate an inventory and candidate map with `inspect_docx.py --structure-map-output`.
2. Review candidates with the caller. Approve only unambiguous targets and set top-level `status` to `approved`.
3. Run `validate_structure_map.py <map.json> --source <input.docx>` immediately before applying it.
4. Pass the same map to `apply_profile.py` and `audit_docx.py`.
5. Regenerate the map and repeat QA whenever the source fingerprint changes.

New maps use schema `1.3`; readers continue to accept `1.0`, `1.1`, and `1.2`. Version 1.0 can authorize its original TOC, heading, caption, first-row table, and trailing-section operations. Version 1.1 retains its explicit legacy `SEQ` conversion behavior. Version 1.2 adds domain-aware manual caption actions and semantic paragraph roles.

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

Each table has `kind`: `data`, `layout`, `callout`, or `unknown`. Only approved `data` tables receive table rules. Use `caption_row`, `header_rows`, `repeat_header_rows`, and `prevent_normal_row_split` to authorize exact rows. Layout tables, image containers, teaching boxes, and unknown tables remain unchanged.

### Trailing sections

Candidates report visible payload, header/footer references and payload, page-number start, first-page behavior, section type, and page geometry. Delete only an approved candidate whose evidence says `safe_to_delete`. Independent or ambiguous settings block deletion.

## Operations

- `toc_ranges`: replace an approved static directory range with a real `TOC` field.
- `headings`: assign approved body paragraphs to Heading 1-4 and remove only verified prefixes.
- `captions`: preserve or style manual identifiers by default; perform only explicitly approved identifier replacement, relocation, or field conversion.
- `tables`: apply exact approved header and row-split controls to data tables.
- `pagination_groups`: apply only approved `keepNext`/`cantSplit` relationships to figures, captions, and tables.
- `trailing_empty_sections`: remove safe approved final sections from the end inward.

Style rules clear conflicting direct formatting only for the properties they control and only on approved role targets. They preserve uncontrolled color, language, character styles, superscript/subscript, hyperlinks, fields, bookmarks, comments, and revisions.

## Privacy and integrity

Maps contain indexes, roles, settings, and hashes, never manuscript text. Keep unpublished DOCX files, detailed inventories, rendered pages, and task reports outside the skill repository.

Application verifies the source fingerprint and approved target hashes. Audit compares effective style plus paragraph/run formatting, normalizes only approved derived-field changes, and keeps all authored text logically significant. A manual identifier replacement receives a separate audit entry proving that the exact approved identifier changed and the caption title did not.
