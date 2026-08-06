# Structure map

Use a source-bound structure map for operations that reinterpret existing document structure. The map is separate from the format profile: profile approval authorizes formatting rules, while structure-map approval authorizes specific changes to a specific DOCX.

## Workflow

1. Generate an inventory and candidate map with `inspect_docx.py --structure-map-output`.
2. Review candidates with the caller. Approve only unambiguous targets.
3. Set top-level `status` to `approved`; leave every rejected or unresolved item unapproved.
4. Run `validate_structure_map.py <map.json> --source <input.docx>` immediately before applying it.
5. Pass the same map to both `apply_profile.py` and `audit_docx.py`.

Never reuse a map after the source fingerprint changes. Regenerate and repeat QA instead.

## Operations

- `toc_ranges`: replace approved static directory paragraphs with a real Word `TOC` field.
- `headings`: map approved paragraphs to `Heading 1` through `Heading 4` and remove only a verified manual prefix.
- `captions`: replace only the verified numbering prefix with `STYLEREF` and chapter-restarting `SEQ` fields; preserve caption wording.
- `tables`: repeat the approved first row and prevent normal rows from splitting across pages.
- `trailing_empty_sections`: remove approved empty final sections from the end inward. Block deletion when the section has content or independent header, footer, or page-number settings.

Do not approve broken cross-references, ambiguous captions, questionable table headers, or nonempty section boundaries. Report them for QA.

## Privacy and integrity

Candidate maps contain indexes, roles, settings, and SHA-256 hashes, not authored paragraph or table text. Keep unpublished DOCX files, inventories that contain manuscript details, rendered pages, and task-specific reports outside the skill repository.

Applying a map must verify the source fingerprint and each approved target hash. Auditing must normalize only the explicitly approved derived-field changes and deleted empty section boundaries; all authored text remains byte-for-byte significant at the logical-content level.
