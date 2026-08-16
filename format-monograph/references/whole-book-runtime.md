# Whole-book runtime

`run_monograph.py` is the portable entry point for resumable whole-book work. It
orchestrates the existing inspection, application, audit, field-finalization,
and rendering scripts; it does not replace their safety checks.

## Local state

The work directory contains `run-state.json`, inventories, a candidate map,
derivatives, audits, and render evidence. The state stores fingerprints,
capabilities, statuses, artifact paths, grouped QA, frozen scopes, cache hits,
and stage durations. It must not contain manuscript paragraphs or extracted
source text.

Each completed stage has an input-key hash. `--resume` reuses the stage only when
its key still matches and all required artifacts exist. A changed manuscript or
profile starts a new run identity. A changed approved map reruns application; a
changed field backend reruns finalization; a changed target PDF or visual
manifest reruns verification.

## Prepare

`prepare` validates the profile, captures environment capabilities, inventories
the whole DOCX, and creates a source-bound schema 1.5 structure map. It groups
repeated decisions, records local frozen scopes, and creates a bounded trial
selection manifest without copying manuscript text into the map.

## Apply

The caller reviews the candidate map, records group decisions and exceptions,
approves only known targets, and sets the map status to `approved`. `apply`
validates that map against the unchanged source, formats the entire approved
scope, creates clean/review/report artifacts, and runs the content audit.
Unresolved frozen objects remain unchanged while approved independent scopes may
continue. Their presence keeps the run at `blocked_qa`.

## Finalize

`finalize` requires all critical grouped QA and frozen scopes to be closed. It
updates approved fields using an external target-application backend or
LibreOffice UNO. The backend writes a disposable copy. The core parses field
boundaries, matches main-story fields by semantic instruction and header/footer
fields by effective section role, and
patches only scalar result text or the validated TOC result span into the safe
baseline package. Whole backend XML parts are never imported. Backend
run splitting, simple-to-complex field conversion, and header/footer part
renumbering are reported and discarded; changed authored text, field
instructions, approved bookmarks, pagination sections, media, equations, or
embeddings rejects the refresh.

Repeated identical field instructions must have unique retained paragraph IDs
or unique authored paragraph context; otherwise finalization blocks instead of
pairing them by occurrence order. Every refreshed approved field must be clean,
and global update-on-open is removed even when clean unapproved fields remain.
The only allowed formula has one directly nested `PAGE` field and the exact
core-generated `= PAGE - 1` structure; other formulas and nonnumeric page
results are rejected.

For a target Word backend, section and footer structure must already pass the
portable core audit. Word first performs a no-save `measure_layout` pass that
returns only page counts, section page ranges, and approved spacer ordinals. The
core iteratively resolves parity section starts and removes only approved empty
spacers that landed at a page top. When an approved page-1 restart must begin on
an even physical page, the core creates the editable `{ = { PAGE } - 1 }`
display field in newly isolated odd/even footer parts so adjacent sections keep
their own `PAGE` fields; no other formula field receives approval. The adapter
must open measurement and verification inputs read-only, report a successful
update for every approved field, and leave the independently checked input hash
unchanged. After selective
writeback, Word reopens the selective DOCX,
repaginates without saving, and exports the PDF used for visual QA. Page count
must match the field-calculation session. Only `selective_verified` satisfies
the completed field gate.

Use `--approve-deferred` only after explicit caller QA. Deferred-on-open remains
`candidate_ready`; it cannot become `final_ready`.

## Verify

`verify` repeats the content audit and renders every page. A local visual-QA
manifest must contain:

```json
{
  "all_pages_inspected": true,
  "target_layout_verified": true,
  "page_count": 120,
  "issues": []
}
```

The page count must equal the render manifest. Any missing page, unresolved
issue, non-target layout, deferred field result, integrity failure, or open QA
prevents `final_ready`.

## Recovery

1. Run `status --json` and inspect the failed or blocked stage.
2. Correct the source-independent input that caused the problem, such as a QA
   decision, approved map, renderer path, or backend authorization.
3. Rerun that command with `--resume`.
4. Rerun later stages whenever an upstream fingerprint changes.

Do not delete the source or overwrite it during recovery. Do not claim success
from an older cached artifact after its input key changes.
