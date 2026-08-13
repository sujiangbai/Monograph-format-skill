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
LibreOffice UNO. The backend writes a disposable copy. The core compares OOXML
with cached field-result text neutralized, then writes only verified field-part
results into the safe baseline package. It never imports refreshed media or
embeddings. Any non-field authored XML change is rejected.

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
