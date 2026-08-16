# Portable run checklist

Use this checklist unchanged in Codex, Claude Code, Gemini CLI, VS Code with
GitHub Copilot, and any other Agent Skills-compatible environment. Platform
adapters may locate the skill and provide an approved field or render backend;
they must not change authority, safety, QA, or delivery rules.

## Before preparation

1. Resolve all skill files relative to `SKILL.md`.
2. Keep the manuscript, evidence, work directory, rendered pages, and reports
   outside the installed skill and outside a public repository.
3. Confirm the caller's current requirements and the selected approved profile.
4. Run `<python> scripts/check_environment.py --json` and retain the complete
   capability snapshot. Multimodal source reading is Agent-declared and must not
   be inferred from Python or DOCX support.
5. Select the capability mode from the reported facts. Do not promote a run
   because another Agent completed a similar task.

## Standard whole-book run

```text
<python> scripts/run_monograph.py prepare <input.docx> --profile <profile.json> --work-dir <directory>
<python> scripts/run_monograph.py status --work-dir <directory> --json
<python> scripts/run_monograph.py apply --work-dir <directory> --structure-map <approved.json>
<python> scripts/run_monograph.py finalize --work-dir <directory>
<python> scripts/run_monograph.py verify --work-dir <directory> --visual-qa-manifest <visual-qa.json>
```

Add `--resume` when retrying the same stage. A source, profile, structure-map,
field-backend, target-PDF, or visual-manifest fingerprint change invalidates the
affected cache. Never copy a `run-state.json` between manuscripts.

## QA gate

- Review grouped questions once, then record object-level exceptions.
- Keep uncertain appendices, heading numbering, TOC scope, tables, and images in
  `frozen_scopes`. Apply may continue for approved objects outside those scopes.
- Preserve existing appendix, figure, table, and equation identifiers unless an
  individual change is explicitly approved.
- Require the book-title section boundary to end on the title paragraph so an
  auxiliary empty paragraph cannot shift vertical centering. Require approved
  Heading 1-4 paragraphs and their numbering levels to pass zero-indent and
  direct-numbering audits before delivery.
- Never approve a table header only because it is the first row.
- Never move an image or table. A floating table remains frozen when changing it
  to inline/no-wrap would alter its anchor or position.
- Representative trial selections cover front matter, heading levels,
  appendices, and one or two examples per figure/table class. One candidate is
  limited to 30 rendered pages and is not whole-book pagination evidence.

## Finalization and delivery gate

- Target-application updates occur in a disposable copy. Parse and uniquely
  match each approved field, then patch only its result into the core-generated
  baseline. Never replace an entire backend XML part.
- Establish sections, visible page-number rules, headers, footers, and field
  instructions before calling Word. Require the adapter to report zero
  structural changes.
- Reopen the selective output in the target application without saving, export
  the verification PDF, and require the same page count as the calculation run.
- `deferred`, `stale`, and `code_only` are not completed field refresh states.
- Inspect every page rendered by the target application. Record the exact page
  count and unresolved issues in a local visual-QA manifest.
- Set `final_ready` only when critical QA is closed, field results are safely
  refreshed, integrity audits pass, and the manifest confirms every page was
  inspected with no unresolved issue.
- Retain `run-state.json` and the generated audit JSON as execution evidence.
  Agent prose or a DOCX that merely opens is not evidence of completion.
- Do not place manuscript text, local task artifacts, or rendered pages in
  public logs, repositories, CI artifacts, or compatibility fixtures.

## State meanings

- `analysis_only`: the environment cannot safely modify the DOCX.
- `prepared`: inventory and candidate structure map are ready for QA.
- `blocked_qa`: one or more decisions or protected scopes remain unresolved.
- `candidate_ready`: a derivative exists but a final delivery gate remains.
- `final_ready`: every required integrity, field, and visual gate passed.
- `failed`: an execution or integrity failure requires investigation or repair.
