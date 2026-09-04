# Mac Word synthetic capability probe

This directory contains opt-in experiments for one generated V0.5.1 fixture. It
is not a production updater, does not implement the external updater protocol,
and cannot produce product `final_ready`. Do not configure any script here as a
production field updater.

## Offline checks

Normal tests do not start Word:

```sh
python -m unittest discover -s tests -p 'test_v051_*.py' -v
```

The fixture generator uses only literal synthetic text/table data and an
in-memory procedural PNG. Its package audit is fixture-specific, not a general
untrusted-DOCX sanitizer. Use an existing runtime with `python-docx`, `lxml`,
Pillow and `pypdf`; do not install or alter system dependencies for this probe.

## Explicit opt-in live entry points

These commands require fresh, specific user/manager authorization, no user
documents or printing activity, and new evidence paths. Do not rerun retained
historical paths automatically.

```sh
osascript adapters/microsoft-word/macos/safe_open_controls.applescript
python tests/v051_macos_word_fixture.py --out-dir <new-synthetic-dir>
osascript adapters/microsoft-word/macos/readonly_settings_probe.applescript \
  "$PWD/<new-synthetic-dir>/v051-synthetic.docx"
python adapters/microsoft-word/macos/run_field_probe.py \
  "$PWD/<new-synthetic-dir>/v051-synthetic.docx" \
  --output-dir "$PWD/<new-evidence-dir>"
python adapters/microsoft-word/macos/run_readonly_pdf_probe.py \
  --output-dir "$PWD/artifacts/probe-evidence/<new-readonly-pdf-dir>"
```

The last entry is intentionally pinned to the accepted live-05 candidate and to
the exact live-05 calculation JSON path, non-symlink entity, 7269-byte size and
SHA-256 `70eb8c837e082ab260b520fef7736b84c001e5508e2973cbe09601d1940761b0`.
It refuses changed bytes, forged convergence, symlinks or path replacement
before starting Word.

## Safety and result boundary

The live scripts refuse existing documents/printing jobs, temporarily tighten
AutomationSecurity and UpdateLinksAtOpen with readback/restoration, and claim a
document only by one exact case/diacritic/punctuation/whitespace-sensitive POSIX
path. They never use `active document` or a basename for ownership. Cleanup may
close only that unique exact-path document without saving; ambiguity or failed
restoration remains blocked. The scripts never quit or kill Word, touch Normal,
click GUI permissions, or lower macro/link protection.

The field prototype updates only the existing approved TOC and six approved
scalar fields on an editable calculation copy. It never updates the unapproved
QUOTE field or uses a collection-wide refresh. It requires exact field
instructions/ranges, approved headings/bookmarks, complete pagination tuples,
two consecutive equal observations, fixture content/position checks, and the
unchanged core selective writeback before creating a candidate. The source is
never edited.

The normal read-only/PDF entry opens only the fixed candidate read-only, requires
Boolean `saved=true` before export, Boolean read-only and false printing-update
settings, and performs no field refresh, repagination or DOCX save. Post-export
`saved` is recorded as a Boolean observation, not used alone as a safety verdict.
Both complete snapshots must be identical and match the pinned calculation;
exact no-save close/restoration, unchanged source/candidate/calculation entities,
and PDF path/hash/size/page count remain hard gates. Visual QA is still separate.

## Current bounded evidence

On Word 16.112.3, live-05 reached two equal complete tuples and passed the
fixture content checks plus the unmodified strict selective writeback. The
observed result was 3 pages, TOC entries on pages 1/2, body physical/logical start
2/1, lower-Roman then decimal sections starting at 1, seven approved writebacks,
and unchanged unapproved QUOTE. This is local synthetic capability evidence, not
a production-ready adapter.

Two later PDF observations remain real blockers:

- `field-010-readonly-pdf-02`: Word returned from PDF export, then Boolean
  `saved=false`; the unchanged normal gate rejected it before the after snapshot.
  Its retained 3-page PDF is a failed-run artifact, not accepted verification.
- `field-010-pdf-state-diagnostic-01`: the diagnostic recorded two pre-export
  state points and a complete pre-snapshot read, but the PDF call timed out
  (`-1712`) before returning. Post-export state, after snapshot and a new PDF are
  missing. A later read-only session check found documents=0 and printing=0.

In a later manual GUI control, Word showed the same save prompt after exporting a
copy of the same candidate; the user chose not to save and the final DOCX SHA-256
remained `338ffb3dedc293f516bf28e088cdd6f0b864bb324254412e50e904ea33795ff6`.
This shows only that post-export `saved=false` is not sufficient evidence of a
disk DOCX change. It does not prove the PDF or in-memory state correct. Unknown
saved type, lost read-only state, snapshot drift, uncertain cleanup/restoration,
or any entity change still fails. The whole batch remains `final_ready=false`.

## Evidence index

Raw synthetic evidence is retained under ignored `artifacts/probe-evidence/`:

- `safe-open-004/` - fixed source fixture and package manifest.
- `field-010-live-05/` - accepted calculation, convergence and strict writeback.
- `field-010-readonly-pdf-01/` and `field-010-readonly-pdf-02/` - normal-entry
  failures and the retained failed-run PDF.
- `field-010-readonly-pdf-clean-retry-01/` - clean-restart reproduction of the
  post-export dirty-state observation.
- `field-010-pdf-state-diagnostic-01/` - bounded diagnostic, cleanup history and
  final read-only session postcheck.
- `interface-probe-010-latest-handoff.json` - current machine-readable summary.

The complete chronology, including failed intermediate probes, remains in
`artifacts/manager/v051-macos-word-probe-executor-report.md`; it is deliberately
not duplicated here.
