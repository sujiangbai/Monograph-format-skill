# Microsoft Word Windows adapter

This optional adapter uses the target Microsoft Word layout engine to calculate
editable field results and to perform a separate no-save verification render. It
does not own pagination structure and never provides the delivery parent DOCX.

## Requirements

- Windows with desktop Microsoft Word installed and activated.
- PowerShell 5.1 or newer.
- Permission from the caller to automate Microsoft Word on a derivative copy.

Run the core finalizer with an argument-array command to avoid shell parsing:

```text
python format-monograph/scripts/finalize_docx.py formatted.docx \
  --source original.docx \
  --profile approved-profile.json \
  --structure-map approved-map.json \
  --output finalized.docx \
  --field-updater external \
  --field-updater-command '["powershell.exe","-NoProfile","-ExecutionPolicy","Bypass","-File","<adapter-path>/word_field_updater.ps1"]' \
  --target-software "Microsoft Word 2021" \
  --pdf-output finalized-word.pdf
```

Always pass the updater as a JSON argument array, especially when paths contain spaces. Measurement, refresh verification, and final verification open their inputs explicitly read-only; only the disposable refresh copy is opened writable. Optional Word preference assignments are best-effort and do not weaken macro blocking, Protected View, or organization policy.

The core first establishes and audits all sections, page-number starts, odd/even
footers, and field instructions. The adapter first opens the baseline without
saving and returns page/section measurements plus approved page-top spacer
ordinals. The core owns any resulting section, spacer, footer, or page-display
field change. The adapter then opens only a temporary copy,
disables macros, link updates, prompts, and visible windows, and updates approved
`TOC`, `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `REF`, and `PAGEREF` fields across Word
story ranges. It may also update the core-generated `{ = { PAGE } - 1 }` formula
when the approved mirrored-page restart requires it; arbitrary formula fields are
rejected by the core. The core selectively imports verified field results into its audited
baseline and discards the Word-saved OOXML package. Finally, the adapter opens the
selective output without saving it, repaginates, exports the target PDF, and reports
the page and field counts used by the verification gate. It does not request
recent-file recording for either operation.

The core independently checks that read-only operations did not change the
input hash and rejects a missing or mismatched verification page count. The
adapter treats a false `Field.Update()` return as failure. A successful process
exit or an unchanged field count alone is not cache verification.

Every `measure_layout`, `refresh_fields`, and `verify_only` response must report
`structural_changes_applied: 0`. The adapter must not create or
delete sections, rewrite headers or footers, remove spacer paragraphs, correct
page-number offsets, or change field instructions. Those operations belong to the
portable core and require the approved structure map.

Do not use this adapter on an untrusted DOCX outside a sandboxed user profile. Protected View, organization policy, Word dialogs, damaged fields, or automation denial must be reported as a blocked or deferred finalization; never bypass those controls.
