# Microsoft Word Windows adapter

This optional adapter refreshes editable Word fields with the target Microsoft Word layout engine. It does not change the portable Skill core.

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

The adapter opens only the copied output, disables macros, link updates, recent-file recording, prompts, and visible windows, then updates approved `TOC`, `PAGE`, `REF`, and `PAGEREF` fields across Word story ranges. It repaginates, saves DOCX, reopens it to verify editable fields, optionally exports PDF, and always closes its Word instance.

Do not use this adapter on an untrusted DOCX outside a sandboxed user profile. Protected View, organization policy, Word dialogs, damaged fields, or automation denial must be reported as a blocked or deferred finalization; never bypass those controls.
