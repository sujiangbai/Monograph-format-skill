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
its key still matches, all required artifacts exist, and the current output
evidence revalidates. Finalization resume rereads the finalization JSON and
rehashes the finalized DOCX and persistent Word verification PDF. Verification
resume also rereads and rehashes the final audit, render manifest, and visual
manifest, and repeats the completion, artifact-binding, and page-count checks.
`status` performs the same read-only checks before reporting an existing
`final_ready`; invalid evidence is atomically downgraded to `candidate_ready`
without regenerating an artifact. A changed manuscript or profile starts a new
run identity. A changed approved map reruns application; a changed field backend
reruns finalization; changed or contradictory finalization or verification
outputs reject cache reuse rather than preserving an older completion claim.

Every real successful finalization execution invalidates the downstream verify
stage, verification-output bindings, audit/render/visual artifact references,
stored visual-QA result, and rendered-page metric before state is saved, even
when the newly finalized DOCX and target PDF are byte-identical to their prior
versions. The next `verify --resume` must rerun audit, render, and visual-manifest
validation and rebuild those bindings; it never reruns finalization or the
updater. A genuine finalization cache hit preserves the already revalidated
verification evidence and `final_ready` state.

Here, a real successful execution means both a zero finalizer process result and
a readable `finalization_evidence_version=1` object that passes the shared closed
shape validator. The validator checks required top-level and nested objects,
field-cache/backend/writeback/completion/target enums, boolean and integer types,
artifact-binding version and identities, non-empty paths, lowercase SHA-256
values, workflow stage, integrity fields, and stored evidence-validation shape.
Malformed JSON, an unknown version, a missing field, or any shape error returns
non-success without saving or clearing the old verification state. Shape validity
does not grant completion: valid deferred, LibreOffice non-final, and other
candidate evidence still pass through the existing consistency and final-ready
gates.

Finalization stores a versioned, allowlisted gate summary in run state. It binds
the finalization result status, content/object/font integrity results, finalized
workflow stage, source/formatted/profile/approved-map/output SHA-256 values,
output and target-PDF paths, target-layout status, artifact identities, and the
complete field shape. Resume, verification, and status rebuild this summary
from the current finalization JSON, compare it field by field with the state
copy, and hash the current entities. A completion-shaped field payload cannot
override a failed integrity result or a conflicting workflow or target field.
This is explicit local consistency checking rather than a hash of the whole
JSON and is not a signature.

Finalization and verification input keys also bind versioned request identities.
Target software is one allowlisted ID: `microsoft_word`, `libreoffice`, or
`unsupported`; only explicit Microsoft Word/Microsoft 365 aliases map to Word.
An omitted finalization target inherits the approved profile default, while an
omitted verification target inherits the persisted finalization ID. Only an
actually invoked renderer is bound by request source, resolved path/hash/size.
An external updater identity is a versioned audit record, not a proof of a
runtime dependency closure. The shared parser and executor use the same argv,
PATH resolution, and `format-monograph` execution directory. The identity records
fixed arguments; the lexical and resolved executable path, hash, size, and
symlink destination; plain file arguments; `--key=file` and `--key file` values;
and `@response-file` bytes. Directory arguments use a sorted recursive manifest
of relative path, entry type, file hash, and size. Directory symlinks, including
symlinks below the root, are recorded but never followed; explicit file symlinks
record both lexical and resolved targets. Python script/module inspection may add
diagnostic facts, but it never authorizes cache reuse.

Without an OS-enforced, auditable hermetic runtime, a native program, wrapper,
direct Python script, `-m` module/package, `/usr/bin/env`, or other external
program can still depend on cwd files, environment variables, PATH subcommands,
site customization, reflection, dynamic imports, or unobserved resources. Every
external identity therefore records `dependency_closure.status=unproven`,
`external_program_not_hermetic`, `runtime_dependencies_unproven`, and
`cache_reusable=false`. Every `finalize --resume` with an external updater runs
the finalizer/updater again, even when argv and all recorded entities are
unchanged. Fresh, otherwise valid evidence is not failed merely because it is
non-cacheable, and `status` remains read-only. Recorded entity changes still
invalidate local consistency claims for existing evidence.

This is unsigned argv/explicit-entity consistency, not a signature,
external-application authentication, or arbitrary-program dependency analysis.
External finalize caching may be reopened only by a separately approved adapter
whose hermetic runtime is enforced and auditable. An unused renderer is neither
probed nor bound. Verification caching is separate and remains bound to the
actual finalized artifacts, target PDF, renderer, and visual request.
Verification tests the new input key before consulting old output bindings, so
a new valid visual manifest path or hash runs audit/render again and replaces
the bindings. Only a true key match can reuse and revalidate cached outputs.

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
updates approved fields using an external target-application backend or the
verified macOS LibreOffice internal-Python macro host. The legacy UNO
server/helper path is disabled. Linux, Windows, unsupported executables, and
ordinary wrappers fail closed; `auto --approve-deferred` may record deferred
without starting that backend. The backend writes a disposable copy. The core parses field
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
must match the field-calculation session and the later render/visual manifests.
The only complete Word cache transition is input `stale` to output `refreshed`;
`absent`, deferred/stale output, and unknown states are rejected. The Word
verification PDF must be a persistent artifact. Finalization records resolved
path, SHA-256, and size for both the finalized DOCX and verification PDF, plus
the PDF page count. Verification rehashes both files and cross-checks the
finalization JSON against its run-state copy. This is unsigned local consistency
and artifact binding, not an unforgeability guarantee; a subject able to modify
all local artifacts and both JSON files is inside the explicit local trust
boundary. Only `selective_verified` satisfies the completed field gate.

Finalization separates gate evidence from diagnostic evidence. The inline,
versioned `field_backend` object is a closed canonical projection: consumers
reject unknown fields, unsupported operations or statuses, invalid counts or
types, and NaN/Infinity at every nested level. The unabridged external or
LibreOffice response is never copied inline. It is serialized as a separate,
versioned backend-audit JSON sidecar with standard-JSON value validation and
fixed byte, nesting-depth, node-count, and string-length limits. Finalization
atomically persists that sidecar before its status JSON and records only the
sidecar path, SHA-256, and size in the latter. Resume, verify, and status bind
the state artifact to that identity and rehash/reparse the sidecar. Missing or
changed diagnostics safely invalidate local final-ready evidence, while
completion and gate decisions continue to use only the canonical projection.
Without a status-output path the result explicitly says the audit was not
persisted and does not inline raw diagnostics.

The finalizer resolves its complete path graph before creating a directory,
unlinking, or writing: input/source/profile/map plus final DOCX, optional PDF,
status, and the status-derived backend-audit sidecar. Persistent outputs are
lexically and resolution-wise unique, do not alias any input or symlink, and
share one parent directory. On POSIX, an already-open parent authority creates
both staging and its producer child through one create-and-bind helper: a random
basename is exclusively made with `mkdir(dir_fd=...)`, immediately no-follow
statted, then parent-relative opened and compared by device/inode/owner/mode.
Collisions retry a fresh name; create-to-open substitutions are retained and
rejected rather than adopted. Every hook and producer action occurs only after
the relevant helper returns. Final staged
DOCX/PDF/status/audit bytes are imported only by authority-relative exclusive,
no-follow creation; a pre-existing regular file, symlink, directory, junction,
reparse point, or unknown entry is never followed, replaced, truncated, moved,
or unlinked. Each import derives its complete identity from target `fstat` while
the write FD is still open plus the digest of bytes actually written. After
close, it immediately reopens relative to the staging authority with no-follow,
compares the full identity, and returns the original writer identity. Main passes
that identity independently to the publisher; its first snapshot must match
rather than define expected, even without a status sidecar. The producer child is retained as reported evidence. Old targets
are snapshotted but not removed during computation.
Publication atomically captures the actual current pathname into a unique backup,
verifies that captured inode against the startup snapshot, and uses only the
platform's atomic no-replace move for capture, publication, rollback quarantine,
restoration, and recovery promotion. The approved production primitives are macOS
`renameatx_np(RENAME_EXCL)` and Linux `renameat2(RENAME_NOREPLACE)`; `EEXIST`, cross-filesystem,
reparse, parent-identity, permission, or unsupported results fail closed with no hard-link, ordinary-rename, copy,
overwrite, or check-then-unlink fallback. A successful move consumes the staged
pathname. Production never deletes the staging directory. On success it verifies
that all new artifact paths were consumed, writes a versioned
`business_gate=false` `publication-record.json` exclusively, binds its complete
write-FD identity, and after hooks requires both an exact no-follow identity and
an exact staging entry set containing only the producer workspace and record. It reports transaction ID,
retained staging path, and `cleanup_policy=manual_only`. On failure it reports
and retains staging/quarantine/backup/unknown evidence. No success or failure
path performs check-then-rmdir or recursive staging deletion, so a substituted
object is never removed as if it were producer-owned. Tests delete only their
enclosing temporary roots after every authority FD/handle closes. Operational
residue reporting therefore distinguishes intentional, reported publication
evidence from unreported residue.

All transaction, recovery, producer, staging, and output-parent authority closes
are attempted independently and their diagnostics are aggregated. Once status
and the exclusive publication record establish the commit point, a close or
console-report failure is reported as structured
`committed_with_cleanup_errors`; it does not roll back committed artifacts or
return a generic pre-commit publication failure. The persisted business gate is
unchanged and remains subject to normal consumer revalidation. Stderr cleanup
diagnostics and stdout status are both best-effort after commit; broken pipes,
OS/encoding errors, and closed streams are swallowed and cannot reverse success.

The producer validates the complete finalization evidence, serializes it exactly
once as standard JSON bytes, and independently records that byte string's
SHA-256 and size. The publisher requires that external identity before doing any
work. Immediately before and after publishing status last, it reopens a regular,
non-symlink status entity, compares the exact identity, reruns the complete
closed-shape and completion/evidence-validation calculation, and checks every
DOCX/PDF/audit binding. A subset binding check is only additional defense; it is
not the status trust root. A status mutation before commit prevents a new disk
transaction, and a mutation after status publication triggers rollback and
evidence retention.

Successful replacement does not delete any captured old inode. Each backup is
atomically moved to the stable
`.format-monograph-recovery/<transaction-id>/` directory. On POSIX the publisher
opens stable no-follow directory FDs, validates effective-user ownership and
rejects group/world-writable authority directories, then performs relative
creation/opening. Windows production publication is intentionally unavailable:
SDDL hashes or trustee/right string denylists do not prove effective access
(including NULL/empty DACL, inherited, conditional, object, callback, unknown
trustee, owner, and deny/allow-order cases), and pathname `MoveFileExW` cannot
bind already verified source/root handles to the no-replace move. The program
therefore fails closed before creating a missing output parent, staging, or producer child. A future Windows
publisher requires separately approved AccessCheck-equivalent authorization and
an audited authority-bound native rename. POSIX entries are rechecked after recovery-parent
open, transaction creation, each backup move, manifest publication, and
immediately before return. Replacing a
parent, transaction, backup, or manifest entry with a symlink, file, or other
directory fails closed and never authorizes deletion of that substitute. Its
versioned manifest records the transaction ID, original target, recovery path,
and startup snapshot. The manifest writer-FD identity is retained and compared
against a full authority-relative no-follow snapshot after the manifest hook and
again before return. At both points the transaction entry set must equal the
expected `*.previous` files plus `recovery-manifest.json`; a same-byte new inode
or any extra entry is blocking and retained.
An open descriptor can continue modifying that old inode after capture, so its
later hash is explicitly mutable diagnostic evidence and never a final-ready
business gate. Recovery and staging consume storage and are never automatically
cleaned; an operator may remove them only after proving that no process retains
an open FD/handle and re-establishing current path ownership. These private
directories are controlled by the current effective user. A same-UID/SID actor
able to coordinate changes across every local artifact is inside the existing
unsigned trust boundary, although production still never actively deletes an
unknown object. This ordered protocol is not a globally atomic multi-file
transaction, offers no `fsync` or power-loss guarantee, and status is not a lock;
every consumer must still rehash and revalidate current artifacts. If rollback cannot complete, staging,
recovery, backups, quarantine, and unknown concurrent objects remain at reported
paths rather than being unconditionally deleted.

The explicit `windows-latest` gate currently proves that the real host rejects
publication before producer execution and leaves existing targets unchanged.
Darwin-backed Windows capability simulation remains supplemental coverage for
generic control flow, not production availability or Windows execution evidence.
Windows publication is a documented NO-GO until both missing OS contracts are
implemented, adversarially tested on a real runner, and separately approved.

After a fresh finalizer process exits zero and the inline shape passes, the
orchestrator independently derives the sidecar path from its own fixed status
path. Before `invalidate_verification_state`, it requires the exact binding
schema/version/status, the exact non-symlink regular path, size/hash, standard
JSON and bounded recursive shape, and the exact sidecar root version/schema.
Failure returns non-success without saving state or clearing the old verify
stage. All evidence versions require `type(version) is int` and the supported
value; `true`, `1.0`, `"1"`, and unknown integers never match.

On macOS, the only LibreOffice field-refresh path runs its Python field helper as a user macro
inside the same disposable profile used by the headless process. It passes only
temporary input, output, and result paths through that process environment,
closes the refreshed document before publishing success, and removes the
profile afterward. It does not change the caller's global macro settings or
place a macro in the delivery DOCX. The document load descriptor uses
`UpdateDocMode.NO_UPDATE` and `MacroExecMode.NEVER_EXECUTE` so external links and
document-embedded macros are not authorized during load. Before launch, the
parent inventories package relationships, external field instructions, and XML
`href`/`src` values. `TargetMode=External`, every URI scheme (at least `file`,
`ftp`, `smb`, `http`, and `https`), URI-relative or backslash network paths,
package escapes, malformed targets, and unresolved references are rejected
without starting LibreOffice. Internal relationship targets are percent-decoded,
backslashes are normalized, and relative targets are resolved from the OPC
source part (`_rels/.rels` from the package root and
`word/_rels/document.xml.rels` from `word/document.xml`), not from the `_rels`
directory. Fragments and relative or package-root members remain allowed only
when the resolved package member exists. The helper receives a hashed,
versioned authorization only when the approved structure map and baseline DOCX
yield one stable identity. It separately receives the canonical approved TOC
contract. The macro independently validates and canonicalizes that list,
recomputes its SHA-256, and requires equality with the authorization's
structure-contract hash before it examines or updates an index. The
authorization binds that hash, exact
OOXML TOC instruction and paragraph/field position, occurrence/ordinal/service,
and expected UNO `CreateFromOutline`, `CreateFromMarks`, and `Level`. The macro
reparses the input package, rereads those UNO properties, rebuilds and rehashes
the observed descriptor, and compares every value before any `update()`. A
different count, service, position, instruction, outline source, marks source,
or level rejects with zero updates. It then performs only that authorized update plus
document-internal calculation. It
never calls collection-level `TextFields.refresh()`, because that operation can
update external-data fields. The helper closes the document successfully before
it publishes `ok=true`; a close failure removes the output and publishes a
structured negative result. Prefer an explicitly configured renderer;
otherwise use the installed LibreOffice app before a wrapper that can be
strictly resolved to that verified bundle. Unrecognized wrappers are not
executed. The separate contract input and macro-side hash recomputation are
an unsigned consistency boundary, not a signature: a subject able to alter all
local inputs or the macro environment coherently remains inside the local trust
boundary.

LibreOffice may reorder the exact TOC instruction
`TOC \\o "1-3" \\h \\z` as `TOC \\z \\o "1-3" \\h`. The core may restore only
that precisely matched field-instruction permutation. It never overwrites
`sectPr`: section boundaries and the complete parsed section-property tree must
already have identical semantics. Any margin, paper, orientation, columns,
page-number start, section type, header/footer reference, property, value, or
boundary difference is blocking. If strict
writeback accepts only matched caches, a macOS smoke reports
`delivery_field_status=libreoffice_refreshed` and
`field_writeback_status=libreoffice_selective`; this proves only a non-final
LibreOffice backend result. If the refreshed TOC contains a non-text payload or
any other unapproved result, the smoke must instead report
`gate_outcome=strictly_deferred`, `backend_result_accepted=false`, and preserve
the rejection evidence. Neither outcome completes Word verification or permits
`final_ready`.

The macOS CI runner is pinned to `macos-15` and records macOS, Python,
LibreOffice, and Homebrew cask metadata. The LibreOffice cask itself is not
version-pinned, so upstream cask drift remains explicit evidence to review; the
safety-rejection smoke is separate from the Word field-completion gate.

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

The render manifest must identify the same normalized target requested by the
verification stage. A completed Word path must render the exact bound persistent
Word PDF and report `renderer_source=target_pdf`; a no-fields path must not gain
a target PDF and must identify the resolved renderer request used to create the
render. Switching Word and LibreOffice requests always reruns verification and
cannot reuse the other request's final-ready evidence.

The canonical finalization gate stores the allowlisted
`field_completion.evidence_validation`. Resume, verification, and status
recompute `final_ready_evidence_errors(completion_evidence(finalization))` and
require the stored validation to be exactly `status=pass` with `errors=[]`.
Missing, stale, or contradictory validation blocks final-ready even when the
remaining completion fields look valid.

## Recovery

1. Run `status --json` and inspect the failed or blocked stage.
2. Correct the source-independent input that caused the problem, such as a QA
   decision, approved map, renderer path, or backend authorization.
3. Rerun that command with `--resume`.
4. Rerun later stages whenever an upstream fingerprint changes.

Do not delete the source or overwrite it during recovery. Do not claim success
from an older cached artifact merely because its input key still matches; all
bound outputs and completion evidence must also revalidate.
