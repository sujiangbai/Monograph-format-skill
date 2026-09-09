# Microsoft Word macOS adapter

## Limited-support automatic finalization

This adapter is optional, not a generally validated Word integration. Local
synthetic evidence is limited to macOS 26.5.1 and Word 16.112.3 and the specific
scenarios exercised; the supported field list below does not establish coverage
of every header/footer story combination. A prior synthetic no-TOC flow and a
current-code synthetic TOC flow each completed normal finalize, independent
full-page visual approval, verify and status through final_ready. The TOC flow
included bounded footer supplementation, strict selective writeback, a second
Word no-save verification and a persistent three-page PDF. Offline tests also
reject incorrect or unconfirmed outgoing footer caches. These results do not
establish general manuscript, complex-object, or all-story coverage, or
completion of the whole product or other project stages.

For approved, non-nested footer PAGE/NUMPAGES with a unique complete instruction
key inside the exact existing story owner, the adapter can supplement a stale
saved cache using the converged result read from that same Word field object.
The owner is attached when the group is created, not guessed from flattened
row positions; these associations participate in the existing convergence check.
The complete story inventory and existing semantic owner/part mapping must
agree. Before changing any cache, the untouched Word candidate passes the
existing selective-writeback content, field and structure checks. Existing
scalar payload/container checks are reused and the supplemented candidate is
checked again. Unchanged caches are not rewritten.

Duplicate, nested or insufficiently associated fields receive no model-result
supplement. Skipping a supplement is not successful cache verification. Before
refresh success, outgoing top-level footer PAGE/NUMPAGES caches must match the
converged owner-associated model results and pass existing scalar checks.
Duplicate identities pass this read-only check only when every observed value
is identical and every saved occurrence equals it; differing values are not
paired by order. Missing associations and nested PAGE/NUMPAGES parents cannot
be confirmed by this mechanism and remain incomplete. This does not impose a
new global rejection on unrelated nested fields. The unique-field path passed
the current live TOC flow; duplicate/nested refusal and confirmation boundaries
are covered by offline tests, not expanded into new writeback capabilities.
No page count is manufactured into a result. A PAGE cache belongs to the exact
observed field, not to every page on which its footer is displayed. TOC, body
and other fields retain their XML route. The raw backend's footer_cache_source
records original/supplemented candidate hashes and a count, not field text.
The final persisted NUMPAGES guard checks the actual outgoing candidate.
Core selective writeback and final read-only/PDF verification are unchanged;
verify_only never supplements, refreshes or saves a DOCX.

This optional adapter implements the existing external field protocol 1.1 for
desktop Microsoft Word on macOS. It is a bounded calculation service, not a
delivery writer: the core selectively imports approved field results from the
disposable Word copy. Missing or ambiguous original roles, changed original
content, fields, bookmarks, activation or pagination remain errors. Only extra
roles backed by independently owned, valid, empty header/footer shells may be
discarded. Shells cannot share an original role's candidate field source or
contain text, fields, symbols, breaks, tables, drawings, objects, relationships
or unknown substantive payload. Localized paragraph style identifiers are not
resolved as an acceptance authority; candidate styles are never imported.
The output retains the baseline role graph, references and inheritance. This
bounded candidate policy does not authorize arbitrary Word package changes.

## Requirements

- macOS with desktop Microsoft Word installed and activated.
- Python 3.11 or newer with `format-monograph/requirements.txt` installed.
- macOS Automation permission for the invoking terminal or agent.
- Explicit caller approval to automate Word for the synthetic or user-selected
  input. The adapter refuses to run while Word has any document open or a
  background print job active; it never closes an unrelated document or quits
  Word.

Respond promptly to macOS Automation or file-access prompts for the authorized
input; refresh and later read-only copies may prompt separately. A timeout is
not an automatic recovery or completion. After permissions are ready and Word
is idle, a newly authorized normal finalization may reuse the existing work and
input paths while preserving failed-run evidence. Do not reuse a standalone
verification PDF as a successful finalization. The observed open timeouts do
not establish a diagnosed or universally resolved Word defect.

Pass the adapter as a JSON argument array. Use exact absolute paths for both the
Python executable and this script:

```text
<python> format-monograph/scripts/finalize_docx.py <formatted.docx> \
  --source <source.docx> \
  --profile <approved-profile.json> \
  --structure-map <approved-structure-map.json> \
  --output <finalized.docx> \
  --field-updater external \
  --field-updater-command '["<python>","<repo>/adapters/microsoft-word/macos/word_field_updater.py"]' \
  --target-software microsoft_word \
  --pdf-output <persistent-word-verification.pdf>
```

The Python entry validates the exact non-symlink paths, target ID, field
whitelist, approved existing TOC, input hash, field instructions and PDF entity.
For verification, Word always exports once to a private short temporary directory
under `/private/tmp`. Only after the existing read-only, snapshot, close/restore
and PDF page-count checks pass are identical PDF bytes copied with exclusive
creation to the original requested output. An occupied target is rejected and
left untouched. Temporary exports and this call's failed partial copy are cleaned
up; the protocol target and core publisher are unchanged. This is one path, not
a length threshold, retry, or alternate PDF backend. Local synthetic long-path
failures motivated it; no universal Word path-length limit is asserted.
Word's `/var/` representation is accepted only for the already resolved target
under `/private/var/`, with matching regular-file device/inode identity. Other
spellings, missing/changed entities and multiple matches fail closed. Lookup
and close share this rule, including the selected-object and close rechecks.
The identity binding can change only immediately after the already-held,
prechecked document's approved save and successful saved-state readback, with
its path and resulting file rechecked; this permits Word's atomic save while
not allowing close to silently rebind. This is bounded local consistency, not
an atomic filesystem transaction or protection from coordinated local tampering.
The AppleScript opens only the requested path with recent-file recording off,
forces macro execution off, disables link updates, temporarily disables
print-time field/link refresh, and verifies every changed preference before and
after restoration. It never lowers a security preference.

Some Word for Mac versions do not expose the three print-time field, link, and
field-code settings and return `missing value`. The adapter never writes such
an unreadable optional property and requires it to remain missing before work,
after work, and during restoration. Boolean values are still set to false,
checked, restored, and checked again. PDF no-update evidence additionally
requires read-only open, identical complete snapshots before and after export,
explicit close without saving, an unchanged DOCX hash, and PDF entity/page
verification; an unavailable optional property does not prove safety by itself.

`measure_layout` and `verify_only` open the requested DOCX read-only and close it
without saving. `refresh_fields` first works on the new disposable output path,
updates the approved existing TOC before repagination, then updates individual
approved `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `PAGEREF`, `REF`, and exact
core-generated PAGE-minus-one formula fields. It never invokes an update-all
collection operation. Two consecutive complete tuples must match within three
rounds. Persisted convergence evidence contains TOC entry counts, order and text
hashes, never TOC body text.

The core additionally compares the complete final refresh and read-only
verification snapshots, ignoring only their top-level round number. Equal
page counts alone are insufficient. Older Windows adapters without equivalent
convergence evidence cannot meet this gate; this batch does not upgrade them.
The current macOS adapter requires an existing effective primary footer in each
section for page-number observations. It refuses missing primary footers rather
than inventing them; approved preprocessing must already satisfy this condition.

PDF export can make Word's in-memory `saved` property false even for a read-only
document. The adapter therefore records Boolean observations, captures equal
complete snapshots before and after export, closes explicitly without saving,
and lets the parent verify the unchanged DOCX hash plus PDF hash/page binding.
A dialog, timeout, missing PDF, ambiguous path, unexpected field instruction,
failed close, or uncertain preference restoration is an error and cannot be
reported as `final_ready`.

Normal automated tests use synthetic DOCX files and mock the `osascript`
boundary. Do not run a live Word smoke without separate approval for the exact
synthetic input and output paths.
