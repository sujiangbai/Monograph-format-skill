# 能力等级

先运行 `scripts/check_environment.py --json`，分别读取 `inspection`、`profile_validation`、`docx_editing`、`rendering`、`word_automation`、`word_field_refresh` 和 `word_pdf_export`。某一项缺失不应被误报为其他能力也缺失。

## 完整模式

需要：

- Python 3.11 或更高版本。
- `python-docx`、`lxml`、`jsonschema` 和 `PyMuPDF`。
- LibreOffice `soffice`，或经调用者授权且能导出 PDF 的目标软件后端；将 PDF 转为逐页图像仍需要 `PyMuPDF`。
- 可读写用户指定的任务目录。
- 能够保留和验证文档中的可编辑公式对象。

允许修改、审计、渲染并交付完整三件套。逐页图像仍必须由 Agent 实际检查，脚本成功不等于视觉通过。

## 结构模式

需要 `inspection`、`profile_validation` 和 `docx_editing` 能力，但没有可用 `rendering` 能力。允许结构修改和审计，但必须先获得用户接受降级，并在报告中写明未进行视觉 QA。

如果必需字体缺失，`apply_profile.py` 默认阻塞。只有当前调用者完成 QA 并明确接受后，才可使用 `--allow-missing-fonts`；报告必须列出缺失字体且视觉验收仍为未完成。

## 分析模式

无法执行核心脚本、缺少 `docx_editing` 能力，或不能保持和验证可编辑公式。只允许分析资料、起草配置、列出冲突和提出 QA。仅缺少 `profile_validation` 时可以盘点 DOCX，但不得批准或应用未经模式验证的配置。不得修改文件、以图片替代公式或声称验证通过。

## 环境原则

- 使用当前 Agent 可用的 Python 命令，不在核心指令中假定 `python`、`python3` 或 `py` 中的某一个。
- 开放 Skill 核心不依赖 Word COM、平台 shell 脚本或 Agent 私有 API；平台适配器可以通过通用外部 JSON 协议提供这些能力。
- 找不到字体时停止并提问，不要使用相近字体静默替换。
- 找到公式图片、断链字段或旧版 Equation Editor 对象时停止并提问。
- LaTeX 只有在能够转换并验证为可编辑 OMML 时才可进入 DOCX。
- 渲染器缺失可以降级；渲染器存在但转换失败时应修复错误，不能直接降级掩盖问题。

## Field finalization

Treat field finalization as an independent capability. `field_finalization=true` means an executable external or LibreOffice/UNO candidate exists; it does not guarantee that the candidate preserves fields or layout for a particular DOCX. `word_automation=true` still requires caller authorization and a successful live run.

On macOS, prefer an explicitly configured renderer and otherwise use the installed `/Applications/LibreOffice.app` before a PATH wrapper. Run the LibreOffice Python field helper as a user macro inside a disposable LibreOffice profile; do not launch the app's embedded Python executable from an unrelated parent process, mix it with another Python runtime, weaken the caller's global macro policy, or embed the helper in the delivery DOCX. Reject `TargetMode=External`, package-field connections, all URI schemes (including `file`, `ftp`, `smb`, `http`, and `https`), URI-relative and backslash network paths, package escapes, malformed targets, and unresolved XML `href`/`src` references before launch. Resolve `.rels` targets from their OPC source part after percent-decoding and backslash normalization; only fragments and relative or package-root references to members that actually exist in the package are accepted. Load accepted input with UNO `UpdateDocMode.NO_UPDATE` and `MacroExecMode.NEVER_EXECUTE`; update only a `ContentIndex` whose complete hashed identity matches the approved structure contract, exact baseline OOXML TOC instruction and position, occurrence/ordinal/service, and observed UNO `CreateFromOutline`, `CreateFromMarks`, and `Level` properties. The canonical TOC contract is a macro input separate from the authorization. The macro validates and canonicalizes it, recomputes SHA-256, and compares that hash with the authorization before any `update()`. This is an unsigned consistency check, not a signature or protection against coherent alteration of every local input or the macro environment. A count, contract, or identity mismatch rejects before any `update()`. Perform only that authorized update plus document-internal calculations, and never call collection-level `TextFields.refresh()`. Publish success only after the document closes successfully. A successful synthetic smoke reports `libreoffice_refreshed` plus `libreoffice_selective` only after exact baseline field instructions, ordering, boundaries, and pagination semantics are verified. `sectPr` is never overwritten to hide a backend difference. This is a non-final backend state. A strict rejection may pass the safety smoke only as `strictly_deferred` with `backend_result_accepted=false` and the rejection evidence retained; it is not a backend field-refresh success. Both states still require Word no-save verification, do not complete the field gate, and cannot become `final_ready`.

- `refreshed`: field instructions remain editable, cached results are present, and final content/object audits pass.
- `selective_verified`: input cache is `stale`, output cache is `refreshed`, the core imported only uniquely matched approved field results into its baseline, and target Word reopened the selective output without saving, reproduced the page count, and exported a persistent verification PDF. The finalized DOCX and PDF are bound by path, SHA-256, size, and page count and are rechecked before `final_ready`.
- `libreoffice_refreshed`: LibreOffice updated only approved internal indexes/calculations and the core selectively imported matched caches while preserving the exact baseline field contract. This is not Word-verified and is not final-ready.
- `refreshed_target_word`: legacy status from earlier implementations; do not use it for a new `final_ready` claim without the V0.3.2 selective and no-save verification evidence.
- `deferred`: fields are marked dirty and Word-compatible update-on-open is enabled. This state requires explicit caller QA and is not a completed refresh.
- `code_only` or `stale`: do not deliver as finalized.
- `absent`: acceptable only when the approved document contains no fields that require refresh.

If a field updater removes field instructions, changes authored content, approved bookmarks, pagination structure, or protected payloads, reject its field results. Discard all backend non-field serialization even after a successful refresh. In `auto` mode, an explicitly approved deferred fallback may be used; otherwise stop.

Before accepting a finalized copy, resolve approved fonts through direct formatting, styles, base styles, document defaults, and the theme font scheme. Reject a backend result that reintroduces a theme-font mismatch even when its field-cache checks pass.

`target_pdf_ready_for_visual_qa` means the target application exported a PDF, not that layout passed. Mark `target_layout_verified` only after every exported page is inspected for page-number sequences, TOC entries, clipping, overlap, tables, captions, headers, footers, and equations.

Whole-book `final_ready` additionally requires the versioned allowlisted
finalization-gate summary and versioned finalization/verification request
identities to match the current state, stage input keys, artifacts, target PDF,
and render manifest. Only an actually used renderer is bound by normalized
request source, resolved path, and current file digest. External updater
identities bind shared-parser argv, fixed arguments, executable, and local
script/file arguments by path, digest, and size. These are local consistency
checks, not signatures or application authentication. Target evidence uses
exact allowlisted IDs; changing a used renderer, target, updater entity or
argument, deferred approval, or visual-manifest path/hash reruns the affected
stage and cannot reuse a previous target's completion claim.

## Portable capability snapshot

V0.3.0 also emits `portable_capabilities` for file reading, Python execution,
DOCX inspection, profile validation, DOCX editing, font discovery, rendering,
target Word, field update, and multimodal source reading. The legacy
`capabilities` keys remain available for schema compatibility. An Agent must
declare multimodal source-reading support separately; the core cannot infer it.
LibreOffice field refresh is available only through the verified macOS internal
macro host. Other platforms or unrecognized executables may still render safe
fixtures, but field refresh is unavailable/deferred; the legacy UNO
server/helper is never a capability fallback.

Use the same snapshot in `run-state.json` for every staged command. An adapter
may add an authorized target-application backend, but it cannot turn an
unverified capability into a completed field refresh or target-layout check.
