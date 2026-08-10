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

- `refreshed`: field instructions remain editable, cached results are present, and final content/object audits pass.
- `refreshed_target_word`: Microsoft Word repaginated the derivative, updated only approved fields, saved and reopened it, and the core audits passed.
- `deferred`: fields are marked dirty and Word-compatible update-on-open is enabled. This state requires explicit caller QA and is not a completed refresh.
- `code_only` or `stale`: do not deliver as finalized.
- `absent`: acceptable only when the approved document contains no fields that require refresh.

If a field updater removes field instructions, changes authored content, or changes protected payloads, reject its output. In `auto` mode, an explicitly approved deferred fallback may be used; otherwise stop.

Before accepting a finalized copy, resolve approved fonts through direct formatting, styles, base styles, document defaults, and the theme font scheme. Reject a backend result that reintroduces a theme-font mismatch even when its field-cache checks pass.

`target_pdf_ready_for_visual_qa` means the target application exported a PDF, not that layout passed. Mark `target_layout_verified` only after every exported page is inspected for page-number sequences, TOC entries, clipping, overlap, tables, captions, headers, footers, and equations.
