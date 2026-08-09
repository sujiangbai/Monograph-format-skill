# 能力等级

先运行 `scripts/check_environment.py --json`，分别读取 `inspection`、`profile_validation`、`docx_editing` 和 `rendering` 四项能力，再选择模式。某一项缺失不应被误报为其他能力也缺失。

## 完整模式

需要：

- Python 3.11 或更高版本。
- `python-docx`、`lxml`、`jsonschema` 和 `PyMuPDF`。
- LibreOffice `soffice` 或当前 Agent 能实际调用并逐页输出图像的等效 DOCX 渲染器。
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
- 不依赖 Word COM、平台 shell 脚本或 Agent 私有 API。
- 找不到字体时停止并提问，不要使用相近字体静默替换。
- 找到公式图片、断链字段或旧版 Equation Editor 对象时停止并提问。
- LaTeX 只有在能够转换并验证为可编辑 OMML 时才可进入 DOCX。
- 渲染器缺失可以降级；渲染器存在但转换失败时应修复错误，不能直接降级掩盖问题。

## Field finalization

Treat field finalization as an independent capability. `field_finalization=true` means an executable LibreOffice/UNO candidate exists; it does not guarantee that the candidate preserves Word fields or layout for a particular DOCX.

- `refreshed`: field instructions remain editable, cached results are present, and final content/object audits pass.
- `deferred`: fields are marked dirty and Word-compatible update-on-open is enabled. This state requires explicit caller QA and is not a completed refresh.
- `code_only` or `stale`: do not deliver as finalized.
- `absent`: acceptable only when the approved document contains no fields that require refresh.

If a field updater removes field instructions, changes authored content, or changes protected payloads, reject its output. In `auto` mode, an explicitly approved deferred fallback may be used; otherwise stop.
