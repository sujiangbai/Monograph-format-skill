# Monograph Format Skill

`format-monograph` 是一个遵循开放 Agent Skills 规范的跨 Agent 能力包，用于依据出版社或专著项目提供的格式资料，对 DOCX 专著进行可审计的格式修改。

当前状态：V0.1 开发中。核心目标包括：

- 从 DOCX、PDF、图片和文字要求中提取结构化格式规则。
- 在用户确认规则后修改专著 DOCX，不改动正文事实与措辞。
- 输出干净格式稿、审阅标注稿和格式报告。
- 在 Codex、Claude Code、Gemini CLI、VS Code / GitHub Copilot 等支持 `SKILL.md` 的 Agent 中复用。

## 仓库结构

- `format-monograph/`：可安装的开放标准 skill。
- `adapters/`：平台安装与发现适配，不复制核心规则。
- `docs/`：治理、兼容矩阵和不可覆盖的方案历史。
- `tests/`：合成测试与兼容性验证。

## 数据原则

不要向本公开仓库提交出版社原始规范、样书、用户书稿或运行产物。仅提交获得许可的规则摘要和合成测试数据。

## 开发流程

首次基线建立后，任何方案、skill 或适配变更都必须在新分支提交 PR，并在 QA 后获得仓库所有者明确同意才可合并。完整规则见 `docs/GOVERNANCE.md`。

## 许可证

Apache License 2.0。
