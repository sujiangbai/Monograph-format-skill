# OpenAI Codex

## 安装

个人级安装应将完整的 `format-monograph/` 目录放入 Codex 的 skills 目录：

- 优先使用 `$CODEX_HOME/skills/format-monograph/`。
- 未设置 `CODEX_HOME` 时，常用默认位置为 `~/.codex/skills/format-monograph/`。
- 支持开放 Agent Skills 项目发现的 Codex 环境也可使用 `.agents/skills/format-monograph/`。

不要只复制 `SKILL.md`，核心需要同级的 `scripts/` 和 `references/`。

## 验证

1. 重新加载 skill 列表或开启新任务。
2. 确认 `format-monograph` 出现在可用 skills 中。
3. 请求“分析出版社规范并为专著建立格式配置”，确认 Codex 先读取 `SKILL.md`。
4. 运行环境检测并确认能力等级。
5. 使用 `tests/compatibility/activation-prompts.json` 完成触发测试。

Codex 安装位置可能随产品形态和管理员策略变化。若当前 Codex 提供 skill 安装器或自定义目录配置，以当前产品显示的目录为准。

整书运行必须遵循 `references/portable-run-checklist.md`，并使用
`scripts/run_monograph.py` 的统一阶段和状态。
