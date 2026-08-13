# Claude Code

## 安装

Claude Code 官方支持：

- 项目级：`.claude/skills/format-monograph/`
- 个人级：`~/.claude/skills/format-monograph/`
- 插件级：`<plugin>/skills/format-monograph/`

复制或链接完整的核心目录，不要复制适配文件到 skill 内。若会话启动时顶层 skills 目录不存在，创建后重新启动 Claude Code。

## 验证

1. 启动 Claude Code。
2. 确认 `/format-monograph` 可用，或用自然语言触发。
3. 请求从未批准规范修改 DOCX，确认 Claude 只生成规则草案并停止等待批准。
4. 运行环境检测和结构测试。

官方说明：https://code.claude.com/docs/en/slash-commands

整书运行必须遵循 `references/portable-run-checklist.md`，并使用
`scripts/run_monograph.py` 的统一阶段和状态。
