# VS Code / GitHub Copilot

## 安装

VS Code 官方支持以下项目位置：

- `.github/skills/format-monograph/`
- `.claude/skills/format-monograph/`
- `.agents/skills/format-monograph/`

个人 skill 可放入：

- `~/.copilot/skills/format-monograph/`
- `~/.claude/skills/format-monograph/`
- `~/.agents/skills/format-monograph/`

还可以通过 `chat.agentSkillsLocations` 配置其他目录。复制或链接完整的核心目录。

## 验证

1. 确认 `chat.useAgentSkills` 已启用。
2. 在 Chat 中运行 `/skills`，确认 `format-monograph` 已发现。
3. 用触发与非触发提示测试自动加载。
4. 确认终端工具在用户授权后可以运行 Python 脚本。

官方说明：https://code.visualstudio.com/docs/agent-customization/agent-skills
