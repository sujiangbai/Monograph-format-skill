# 兼容性目标

核心 skill 遵循 Agent Skills 的 `SKILL.md + scripts + references + assets` 结构。平台适配只处理安装位置和发现方式。

| Agent | 项目级位置 | 个人级位置 | 当前状态 |
| --- | --- | --- | --- |
| OpenAI Codex | `.agents/skills/format-monograph/`（需当前客户端支持） | `$CODEX_HOME/skills/format-monograph/` | 待实机验证 |
| Claude Code | `.claude/skills/format-monograph/` | `~/.claude/skills/format-monograph/` | 待实机验证 |
| Gemini CLI | `.gemini/skills/format-monograph/` 或 `.agents/skills/format-monograph/` | `~/.gemini/skills/format-monograph/` | 待实机验证 |
| VS Code / GitHub Copilot | `.github/skills/`、`.claude/skills/` 或 `.agents/skills/` | `~/.copilot/skills/`、`~/.claude/skills/` 或 `~/.agents/skills/` | 待实机验证 |
| 其他 Agent Skills 客户端 | 按客户端配置 | 按客户端配置 | 规范级验证 |

## 能力等级

- **完整模式**：可执行 Python、处理 DOCX，并可调用 LibreOffice 或经授权的目标软件后端完成渲染。
- **结构模式**：可处理 DOCX 但不能渲染；需要用户明确接受降级。
- **分析模式**：不能执行脚本；只能提取规则和提出 QA。

## 验证原则

- 开放规范通过 `skills-ref validate` 检查。
- 平台发现、触发、资源读取和脚本调用必须在真实客户端分别验证。
- 没有可用账号或运行环境时保持“待实机验证”，不得根据规范兼容推断为实机通过。
- 使用 `tests/compatibility/smoke-test.md` 和统一提示集记录结果。
- Microsoft Word Windows 适配器是独立目标软件能力；CI 只验证协议和脚本静态安全，未配置 Word 的环境不得标记实机通过。

## 官方依据

- Agent Skills 规范：https://agentskills.io/specification
- Claude Code skills：https://code.claude.com/docs/en/slash-commands
- Gemini CLI skills：https://geminicli.com/docs/cli/tutorials/skills-getting-started/
- VS Code Agent Skills：https://code.visualstudio.com/docs/agent-customization/agent-skills

Codex 的准确发现位置可能因产品形态和管理员策略变化；适配文档要求优先采用当前 Codex 显示的 skills 目录，并在真实环境验证。
