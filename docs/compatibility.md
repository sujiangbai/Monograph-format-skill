# 兼容性目标

核心 skill 遵循 Agent Skills 的 `SKILL.md + scripts + references + assets` 结构。平台适配只处理安装位置和发现方式。

| Agent | 目标级别 | 当前状态 |
| --- | --- | --- |
| OpenAI Codex | 核心发现、引用加载、Python 脚本调用 | 待实机验证 |
| Claude Code | 核心发现、引用加载、Python 脚本调用 | 待实机验证 |
| Gemini CLI | 核心发现、引用加载、Python 脚本调用 | 待实机验证 |
| VS Code / GitHub Copilot | 项目级 skill 发现和脚本调用 | 待实机验证 |
| 其他 Agent Skills 客户端 | 按开放规范兼容 | 规范级验证 |

## 能力等级

- **完整模式**：可执行 Python、处理 DOCX、调用 LibreOffice 或等效渲染器。
- **结构模式**：可处理 DOCX 但不能渲染；需要用户明确接受降级。
- **分析模式**：不能执行脚本；只能提取规则和提出 QA。

任何平台只有在真实环境完成发现、触发、资源读取和端到端测试后，才能标记为“已验证”。
