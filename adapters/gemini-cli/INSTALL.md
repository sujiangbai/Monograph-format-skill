# Gemini CLI

## 安装

Gemini CLI 支持：

- 项目级：`.gemini/skills/format-monograph/`
- 项目级开放标准别名：`.agents/skills/format-monograph/`
- 个人级：`~/.gemini/skills/format-monograph/`

也可以使用：

- `gemini skills install <url-or-path>`
- `gemini skills link <path>`

工作区 skill 只在受信任目录加载。复制或链接完整核心目录后，运行 `/skills reload`。

## 验证

1. 运行 `/skills list` 并确认 skill 已发现。
2. 用自然语言请求专著格式分析，批准 `activate_skill`。
3. 确认引用文件和 Python 脚本均可访问。
4. 执行环境检测和结构测试。

官方说明：https://geminicli.com/docs/cli/tutorials/skills-getting-started/

整书运行必须遵循 `references/portable-run-checklist.md`，并使用
`scripts/run_monograph.py` 的统一阶段和状态。
