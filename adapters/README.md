# 平台适配

本目录说明各 Agent 的发现方式和可选平台能力。不要在适配目录复制 `SKILL.md`、核心脚本或规则文件。平台专用自动化只能通过核心定义的外部 JSON 协议接入。

安装时应复制或链接完整的 `format-monograph` 目录。安装后先审阅脚本，再安装 `requirements.txt` 中的依赖，并运行：

`<python> scripts/check_environment.py --json`

平台适配不得改变规则优先级、批准闸门、内容保护或逐页视觉验收要求。

`microsoft-word/windows/` 是可选目标软件适配器。它不影响其他平台发现 Skill；只有 Windows、Microsoft Word 和调用者授权均满足时才可使用。
