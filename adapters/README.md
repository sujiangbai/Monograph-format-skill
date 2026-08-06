# 平台适配

本目录只说明如何让各 Agent 发现仓库根目录下的 `format-monograph/`。不要在适配目录复制 `SKILL.md`、脚本或规则文件。

安装时应复制或链接完整的 `format-monograph` 目录。安装后先审阅脚本，再安装 `requirements.txt` 中的依赖，并运行：

`<python> scripts/check_environment.py --json`

平台适配不得改变规则优先级、批准闸门、内容保护或逐页视觉验收要求。
