# 平台适配

本目录说明各 Agent 的发现方式和可选平台能力。不要在适配目录复制 `SKILL.md`、核心脚本或规则文件。平台专用自动化只能通过核心定义的外部 JSON 协议接入。

核心的 external 接入边界在协议与 Word 目标检查通过后，将既有 `external` 或 macOS 具体实现标签 `microsoft_word_macos_applescript` 归为业务类别 `external`。三个操作（refresh_fields、measure_layout、verify_only）使用同一分类；原标签保存在 raw backend audit 的 `implementation_backend`，不参与完成门禁。其他标签不改写、不授予 external 身份，继续经过原有 canonical/字段完成验证。Canonical 证据及字段完成态继续使用既有业务类别，不扩充枚举或 Schema；具体实现名称只是诊断，不是应用认证或成功证明。

安装时应复制或链接完整的 `format-monograph` 目录。安装后先审阅脚本，再安装 `requirements.txt` 中的依赖，并运行：

`<python> scripts/check_environment.py --json`

平台适配不得改变规则优先级、批准闸门、内容保护或逐页视觉验收要求。

`microsoft-word/windows/` 和 `microsoft-word/macos/` 是可选目标软件适配器。它们不影响其他平台发现 Skill；只有对应平台、Microsoft Word 和调用者对精确输入的自动化授权均满足时才可使用。macOS 适配器还会在 Word 已打开任何文档或正在后台打印时拒绝运行。

macOS Word 适配器提供可选、有限支持的自动最终化：在 macOS 26.5.1 / Word 16.112.3 下，前次无 TOC 合成场景和当前代码含 TOC 合成场景均已通过正常 finalize、逐页视觉批准及 verify/status 闭合。唯一字段可按严格契约补写；重复字段只有模型值相同且保存值逐项一致时可确认，异值重复、缺关联或嵌套目标无法确认则不完成。完整边界与首次文件访问授权要求见其 INSTALL.md；字段白名单不代表所有 story、复杂图表或整个产品均已验证。超时不自动重试或冒充完成，权限就绪后可经新授权从现有正常入口重新运行并保留失败证据。

所有 Agent 适配器必须执行 `format-monograph/references/portable-run-checklist.md`
中的统一阶段。适配器只能改变发现和调用方式，不得改变运行状态、QA 或交付闸门。
