# Monograph Format Skill

`format-monograph` 是一个遵循开放 Agent Skills 规范的跨 Agent 能力包，用于对中文或多语种专著、教材类 DOCX 文稿进行可审计的格式修改。

它面向需要在 Codex、Claude Code、Gemini CLI、VS Code / GitHub Copilot 等支持 `SKILL.md` 的 Agent 中复用排版能力的人。这个 skill 的目标不是替作者改写内容，而是在用户确认规则后，依据出版社规范、样书、模板或明确要求，对 Word 文档进行受控格式化、审计和交付。

当前状态：**V0.3.1 候选实现**。

## 适合谁使用

- **Agent 操作者**：需要让 AI 依据明确规则处理专著或教材 DOCX 排版。
- **项目维护者和贡献者**：需要理解 skill 结构、规则边界、测试与治理流程。
- **审校或出版项目负责人**：需要确认格式修改不会改变正文事实、公式语义、答案、引用或原始内容。

## 核心能力

- 从 DOCX、PDF、图片、样书和文字要求中提取结构化格式规则。
- 以当前调用者明确要求和后续 QA 决定作为最高格式优先级。
- 在规则获得确认后修改 DOCX，并保留原始输入文件。
- 不改动正文事实、措辞、公式语义、答案或引用。
- 保留 OMML、MathType/OLE 等可编辑公式对象，阻止公式被静默转换为图片。
- 支持真实 Word 字段、多级标题编号、镜像页边距、中西文字体分离和固定磅值行距。
- 解析 DOCX 样式、主题和继承链后的实际字体，避免主题字体造成审计误判。
- 通过结构映射处理目录、正文页码分区、奇偶页脚、表格、题注、图片和章节结构。
- 通过统一的分阶段命令、运行状态和指纹缓存支持整书断点续跑。
- 按组提出重复 QA，只冻结未决章节或对象，并保留书稿已有附录编号。
- 将代表性试排限制为不超过约 30 个渲染页，不把试排误作整书验收。
- 在临时目标软件副本中更新字段，只回写通过审计的字段结果。
- 输出格式稿、审阅标注稿、格式报告、审计结果和渲染检查依据。

## 安全边界

使用本 skill 时必须遵守以下边界：

- 不覆盖原始 DOCX。
- 不擅自改写作者内容。
- 不静默修复、补写或推断正文、答案、引用和公式。
- 不把可编辑公式转换为 PNG、JPEG、SVG、EMF 等图片。
- 不在未确认字体缺失、字段刷新、表格拆分、页码分区或结构映射时宣称交付通过。
- 不向公开仓库提交用户书稿、出版社规范、样书截图、格式成品或渲染页面。

## 基本使用流程

1. 在支持 Agent Skills 的工具中安装或加载 `format-monograph/` 目录。
2. 让 Agent 读取 `format-monograph/SKILL.md` 作为执行入口。
3. 将用户书稿、出版社规范、样书和运行产物放在 skill 目录之外。
4. 先提取并生成格式 profile。
5. 由用户确认格式规则、冲突项和阻塞问题。
6. 检查 DOCX 结构并生成 structure map。
7. 由用户逐项确认结构操作。
8. 应用已批准规则，生成格式稿和审阅稿。
9. 运行审计、字段终稿化和渲染检查。
10. 交付前确认内容一致性、公式对象、媒体对象、字段缓存、页码和视觉 QA。

V0.3.1 整书运行使用统一入口：

```text
<python> format-monograph/scripts/run_monograph.py prepare <input.docx> --profile <profile.json> --work-dir <directory>
<python> format-monograph/scripts/run_monograph.py apply --work-dir <directory> --structure-map <approved.json>
<python> format-monograph/scripts/run_monograph.py finalize --work-dir <directory>
<python> format-monograph/scripts/run_monograph.py verify --work-dir <directory> --visual-qa-manifest <visual-qa.json>
```

中断后在原命令上添加 `--resume`。`run-state.json` 和所有书稿产物只保存在用户指定的本地任务目录。

详细执行规则见 [`format-monograph/SKILL.md`](format-monograph/SKILL.md)。

## 仓库结构

- `format-monograph/`：可安装的开放 Agent Skill 核心。
- `adapters/`：不同平台或目标软件的安装、发现和字段刷新适配。
- `docs/`：治理规则、方案历史和设计说明。
- `tests/`：合成测试与兼容性验证。

## 数据原则

本仓库是公开仓库，只应提交可以公开的 skill 规则、脚本、文档和合成测试数据。

禁止提交：

- 用户原始书稿。
- 出版社未公开授权的规范、模板、样书或截图。
- 格式修改成品、渲染页面和运行产物。
- 含个人信息、密钥或访问令牌的文件。

## 开发与治理

`main` 是稳定分支。首次基线之后，方案、skill、适配器或修复变更应通过新分支和 PR 进行。

方案文件是不可变快照；如果方案需要调整，应创建更高版本的新方案文件，而不是覆盖旧文件。

完整规则见 [`docs/GOVERNANCE.md`](docs/GOVERNANCE.md)，方案索引见 [`docs/plans/INDEX.md`](docs/plans/INDEX.md)。

## 许可证

本项目采用 [Apache License 2.0](LICENSE)。

可以使用、修改和分发本项目，也可以用于商业场景；但需要保留许可证和版权说明，且项目按原样提供，不附带担保。
