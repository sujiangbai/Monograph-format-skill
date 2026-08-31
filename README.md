# Monograph Format Skill

`format-monograph` 是一个遵循开放 Agent Skills 规范的跨 Agent 能力包，用于对中文或多语种专著、教材类 DOCX 文稿进行可审计的格式修改。

它面向需要在 Codex、Claude Code、Gemini CLI、VS Code / GitHub Copilot 等支持 `SKILL.md` 的 Agent 中复用排版能力的人。这个 skill 的目标不是替作者改写内容，而是在用户确认规则后，依据出版社规范、样书、模板或明确要求，对 Word 文档进行受控格式化、审计和交付。

当前状态：**V0.3.3 候选实现**。

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
- 修复独立嵌入图片因固定行距或简单固定表格行高导致的裁切，不改变图片尺寸、锚点、环绕或文中位置。
- 目录只读取已批准语义标题；标题样式来源受图片污染时整份目录切换到纯文本 `TC` 来源，并拒绝含图片或越界条目的回写结果。
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

V0.3.3 继续使用统一运行入口：

```text
<python> format-monograph/scripts/run_monograph.py prepare <input.docx> --profile <profile.json> --work-dir <directory>
<python> format-monograph/scripts/run_monograph.py apply --work-dir <directory> --structure-map <approved.json>
<python> format-monograph/scripts/run_monograph.py finalize --work-dir <directory>
<python> format-monograph/scripts/run_monograph.py verify --work-dir <directory> --visual-qa-manifest <visual-qa.json>
```

中断后在原命令上添加 `--resume`。`run-state.json` 和所有书稿产物只保存在用户指定的本地任务目录。

终稿发布采用同目录 staging、既存目标原子捕获和平台原生的 no-replace 原子移动，并把状态 JSON 作为最后的提交标记。macOS 使用 `renameatx_np(RENAME_EXCL)`，Linux 使用 `renameat2(RENAME_NOREPLACE)`。创建 staging 与 producer child 时，程序从已打开的 parent authority 生成随机 basename，以 `mkdir(dir_fd=...)` exclusive 创建，立即 no-follow stat 捕获完整 identity，再 parent-relative open 并逐项比对；冲突只重试新名称，绝不收养既存目录。任何 hook 或 producer operation 都晚于 create-and-bind helper 成功返回。最终 staged DOCX/PDF/status/audit 只通过 authority-relative `O_EXCL|O_NOFOLLOW` 导入；main 保存每个写入 FD 的完整 identity，publisher 必须先与该独立 identity 逐项相等，不能从首次重读重新定义 expected，也不能覆盖、截断或删除预存同名对象。不支持、跨文件系统或权限失败一律关闭失败，不回退到 hard link、普通 rename、copy 或覆盖。Windows 当前明确不可用，并在创建缺失 output parent 前 fail-close：pathname `MoveFileExW` 无法把已验证 source/root handle 与 no-replace 操作绑定，且尚无 AccessCheck 有效权限证明；SDDL 字符串、Darwin 模拟和 Windows CI 的拒绝测试都不构成 Windows publisher 可用证据。完整 status 只序列化一次，并在提交前后核验独立 identity、closed shape、completion 与工件绑定。输出父目录、staging、producer workspace、recovery 和 transaction 目录持续核验其权威及路径条目；成功替换的旧 inode 被保留到 `.format-monograph-recovery/<transaction-id>/`。生产代码不自动删除 staging：成功时由原子 move 消耗 staged 工件，以 exclusive 写入 FD 的完整 identity 绑定非业务门禁 `publication-record.json`，hook 后重新核对 record identity 与精确 staging entry set，并保留 producer evidence；失败时保留 staging/quarantine/backup/unknown。所有目录都只允许人工清理。commit point 后的 authority close/report 异常以 `committed_with_cleanup_errors` 诊断返回；stderr/stdout 都是吞掉控制台故障的 best-effort 报告，不回滚、不把已提交状态误报为普通发布失败，且所有 close 仍逐项尝试。私有 staging/recovery 以当前有效用户为控制边界；能以同一 UID 协调修改全部本地证据的主体位于既有无签名信任边界内。该协议不是全局多文件原子事务，不承诺 `fsync`/掉电持久性，status 也不是锁；consumer、`resume`、`verify` 和 `status` 必须继续重验当前工件。当前 macOS 有本机实测；Windows 保持 NO-GO，只有未来另行批准的 AccessCheck 与 authority-bound 原生 rename 实现通过真实 Windows 对抗门禁后才可开放。

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
