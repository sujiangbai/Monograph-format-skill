# QA 与报告

## 何时提问

仅在答案会改变规则或交付结果时提问。每轮优先提出一至三个问题，并给出推荐选项及影响。

必须提问的情况：

- 当前调用者自己的要求互相冲突，或多个资料的适用关系不清。
- 低置信度规则、未识别对象或静态字段迁移存在歧义。
- 图表编号可能同时包含出版编号与建筑、土木、结构工程图纸中的剖面、节点、详图或大样标识。
- 字体缺失或单位不明确。
- 公式图片、旧版 Equation Editor 对象、无法验证的 LaTeX 转换或公式对象哈希变化。
- `REF/PAGEREF` 指向不存在的书签。
- 目录或正文起点不明确，偶数页脚缺失，正文存在意外页码重启，或其他前置页是否显示页码未确定。
- 镜像边距导致目标软件在从 1 重启的目录或正文分节前自动插入奇偶补齐页，而当前后端不能自动采用匹配物理奇偶性的分节起始、可编辑显示偏移页码字段和 `PAGEREF` 校正。
- 表格列角色不明、复杂合并、浮动对象、可见控制标记、超宽溢出或横向分节未经逐表确认。
- 目标软件自动化未授权、受保护视图阻止、字段后端失败或后端返回未经批准的字段类型。
- 规则适用范围不清、结构模式交付或自动规则超出脚本能力。
- PR 合并授权。

明确的调用者格式要求高于较低级来源，但仍需在报告中记录冲突。安全边界不能被格式优先级取消。

## 报告结构

`<stem>-format-report.md` 必须包含：

1. 输入文件、配置名称、配置版本、批准信息和来源优先级。
2. 能力模式、环境检查摘要、缺失字体及任何明确批准的降级。
3. 已应用规则、命中数量、未命中规则和 `manual_review` 项。
4. 批准的派生字段变更清单、逐项批准的人工题注标识替换，以及 TOC、SEQ、REF、PAGEREF、PAGE 和书签状态。
5. 正文指纹、OMML、嵌入对象和媒体哈希对比结果。
6. 公式对象分类、公式图片候选和旧版公式对象状态。
7. 目录/正文页码分区、每个重启点、奇偶 PAGE 页脚、首页显示状态和孤立页眉页脚清理数量。
8. 每个获批数据表的列角色、宽度、边距、边框、跨页、横向分节及未决视觉 QA。
9. 冲突、QA 决定、字段刷新后端、目标软件版本、兼容性和其他限制。
10. 渲染页数、PDF 来源、逐页检查状态和发现的问题。
11. 最终结论：通过、带限制通过或失败。

逐表视觉 QA 必须检查短标签、编号、数值与单位是否出现可避免的孤立换行。不得让正文首行缩进泄漏到表格单元格；若短内容因列宽、单元格边距、段落缩进或错误对齐被拆行，应先恢复原有合理对齐并消除非语义缩进，再考虑调整列宽。不得通过改写文字掩盖排版问题。

逐页检查还必须确认书名页无页码、书名与目录分属不同页面、书名符合批准的字体/字号/字重、目录题名为批准文字且居中加粗。逐级比较多级列表编号与对应 Heading 样式的中西文字体、字号和字重，并确认 Heading 1-4 首行缩进为 0。检查每个图表后的批准空段落：同页后续内容前必须恰有一个，跨页时新页顶部不得保留。逐表报告底纹、主线/次线磅值、左右外边框、数据行横线，以及多行表头的局部分隔线是否避开纵向合并单元格。

## V0.2.6 确定性字体与分页证据

对每条获批的自动字体规则，报告目标字体、样式中声明的显式字体、原主题引用、沿继承链解析出的实际生效字体及其来源。正文、标题、目录、题注、表格文字、脚注/尾注、长引文、参考文献和教学框中凡已获批自动处理的对象，都必须按同一方法检查。检测到等线、等线 Light 或其他主题字体并不自动表示错误；只有它与当前任务批准字体不一致时才失败。未获批对象保持原状并列入未覆盖范围。

对每个页码页脚报告 `PAGE` 字段数量、奇偶页位置、是否包含非页码内容，以及本次运行是否执行了重复字段规范化、获批的静态数字转换或 `{ = { PAGE } - 1 }` 显示偏移。报告物理分节起始、各节显示偏移、校正的 `PAGEREF` 数量和 TOC 锁定状态；这些项目全部成功时不得再要求调用者手动更新目录或页码。每个渲染空白页必须标记为 `intentional_recto_blank`、`removable_trailing_blank` 或 `unexpected_blank`，并记录保留、删除或阻塞的依据。

## 审阅标注稿

优先将批注锚定到受影响样式的第一个非空正文段落。批注必须包含规则编号、格式变化摘要和来源编号。页面或全局规则无法精确锚定时，将批注放在第一个可用正文段落，并在报告中标注为全局变化。

不要通过新增可见正文来制作说明。不要用高亮覆盖出版规范要求的底色或字体颜色。

## 内容一致性

比较时排除字段结果显示文本，并仅对明确批准的字段标记、无歧义标题前缀和逐项确认的题注标识替换进行规范化。题注替换必须证明标签、分隔符和标题不变。其余正文、表格文字、脚注、尾注、长引文、引用、答案、索引词和公式语义必须一致。

审计失败时不得交付为通过。若用户批准的派生字段在 Word 打开后更新，只在报告中记录，不将其视为正文修改。任何公式栅格化、嵌入对象变化或媒体替换均视为失败。

## Finalization evidence

Record the input, profile, structure-map, and final-output SHA-256 values; field-cache status before and after finalization; updater backend and software version; updated field types and counts; repagination, save, reopen, and cache-verification status; optional target PDF; and content/protected-object audit outcomes. Do not include manuscript text in the status file.

Finalization evidence must also record effective-font integrity before and after refresh. A target-software save is rejected when it reintroduces a conflicting theme font or changes an approved effective font.

`deferred` is a limitation, not a refresh success. State who approved it and require the target application to update fields before final visual QA. If the renderer and target software differ, report both and set target layout to unverified until that application is checked.

`refreshed_target_word` requires a successful Microsoft Word backend plus core integrity audits. `target_pdf_ready_for_visual_qa` only means Word exported a PDF. Promote to `target_layout_verified` only after every PDF page has been inspected; never infer the physical page count from PAGE-field count.
