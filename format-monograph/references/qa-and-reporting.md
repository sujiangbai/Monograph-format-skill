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
7. 冲突、QA 决定、字段刷新、兼容性和其他限制。
8. 渲染页数、逐页检查状态和发现的问题。
9. 最终结论：通过、带限制通过或失败。

## 审阅标注稿

优先将批注锚定到受影响样式的第一个非空正文段落。批注必须包含规则编号、格式变化摘要和来源编号。页面或全局规则无法精确锚定时，将批注放在第一个可用正文段落，并在报告中标注为全局变化。

不要通过新增可见正文来制作说明。不要用高亮覆盖出版规范要求的底色或字体颜色。

## 内容一致性

比较时排除字段结果显示文本，并仅对明确批准的字段标记、无歧义标题前缀和逐项确认的题注标识替换进行规范化。题注替换必须证明标签、分隔符和标题不变。其余正文、表格文字、脚注、尾注、长引文、引用、答案、索引词和公式语义必须一致。

审计失败时不得交付为通过。若用户批准的派生字段在 Word 打开后更新，只在报告中记录，不将其视为正文修改。任何公式栅格化、嵌入对象变化或媒体替换均视为失败。

## Finalization evidence

Record the input, profile, structure-map, and final-output SHA-256 values; field-cache status before and after finalization; the updater backend; and content/protected-object audit outcomes. Do not include manuscript text in the status file.

`deferred` is a limitation, not a refresh success. State who approved it and require the target application to update fields before final visual QA. If the renderer and target software differ, report both and set target layout to unverified until that application is checked.
