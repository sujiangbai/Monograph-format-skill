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
9. 每个获批图片的放置类别、原/新显示尺寸、缩放比例、有效 DPI、媒体哈希，以及锚点、容器和对象顺序保护结果；未获批图片及原因。
10. 图片可见性统计：嵌入/浮动/表内分类、固定行距或固定行高候选、获批修复、保持原状和阻塞 QA 数，以及显示尺寸未变证明。
11. 目录来源模式、批准源数量、纯文本 `TC` 来源数量、结果条目数量，以及 `verified_text_only`、`deferred` 或 `rejected` 状态；不得记录标题正文。
12. 冲突、QA 决定、字段刷新后端、目标软件版本、兼容性和其他限制。
13. 渲染页数、PDF 来源、逐页检查状态和发现的问题。
14. 最终结论：通过、带限制通过或失败。

逐表视觉 QA 必须检查短标签、编号、数值与单位是否出现可避免的孤立换行。不得让正文首行缩进泄漏到表格单元格；若短内容因列宽、单元格边距、段落缩进或错误对齐被拆行，应先恢复原有合理对齐并消除非语义缩进，再考虑调整列宽。不得通过改写文字掩盖排版问题。

逐页检查还必须确认书名页无页码、书名与目录分属不同页面、书名符合批准的字体/字号/字重、至少 33 pt 行距、不截字，并在页面内水平和垂直居中；目录题名为批准文字且居中加粗。逐级比较多级列表编号与对应 Heading 样式的中西文字体、字号和字重，并确认 Heading 1-4 首行缩进为 0。检查每个图表后的批准空段落：同页后续内容前必须恰有一个，跨页时新页顶部不得保留。逐表报告底纹、主线/次线磅值、左右外边框、数据行横线，以及多行表头的局部分隔线是否避开纵向合并单元格。

逐图检查必须将原稿与格式稿按锚点逐一对应，确认图片仍在同一段落或单元格、承载表格仍在相同正文对象位置、前后相邻正文顺序未改变、媒体内容与裁剪状态未变化、宽高比未失真且没有溢出或遮挡。同一图版行还要检查显示高度一致。不得把自然分页变化误报为对象移动；对象锚点、容器或顺序无法证明保持时，审计必须失败。

逐图还要检查图片段落是否因固定行距或表格固定行高而裁切。获批修复只能把独立嵌入图片段落改为自动单倍行距，或把证据明确的简单表格行从 `exact` 放宽为 `atLeast`；尺寸、裁剪、锚点、环绕、容器和顺序必须全部不变。逐目录检查不得仅凭“看起来正常”：字段结果必须与批准源逐项匹配，并确认目录区域不含图片、VML、OLE、文本框、表格、外部关系、空项或额外项。

## V0.2.6 确定性字体与分页证据

对每条获批的自动字体规则，报告目标字体、样式中声明的显式字体、原主题引用、沿继承链解析出的实际生效字体及其来源。正文、标题、目录、题注、表格文字、脚注/尾注、长引文、参考文献和教学框中凡已获批自动处理的对象，都必须按同一方法检查。检测到等线、等线 Light 或其他主题字体并不自动表示错误；只有它与当前任务批准字体不一致时才失败。未获批对象保持原状并列入未覆盖范围。

对每个页码页脚报告 `PAGE` 字段数量、奇偶页位置、是否包含非页码内容，以及本次运行是否执行了重复字段规范化、获批的静态数字转换或 `{ = { PAGE } - 1 }` 显示偏移。报告物理分节起始、各节显示偏移、校正的 `PAGEREF` 数量和 TOC 锁定状态；这些项目全部成功时不得再要求调用者手动更新目录或页码。每个渲染空白页必须标记为 `intentional_recto_blank`、`removable_trailing_blank` 或 `unexpected_blank`，并记录保留、删除或阻塞的依据。

## 审阅标注稿

优先将批注锚定到受影响样式的第一个非空正文段落。批注必须包含规则编号、格式变化摘要和来源编号。页面或全局规则无法精确锚定时，将批注放在第一个可用正文段落，并在报告中标注为全局变化。

不要通过新增可见正文来制作说明。不要用高亮覆盖出版规范要求的底色或字体颜色。

## 内容一致性

比较时排除字段结果显示文本，并仅对明确批准的字段标记、无歧义标题前缀和逐项确认的题注标识替换进行规范化。题注替换必须证明标签、分隔符和标题不变。其余正文、表格文字、脚注、尾注、长引文、引用、答案、索引词和公式语义必须一致。

审计失败时不得交付为通过。若用户批准的派生字段在 Word 打开后更新，只在报告中记录，不将其视为正文修改。任何公式栅格化、嵌入对象变化或媒体替换均视为失败。

报告应分别列出数据表边框模型、语义分隔行、图版布局表、无编号图名和表格单元格空白清理。图版标签和无编号图名必须证明文字未变化且未添加编号；空白清理必须证明只删除了获批的前导空段落。旧单元格边框未被清除、图版表仍有可见边框或环绕、汇总行缺少获批分隔线，均视为结构审计失败。

## Finalization evidence

Record the input, profile, structure-map, and final-output SHA-256 values; field-cache status before and after finalization; updater backend and software version; approved, matched, updated, and rejected field counts; discarded backend-difference categories; each read-only layout measurement; core section-start, page-display-offset, and page-top-spacer adjustments; calculation and no-save verification page counts; target PDF; and content/protected-object audit outcomes. For a completed Word path, persist the verification PDF and bind both it and the finalized DOCX by resolved path, SHA-256, size, and page count where applicable. Store a versioned allowlisted finalization-gate summary in state: top-level pass status, all three integrity results, finalized workflow stage, source/formatted/profile/map/output hashes, output/target paths and layout status, artifact identities, and complete field evidence. Rebuild and compare it during resume, verification, and status. Also bind and revalidate the final audit, render manifest, visual manifest, and versioned target/renderer request identity before a verification resume or a `status` report preserves `final_ready`; repeat the Word/PDF/render/visual page-count and completion checks rather than trusting the stored status. A legitimate new input or request is a cache miss and reruns the stage before replacing old bindings; an unchanged input key with altered output is invalid. Invalid evidence is stale and must not remain final-ready, and `status` must not regenerate it. These checks detect single-sided local file or JSON changes but are not a signature or application authentication: a subject that can coherently alter every unsigned local file and request record is inside the explicit local trust boundary. Do not include manuscript or field-result text in the status file.

Finalization evidence must also record effective-font integrity before and after refresh. A target-software save is rejected when it reintroduces a conflicting theme font or changes an approved effective font.

A real successful finalization execution always clears the previous verify
stage and all audit/render/visual bindings and derived visual/page state before
saving the new candidate, even if the finalized DOCX and persistent PDF bytes
did not change. The next verification resume must run audit and render again and
validate the requested visual manifest before restoring final-ready. Only a
true, revalidated finalization cache hit may preserve the existing verification
cache and final-ready state.

Treat the finalizer as successfully producing evidence only after process exit
zero and validation of the shared versioned finalization-evidence shape. Require
all declared cache, backend, completion, integrity, workflow, target, output, and
artifact-binding fields with their exact object/scalar types and supported
enums; reject missing fields, unknown versions, booleans used as integers, empty
paths, and malformed identities before changing run state. This is a producer
schema gate, not the final-ready business gate: structurally valid deferred and
LibreOffice non-final evidence remains candidate evidence.

`field_backend` is a versioned closed canonical projection used by completion
and final-ready gates. It contains only allowlisted finite enums, booleans,
counts, and explicitly shaped selective-writeback, read-only-verification, and
fallback-failure facts; unknown keys, invalid operations/counts, and non-finite
numbers are rejected. Never copy a raw external or LibreOffice response into
that object. Detailed backend diagnostics belong in a separate versioned JSON
audit sidecar. The finalization JSON stores only the sidecar's resolved path,
SHA-256, and byte size, and resume, verify, and status rehash it before trusting
the evidence. The sidecar accepts standard JSON only and enforces bounded byte,
depth, node, and string limits; it must be written atomically. A command without
`--status-output` reports `backend_audit.status=not_persisted` and never embeds
the raw response. The sidecar remains diagnostic, not business-gate input, and
must not contain manuscript or field-result text.

Resolve all finalization inputs and persistent outputs before any filesystem
mutation. The DOCX, PDF, status, and derived backend-audit outputs must be
pairwise distinct both lexically and after symlink resolution, must not resolve
to any input, must not themselves be symlinks, and must share one parent
directory. Reject NUL/control-character and unresolvable paths. Existing targets
require `--force` and must remain unchanged until a fully validated staging set
is ready. Publish status last as the commit marker. If a multi-file replace
fails, restore the captured old targets; retain staging rather than deleting an
unrestored backup when rollback cannot complete.

The orchestrator does not accept a fresh sidecar path from the candidate JSON or
old state. It independently derives the expected path from its fixed
finalization-status path, requires an exact versioned binding and ordinary
non-symlink file, and checks size, SHA-256, standard JSON, limits, and root schema
before invalidating prior verify evidence. Every evidence `version` is an exact
integer: booleans, floats, strings, and unknown integers are unsupported.

Use only the allowlisted target IDs `microsoft_word`, `libreoffice`, and
`unsupported`; backend software, finalization, persistent Word PDF, verification
request, and render manifest must resolve to the same applicable ID. The
canonical gate includes the stored `field_completion.evidence_validation` and
requires it to equal a fresh canonical recalculation with `status=pass` and no
errors. Bind a renderer only when the stage actually invokes it. When an
external updater is used, record its versioned, explicitly unproven dependency
audit and `cache_reusable=false`. Record the shared-parser argv and fixed
arguments; the PATH-resolved executable; plain, option-value, and
option-assignment files; response files; deterministic directory manifests; and
diagnostic direct Python script or local `-m` identities by resolved path, type,
SHA-256, and size where applicable. Resolve relative paths from the same execution
cwd used by the finalizer and never follow directory symlinks. None of these
observations proves runtime hermeticity: native programs, wrappers, Python
reflection/site customization, environment variables, PATH subcommands, and
other hidden inputs remain possible. Consequently every external finalize resume
must rerun the updater even when all recorded entities are unchanged. Fresh valid
evidence remains valid and status remains read-only. This is unsigned argv and
explicit-entity consistency—not a signature, application authentication, or
general dependency analysis. Reopening external finalize caching requires a
separately approved, auditable, runtime-enforced hermetic adapter.

When `toc_source` is approved, finalization evidence must record only the source
mode, source count, approved TC source-field count, result count, and
`toc_result_status`; do not record heading or TOC text. `verified_text_only`
requires one matching result per approved source in order and no non-text
payload. Any mismatch rejects the disposable backend result before writeback.

`deferred` is a limitation, not a refresh success. State who approved it and require the target application to update fields before final visual QA. If the renderer and target software differ, report both and set target layout to unverified until that application is checked.

`selective_verified` requires successful Microsoft Word field calculation, core selective writeback, integrity audits, and a second no-save Word verification with matching page count. `target_pdf_ready_for_visual_qa` only means Word exported a PDF. Promote to `target_layout_verified` only after every PDF page has been inspected; never infer the physical page count from PAGE-field count.

## V0.3.0 whole-book run evidence

For an orchestrated whole-book run, also report the `run-state.json` status,
cache hits, duration of each stage, grouped QA decisions, object-level
exceptions, frozen scopes, field-result writeback parts, rendered page count,
target-layout evidence, and visual-QA manifest status. Do not copy local paths,
manuscript text, inventories, or rendered pages into a public report or CI log.

`prepared` means only that inventory is complete. `blocked_qa` may include a
valid partial candidate while uncertain objects remain unchanged.
`candidate_ready` must list the remaining field, target-layout, or visual gate.
Only `final_ready` is an unrestricted whole-book completion claim.
