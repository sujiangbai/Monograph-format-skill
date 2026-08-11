# 专著要素与规则映射

## 要素清单

检查页面与分节、前置部分、目录、四级以上标题、正文混排、页眉页码、图表公式、题注与交叉引用、脚注尾注、长引文、参考文献、索引、答案、附录和教学提示框。

## 自动选择器

V1.1 可自动执行：

- `document: all` 或 `section_role: all`：现有各节页面属性。
- `style_name: <Word style name>`：指定 Word 样式。
- `paragraph_role`：`body`、`body_text`、`title`、`subtitle`、`chapter_title`、`level_2_section` 至 `level_4_section`、`long_quote`、`heading1` 至 `heading9`。
- `caption_role: all`、`table_role: all`、`bibliography_role: all`。
- `field_role`：批准的字段标记、多级标题编号和无歧义手工前缀迁移。
- `equation_role`：可编辑公式保护和公式图片阻塞策略。

`note_role`、`index_role`、`callout_role`、答案、自定义语义对象、宽表分节、续表断点、文本框和浮动对象在 V1.1 中仍使用 `manual_review`，除非已有更窄且经过测试的自动规则。

## 样式属性

- 字体：`font_name`、`font_name_ascii`、`font_name_east_asia`、`font_name_complex_script`、`font_size_pt`、`bold`、`italic`、`color_hex`。
- 段落：`alignment`、`space_before_pt`、`space_after_pt`、`line_spacing`、`line_spacing_rule`、`line_spacing_pt`。
- 缩进：`first_line_indent_pt`、`first_line_indent_chars`、`left_indent_pt`、`right_indent_pt`。
- 分页：`keep_with_next`、`keep_together`、`page_break_before`、`widow_control`。

`line_spacing_rule` 为 `exact` 时必须同时提供 `line_spacing_pt`；字符缩进以 1/100 字符写入 OOXML。

题注规则还可声明 `numbering_mode`、`preserve_identifier`、`domain_context`、`allow_automatic_renumbering` 和 `preserve_table_cell_caption_position`。默认使用 `manual_text`、保留编号、禁止自动重编号并保留表格首行题注位置。`seq_field` 只有在当前调用者明确要求且结构图逐项批准时才能执行。

## 页面与分节属性

- 固定值：`page_width_mm`、`page_height_mm`、`orientation`、四侧 `margin_*_mm`、`gutter_mm`。
- 保持纸张：`page_size_policy: preserve`。
- 镜像比例：`mirror_margins`、`margin_inner_ratio`、`margin_outer_ratio`、`margin_top_ratio`、`margin_bottom_ratio`。
- 页眉页脚：`header_distance_ratio`、`footer_distance_ratio`、`different_first_page_header_footer`、`odd_and_even_pages_header_footer`。

比例值基于当前节的纸张宽高计算。扫描页只能支持候选比例，不能直接证明准确毫米值。

## 表格属性

- `table_style`
- `alignment`
- `repeat_header_row`
- `prevent_row_split`
- `available_width_percent`、`preferred_column_widths_percent`、`allow_autofit`
- `cell_margins_mm`、`vertical_alignment`
- `border_preset`：`preserve`、`three_line`、`full_grid` 或 `technical_textbook`
- `column_roles`、`column_alignments`
- 未批准列角色时保留原单元格段落对齐；不得把正文首行缩进泄漏到表格。仅由正文样式继承的首行缩进应归零，原单元格直接设置的有意缩进保持不变。
- `header_bold`、`header_shading_hex`
- `font_name_east_asia`、`font_name_ascii`、`font_size_pt`、`line_spacing_pt`

`prevent_row_split` 适用于普通行。视觉属性只作用于结构映射逐表批准的数据表；逐列角色不明确时不得自动对齐。超高行、复杂合并、可见控制标记、续表标签、断点和横向分节必须先盘点并通过 QA。

`technical_textbook` 取消全部底纹、左右外边框和数据行内部横线。表题行上方无线；表题下方/首层表头上方及末行下方使用 1.0 pt，表头层级线、完整表头下方线和内部竖线使用 0.5 pt。多行表头的层级线只覆盖实际拆分的列，不得切穿纵向合并单元格。

获批 `front_matter` 将整本书书名置于独立、不显示且不计入页码的书名页；下一页插入居中加粗的“目录”，目录页码从 1 开始。获批 `block_spacing` 在图表完整内容块后插入一个真实空段落；目标软件重新分页后，若该空段落落在新页页首，则删除。空段落属于批准的派生结构，重复运行不得叠加。

镜像边距与连续物理页可能发生冲突：Microsoft Word 在镜像版式的新分节从 1 重启时可能自动补齐奇偶空白页。若调用者同时批准镜像边距、书名页后立即开始目录、目录后立即开始正文，以及目录和正文分别从可见页码 1 开始，则 Word 适配器应按下一内容页的物理奇偶性选择分节起始类型；物理起始页为偶数时使用可编辑的 `{ = { PAGE } - 1 }` 字段，并在最终更新目录后校正受影响的 `PAGEREF` 结果。skill 必须自动完成刷新、锁定交付缓存并逐页验证，不要求调用者手动更新；分页变化后应重新运行 skill。不得静默保留补齐空白页或关闭镜像边距。

## 字段属性

- `update_on_open`
- `mark_fields_dirty`
- `convert_explicit_markers`
- `rebuild_heading_numbering`
- `heading_levels`
- `strip_manual_heading_prefixes`

支持的显式标记为 `[[TOC]]`、`[[PAGE]]`、`[[SEQ:name]]`、`[[REF:name]]` 和 `[[PAGEREF:name]]`。不得根据相似文本猜测静态目录、题注或交叉引用。

目录与正文的物理页码分区属于结构映射 1.4，不是普通字段属性。页码序列必须由目标软件重新分页验证；`PAGE` 字段数量不等于文档页数。

## 公式属性

- `require_editable_equations`
- `preserve_editable_objects`
- `block_formula_images`

可编辑对象包括 OMML 以及仍可由相应编辑器打开的 MathType/OLE。LaTeX 只能转换为可编辑 OMML。任何公式图片候选、旧版 Equation Editor 对象或对象哈希变化都必须阻塞或进入 QA，不得静默栅格化或重建。

任何未列出的自动属性必须改为 `manual_review`，除非后续不可变版本方案明确增加支持并有测试覆盖。
