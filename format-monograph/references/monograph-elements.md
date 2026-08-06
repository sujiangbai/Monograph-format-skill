# 专著要素与规则映射

## 要素清单

检查以下要素是否存在并分别建立规则：

- 页面尺寸、装订线、页边距、分节、横竖版和起始页。
- 封面、扉页、版权页、序言、前言、摘要、目录和其他前置部分。
- 章、节和小节标题的字体、字号、间距、对齐、分页及编号。
- 正文、中西文混排、首行缩进、行距、段前段后、孤行和段中分页。
- 页眉、页脚、奇偶页、首页差异和页码体系。
- 图片、表格、公式、题注、编号和交叉引用。
- 脚注、尾注、参考文献、索引和附录。

## 自动选择器

V1 自动执行以下选择器：

- `document: all`：页面尺寸和默认分节属性。
- `section_role: all`：对所有现有节应用页面属性。
- `style_name: <Word style name>`：修改指定 Word 样式定义。
- `paragraph_role`：`body`、`title`、`subtitle`、`heading1` 至 `heading9`。
- `caption_role: all`：映射到 Word `Caption` 样式。
- `table_role: all`：应用表格样式和对齐。
- `bibliography_role: all`：映射到 Word `Bibliography` 样式。

无法可靠定位的封面块、版权页、特定附录、浮动对象、文本框、公式版式或自定义语义对象应使用 `manual_review`。

## 自动属性

样式和段落属性：

- `font_name`
- `font_size_pt`
- `bold`
- `italic`
- `color_hex`
- `alignment`
- `space_before_pt`
- `space_after_pt`
- `line_spacing`
- `first_line_indent_pt`
- `left_indent_pt`
- `right_indent_pt`
- `keep_with_next`
- `keep_together`
- `page_break_before`
- `widow_control`

页面与分节属性：

- `page_width_mm`
- `page_height_mm`
- `orientation`
- `margin_top_mm`
- `margin_bottom_mm`
- `margin_left_mm`
- `margin_right_mm`
- `gutter_mm`
- `different_first_page_header_footer`
- `odd_and_even_pages_header_footer`

表格属性：

- `table_style`
- `alignment`
- `repeat_header_row`

任何未列出的属性必须标记为 `manual_review`，除非后续版本方案明确增加自动支持。
