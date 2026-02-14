# 指令 JSON Schema 完整参考

## 顶层结构

```json
{
  "plan": "简述执行计划（中文）",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "技能ID",
      "params": { ... },
      "description": "操作说明（中文）"
    }
  ]
}
```

## 目标类型 (target)

所有格式操作的 `params` 中包含 `target` 字段，指定操作范围：

```json
// 当前选区
{ "type": "selection" }

// 整篇文档
{ "type": "document" }

// 所有段落
{ "type": "all_paragraphs" }

// 指定段落 (1-based)
{ "type": "paragraph_index", "index": 3 }

// 段落范围
{ "type": "paragraph_range", "from": 1, "to": 5 }

// 搜索文本 (occurrence: 0=全部)
{ "type": "search", "text": "第一章", "occurrence": 0 }

// 标题级别
{ "type": "heading_level", "level": 1 }

// 语义角色
{ "type": "role", "role": "body" }

// 指定节 (1-based)
{ "type": "section_index", "index": 2 }

// 指定表格 (1-based)
{ "type": "table_index", "index": 1 }
```

### role 可选值

`title`, `heading_1`, `heading_2`, `heading_3`, `heading_4`,
`abstract`, `keywords`, `body`, `figure_caption`, `table_caption`,
`references`, `acknowledgment`, `toc`, `footnotes`

---

## 第一层：结构识别技能

| skill | params | 说明 |
|-------|--------|------|
| `detect.title` | — | 识别主标题 |
| `detect.headings` | `levels?` (array, default [1,2,3,4]) | 识别多级标题 |
| `detect.abstract` | — | 识别摘要段落 |
| `detect.keywords` | — | 识别关键词段落 |
| `detect.toc_area` | — | 识别目录区域 |
| `detect.body_paragraphs` | — | 识别正文段落 |
| `detect.figure_captions` | — | 识别图标题 |
| `detect.table_captions` | — | 识别表标题 |
| `detect.images` | — | 识别图片段落 |
| `detect.tables` | — | 识别表格区域 |
| `detect.references` | — | 识别参考文献 |
| `detect.acknowledgment` | — | 识别致谢/附录 |
| `detect.footnotes` | — | 识别脚注 |

---

## 第二层：文字与段落格式技能

### 字体操作

| skill | params |
|-------|--------|
| `text.font.set` | `target`, `zhFont?`, `enFont?`, `fontSize?`, `bold?`, `italic?`, `underline?`, `color?`, `strikethrough?` |
| `text.font.set_zh` | `target`, `fontName` |
| `text.font.set_en` | `target`, `fontName` |
| `text.font.set_size` | `target`, `size` (pt 或中文名如"小四") |
| `text.font.set_bold` | `target`, `bold` (boolean) |
| `text.font.set_color` | `target`, `color` (hex 如 "#FF0000" 或 RGB 整数) |
| `text.clear_formatting` | `target` |
| `text.superscript` | `target`, `enabled` (boolean) |
| `text.subscript` | `target`, `enabled` (boolean) |
| `text.highlight` | `target`, `color` |

### 段落操作

| skill | params |
|-------|--------|
| `paragraph.alignment.set` | `target`, `alignment` ("left"/"center"/"right"/"justify") |
| `paragraph.spacing.set` | `target`, `spaceBefore?` (pt), `spaceAfter?` (pt) |
| `paragraph.line_spacing.set` | `target`, `mode` ("multiple"/"exact"/"atLeast"), `value` |
| `paragraph.indent.set` | `target`, `firstLineChars?`, `firstLinePoints?`, `hanging?`, `left?`, `right?` |
| `paragraph.columns.set` | `target`, `count`, `spacing?` (pt) |
| `paragraph.borders.set` | `target`, `borderType`, `lineWidth?`, `color?` |
| `paragraph.numbering.set` | `target`, `format`, `startAt?` |
| `paragraph.bullets.set` | `target`, `symbol?` |

---

## 第三层：页面/页眉页脚/页码技能

### 页面设置

| skill | params |
|-------|--------|
| `page.margins.set` | `section?` (index), `top`, `bottom`, `left`, `right`, `gutter?` (全部 pt) |
| `page.size.set` | `section?`, `paperSize` ("A3"/"A4"/"A5"/"B5"/"Letter"/"Legal"/"16K"), `width?`, `height?` |
| `page.orientation.set` | `section?`, `orientation` ("portrait"/"landscape") |

### 页眉页脚

| skill | params |
|-------|--------|
| `header.set` | `section?`, `text`, `alignment?`, `fontName?`, `fontSize?` |
| `header.clear` | `section?` |
| `footer.set` | `section?`, `text`, `alignment?`, `fontName?`, `fontSize?` |
| `footer.clear` | `section?` |

### 页码

| skill | params |
|-------|--------|
| `page_number.set` | `section?`, `format` ("arabic"/"roman_lower"/"roman_upper"), `position` ("bottom_center" 等) |
| `page_number.restart` | `section`, `startFrom`, `format?` |

### 分节/分页

| skill | params |
|-------|--------|
| `section.break.insert` | `target`, `breakType` ("nextPage"/"continuous"/"evenPage"/"oddPage") |
| `section.page_break_before` | `target` |
| `section.keep_with_next` | `target` |

---

## 第四层：表格样式技能

| skill | params |
|-------|--------|
| `table.borders.set` | `tableIndex`, `style?` ("all"/"box"/"none"/"three_line"/"dashed"), `lineWidth?`, `color?` |
| `table.cell_alignment.set` | `tableIndex`, `row?`, `col?`, `horizontal?` ("left"/"center"/"right"/"justify"), `vertical?` ("top"/"center"/"bottom") |
| `table.shading.set` | `tableIndex`, `row?`, `col?`, `color` |
| `table.row_height.set` | `tableIndex`, `row?`, `height` (pt), `rule?` ("auto"/"atLeast"/"exact") |
| `table.column_width.set` | `tableIndex`, `col?`, `width` (pt) |
| `table.autofit.set` | `tableIndex`, `mode` ("content"/"window"/"fixed") |
| `table.font.set` | `tableIndex`, `row?`, `zhFont?`, `enFont?`, `fontSize?`, `bold?` |
| `table.header_row.set` | `tableIndex`, `repeat` (boolean) |
| `table.alignment.set` | `tableIndex`, `alignment` ("left"/"center"/"right") |

---

## 第五层：样式/目录技能

| skill | params |
|-------|--------|
| `style.apply` | `target`, `styleName` |
| `style.modify` | `styleName`, `zhFont?`, `enFont?`, `fontSize?`, `bold?`, `italic?`, `alignment?`, `lineSpacingMode?`, `lineSpacingValue?`, `spaceBefore?`, `spaceAfter?`, `firstLineIndent?` |
| `style.create` | `styleName`, `basedOn?`, `zhFont?`, `enFont?`, `fontSize?`, `bold?`, `alignment?` |
| `toc.insert` | `target?`, `levels?` (default 3) |
| `toc.update` | — |
| `toc.remove` | — |
| `watermark.text.set` | `text`, `fontName?`, `fontSize?`, `color?`, `layout?` ("diagonal"/"horizontal") |
| `watermark.clear` | — |
