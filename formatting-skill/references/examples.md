# Few-Shot 示例集

## 示例 1: 简单全文字体+行距

**用户输入：** "宋体小四，1.5倍行距"

```json
{
  "plan": "设置全文字体为宋体小四，行距1.5倍",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "document" },
        "zhFont": "SimSun",
        "fontSize": 12
      },
      "description": "设置全文字体为宋体小四(12pt)"
    },
    {
      "id": "cmd_2",
      "skill": "paragraph.line_spacing.set",
      "params": {
        "target": { "type": "all_paragraphs" },
        "mode": "multiple",
        "value": 1.5
      },
      "description": "设置全文行距为1.5倍"
    }
  ]
}
```

## 示例 2: 标题+正文分角色排版

**用户输入：** "标题居中加粗黑体三号，正文首行缩进2字符"

```json
{
  "plan": "设置一级标题为黑体三号加粗居中，正文首行缩进2字符",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "heading_1" },
        "zhFont": "SimHei",
        "fontSize": 16,
        "bold": true
      },
      "description": "设置一级标题为黑体三号加粗"
    },
    {
      "id": "cmd_2",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "role", "role": "heading_1" },
        "alignment": "center"
      },
      "description": "设置一级标题居中"
    },
    {
      "id": "cmd_3",
      "skill": "paragraph.indent.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "firstLineChars": 2
      },
      "description": "设置正文首行缩进2字符"
    }
  ]
}
```

## 示例 3: 页面设置

**用户输入：** "A4纸横向，页边距上下2厘米左右2.5厘米，装订线1厘米"

```json
{
  "plan": "设置A4纸横向，页边距上下2cm左右2.5cm，装订线1cm",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "page.size.set",
      "params": { "paperSize": "A4" },
      "description": "设置纸张为A4"
    },
    {
      "id": "cmd_2",
      "skill": "page.orientation.set",
      "params": { "orientation": "landscape" },
      "description": "设置纸张方向为横向"
    },
    {
      "id": "cmd_3",
      "skill": "page.margins.set",
      "params": {
        "top": 56.7,
        "bottom": 56.7,
        "left": 70.9,
        "right": 70.9,
        "gutter": 28.35
      },
      "description": "设置页边距上下2cm、左右2.5cm、装订线1cm"
    }
  ]
}
```

## 示例 4: 页码规则

**用户输入：** "前言用罗马数字页码，正文从第1页重新编号"

```json
{
  "plan": "设置前言节罗马数字页码，正文节阿拉伯数字从1开始",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "page_number.set",
      "params": {
        "section": 1,
        "format": "roman_lower",
        "position": "bottom_center"
      },
      "description": "设置前言节页码为小写罗马数字，底部居中"
    },
    {
      "id": "cmd_2",
      "skill": "page_number.restart",
      "params": {
        "section": 2,
        "startFrom": 1,
        "format": "arabic"
      },
      "description": "正文节页码从1重新开始，阿拉伯数字"
    }
  ]
}
```

## 示例 5: 选区操作

**用户输入：** "把选中的文字加粗变红"

```json
{
  "plan": "将当前选区文字设为加粗、红色",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set_bold",
      "params": {
        "target": { "type": "selection" },
        "bold": true
      },
      "description": "选区文字加粗"
    },
    {
      "id": "cmd_2",
      "skill": "text.font.set_color",
      "params": {
        "target": { "type": "selection" },
        "color": "#FF0000"
      },
      "description": "选区文字设为红色"
    }
  ]
}
```

## 示例 6: 段落间距

**用户输入：** "标题段前段后各12磅，正文段前段后0"

```json
{
  "plan": "设置标题段前段后12磅，正文段前段后0",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "paragraph.spacing.set",
      "params": {
        "target": { "type": "role", "role": "heading_1" },
        "spaceBefore": 12,
        "spaceAfter": 12
      },
      "description": "一级标题段前段后各12磅"
    },
    {
      "id": "cmd_2",
      "skill": "paragraph.spacing.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "spaceBefore": 0,
        "spaceAfter": 0
      },
      "description": "正文段前段后设为0"
    }
  ]
}
```

## 示例 7: 固定行距

**用户输入：** "全文固定行距20磅"

```json
{
  "plan": "设置全文固定行距20磅",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "paragraph.line_spacing.set",
      "params": {
        "target": { "type": "all_paragraphs" },
        "mode": "exact",
        "value": 20
      },
      "description": "设置全文固定行距20磅"
    }
  ]
}
```

## 示例 8: 中英文字体分开设置

**用户输入：** "正文中文宋体英文Times New Roman"

```json
{
  "plan": "设置正文中文字体宋体、英文字体Times New Roman",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "zhFont": "SimSun",
        "enFont": "Times New Roman"
      },
      "description": "正文中文宋体、英文Times New Roman"
    }
  ]
}
```

## 示例 9: 页眉页脚

**用户输入：** "添加页眉显示论文标题，宋体小五号"

```json
{
  "plan": "设置页眉为论文标题，宋体小五号",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "header.set",
      "params": {
        "text": "论文标题",
        "alignment": "center",
        "fontName": "SimSun",
        "fontSize": 9
      },
      "description": "设置页眉为论文标题，宋体小五居中"
    }
  ]
}
```

## 示例 10: 标题分页

**用户输入：** "每一级标题另起一页"

```json
{
  "plan": "设置一级标题段前分页",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "section.page_break_before",
      "params": {
        "target": { "type": "role", "role": "heading_1" }
      },
      "description": "一级标题另起新页"
    }
  ]
}
```

## 示例 11: 图表标题格式

**用户输入：** "图标题黑体五号居中，表标题也是"

```json
{
  "plan": "设置图标题和表标题为黑体五号居中",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "figure_caption" },
        "zhFont": "SimHei",
        "fontSize": 10.5
      },
      "description": "图标题设为黑体五号"
    },
    {
      "id": "cmd_2",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "role", "role": "figure_caption" },
        "alignment": "center"
      },
      "description": "图标题居中"
    },
    {
      "id": "cmd_3",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "table_caption" },
        "zhFont": "SimHei",
        "fontSize": 10.5
      },
      "description": "表标题设为黑体五号"
    },
    {
      "id": "cmd_4",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "role", "role": "table_caption" },
        "alignment": "center"
      },
      "description": "表标题居中"
    }
  ]
}
```

## 示例 12: 参考文献格式

**用户输入：** "参考文献宋体小五，悬挂缩进2字符，行距固定16磅"

```json
{
  "plan": "设置参考文献宋体小五、悬挂缩进2字符、固定行距16磅",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "references" },
        "zhFont": "SimSun",
        "fontSize": 9
      },
      "description": "参考文献字体宋体小五"
    },
    {
      "id": "cmd_2",
      "skill": "paragraph.indent.set",
      "params": {
        "target": { "type": "role", "role": "references" },
        "hanging": 2
      },
      "description": "参考文献悬挂缩进2字符"
    },
    {
      "id": "cmd_3",
      "skill": "paragraph.line_spacing.set",
      "params": {
        "target": { "type": "role", "role": "references" },
        "mode": "exact",
        "value": 16
      },
      "description": "参考文献固定行距16磅"
    }
  ]
}
```

## 示例 13: 学术论文全套排版

**用户输入：** "按照毕业论文格式排版"

```json
{
  "plan": "按毕业论文标准格式排版：A4纸、标准页边距、标题黑体、正文宋体小四1.5倍行距首行缩进",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "page.size.set",
      "params": { "paperSize": "A4" },
      "description": "设置A4纸张"
    },
    {
      "id": "cmd_2",
      "skill": "page.margins.set",
      "params": { "top": 72, "bottom": 72, "left": 90, "right": 72 },
      "description": "设置页边距上下2.54cm、左3.17cm、右2.54cm"
    },
    {
      "id": "cmd_3",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "title" },
        "zhFont": "SimHei",
        "fontSize": 22,
        "bold": true
      },
      "description": "论文题目黑体二号加粗"
    },
    {
      "id": "cmd_4",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "role", "role": "title" },
        "alignment": "center"
      },
      "description": "论文题目居中"
    },
    {
      "id": "cmd_5",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "heading_1" },
        "zhFont": "SimHei",
        "fontSize": 16,
        "bold": true
      },
      "description": "一级标题黑体三号加粗"
    },
    {
      "id": "cmd_6",
      "skill": "paragraph.spacing.set",
      "params": {
        "target": { "type": "role", "role": "heading_1" },
        "spaceBefore": 24,
        "spaceAfter": 12
      },
      "description": "一级标题段前24磅段后12磅"
    },
    {
      "id": "cmd_7",
      "skill": "section.page_break_before",
      "params": {
        "target": { "type": "role", "role": "heading_1" }
      },
      "description": "一级标题另起一页"
    },
    {
      "id": "cmd_8",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "heading_2" },
        "zhFont": "SimHei",
        "fontSize": 14,
        "bold": true
      },
      "description": "二级标题黑体四号加粗"
    },
    {
      "id": "cmd_9",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "zhFont": "SimSun",
        "enFont": "Times New Roman",
        "fontSize": 12
      },
      "description": "正文宋体/Times New Roman小四"
    },
    {
      "id": "cmd_10",
      "skill": "paragraph.line_spacing.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "mode": "multiple",
        "value": 1.5
      },
      "description": "正文1.5倍行距"
    },
    {
      "id": "cmd_11",
      "skill": "paragraph.indent.set",
      "params": {
        "target": { "type": "role", "role": "body" },
        "firstLineChars": 2
      },
      "description": "正文首行缩进2字符"
    },
    {
      "id": "cmd_12",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "figure_caption" },
        "zhFont": "SimHei",
        "fontSize": 10.5
      },
      "description": "图标题黑体五号"
    },
    {
      "id": "cmd_13",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "role", "role": "figure_caption" },
        "alignment": "center"
      },
      "description": "图标题居中"
    },
    {
      "id": "cmd_14",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "references" },
        "zhFont": "SimSun",
        "fontSize": 10.5
      },
      "description": "参考文献宋体五号"
    },
    {
      "id": "cmd_15",
      "skill": "paragraph.spacing.set",
      "params": {
        "target": { "type": "role", "role": "references" },
        "spaceBefore": 0,
        "spaceAfter": 0
      },
      "description": "参考文献段前段后0"
    }
  ]
}
```

## 示例 14: 无法实现的需求

**用户输入：** "把文档翻译成英文"

```json
{
  "plan": "无法执行: 翻译功能不在排版技能范围内，仅支持文档格式调整操作",
  "commands": []
}
```

## 示例 15: 二级三级标题

**用户输入：** "二级标题黑体四号，三级标题黑体小四"

```json
{
  "plan": "设置二级标题黑体四号、三级标题黑体小四",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "heading_2" },
        "zhFont": "SimHei",
        "fontSize": 14
      },
      "description": "二级标题黑体四号"
    },
    {
      "id": "cmd_2",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "heading_3" },
        "zhFont": "SimHei",
        "fontSize": 12
      },
      "description": "三级标题黑体小四"
    }
  ]
}
```

## 示例 16: 两端对齐

**用户输入：** "全文两端对齐"

```json
{
  "plan": "设置全文段落两端对齐",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "paragraph.alignment.set",
      "params": {
        "target": { "type": "all_paragraphs" },
        "alignment": "justify"
      },
      "description": "全文两端对齐"
    }
  ]
}
```

## 示例 17: 指定段落操作

**用户输入：** "第3段到第5段字体改成楷体"

```json
{
  "plan": "将第3到第5段字体改为楷体",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set_zh",
      "params": {
        "target": { "type": "paragraph_range", "from": 3, "to": 5 },
        "fontName": "KaiTi"
      },
      "description": "第3-5段中文字体改为楷体"
    }
  ]
}
```

## 示例 18: 搜索文本格式化

**用户输入：** "把所有的'重要'两个字加粗变红"

```json
{
  "plan": "搜索所有'重要'文本并设为加粗红色",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "search", "text": "重要", "occurrence": 0 },
        "bold": true,
        "color": "#FF0000"
      },
      "description": "所有'重要'文字加粗变红"
    }
  ]
}
```

## 示例 19: 摘要格式

**用户输入：** "摘要楷体小四，标题加粗"

```json
{
  "plan": "设置摘要段落楷体小四，摘要标题加粗",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.font.set",
      "params": {
        "target": { "type": "role", "role": "abstract" },
        "zhFont": "KaiTi",
        "fontSize": 12
      },
      "description": "摘要段落楷体小四"
    }
  ]
}
```

## 示例 20: 清除格式

**用户输入：** "清除选中内容的所有格式"

```json
{
  "plan": "清除当前选区的所有格式",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "text.clear_formatting",
      "params": {
        "target": { "type": "selection" }
      },
      "description": "清除选区所有格式"
    }
  ]
}
```

## 示例 21: 三线表样式

**用户输入：** "第一个表格设为三线表，表头加粗居中"

```json
{
  "plan": "第一个表格设为三线表样式，表头行加粗居中",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "table.borders.set",
      "params": { "tableIndex": 1, "style": "three_line" },
      "description": "第一个表格设为三线表"
    },
    {
      "id": "cmd_2",
      "skill": "table.font.set",
      "params": { "tableIndex": 1, "row": 1, "bold": true },
      "description": "表头行加粗"
    },
    {
      "id": "cmd_3",
      "skill": "table.cell_alignment.set",
      "params": { "tableIndex": 1, "row": 1, "horizontal": "center", "vertical": "center" },
      "description": "表头行居中"
    }
  ]
}
```

## 示例 22: 表格字体+适应窗口

**用户输入：** "所有表格宋体五号，自动适应窗口"

```json
{
  "plan": "所有表格宋体五号、自动适应窗口宽度",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "table.font.set",
      "params": { "tableIndex": 1, "zhFont": "SimSun", "fontSize": 10.5 },
      "description": "第一个表格宋体五号"
    },
    {
      "id": "cmd_2",
      "skill": "table.autofit.set",
      "params": { "tableIndex": 1, "mode": "window" },
      "description": "第一个表格自动适应窗口"
    }
  ]
}
```

## 示例 23: 表格底纹

**用户输入：** "表1的第一行背景设为浅蓝色"

```json
{
  "plan": "表1第一行背景设为浅蓝色",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "table.shading.set",
      "params": { "tableIndex": 1, "row": 1, "color": "#DCE6F1" },
      "description": "表1第一行背景浅蓝色"
    }
  ]
}
```

## 示例 24: 修改标题样式

**用户输入：** "修改标题1样式为黑体三号加粗，段前24磅段后12磅"

```json
{
  "plan": "修改标题1样式属性",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "style.modify",
      "params": {
        "styleName": "标题 1",
        "zhFont": "SimHei",
        "fontSize": 16,
        "bold": true,
        "spaceBefore": 24,
        "spaceAfter": 12
      },
      "description": "修改标题1样式为黑体三号加粗，段前24磅段后12磅"
    }
  ]
}
```

## 示例 25: 插入目录

**用户输入：** "在文档开头插入三级目录"

```json
{
  "plan": "在文档开头插入三级目录",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "toc.insert",
      "params": { "levels": 3 },
      "description": "插入三级目录"
    }
  ]
}
```

## 示例 26: 添加斜向水印

**用户输入：** "加一个'机密'水印"

```json
{
  "plan": "添加斜向文字水印'机密'",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "watermark.text.set",
      "params": {
        "text": "机密",
        "fontName": "SimSun",
        "fontSize": 54,
        "color": "#C0C0C0",
        "layout": "diagonal"
      },
      "description": "添加斜向水印'机密'"
    }
  ]
}
```

## 示例 27: 表头跨页重复+全表居中

**用户输入：** "第一个表格居中显示，表头跨页重复"

```json
{
  "plan": "第一个表格居中，表头跨页重复",
  "commands": [
    {
      "id": "cmd_1",
      "skill": "table.alignment.set",
      "params": { "tableIndex": 1, "alignment": "center" },
      "description": "表格居中"
    },
    {
      "id": "cmd_2",
      "skill": "table.header_row.set",
      "params": { "tableIndex": 1, "repeat": true },
      "description": "表头跨页重复"
    }
  ]
}
```
