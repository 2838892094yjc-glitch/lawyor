# 中英文字体名映射表

## 中文字体

| 中文名 | 英文 API 名 | 说明 |
|--------|------------|------|
| 宋体 | SimSun | 最常用正文字体 |
| 黑体 | SimHei | 常用标题字体 |
| 楷体 | KaiTi | 引用/摘要常用 |
| 仿宋 | FangSong | 公文常用 |
| 微软雅黑 | Microsoft YaHei | 现代无衬线 |
| 华文中宋 | STZhongsong | 学术论文正文 |
| 华文仿宋 | STFangsong | 公文正文 |
| 华文楷体 | STKaiti | 引用/注释 |
| 华文宋体 | STSong | macOS 宋体替代 |
| 华文细黑 | STXihei | macOS 黑体替代 |
| 等线 | DengXian | Win10 默认字体 |
| 新宋体 | NSimSun | 等宽宋体 |
| 方正小标宋 | FZXiaoBiaoSong-B05 | 公文标题 |
| 方正仿宋 | FZFangSong-Z02 | 公文正文 |

## 英文字体

| 字体名 | 说明 |
|--------|------|
| Times New Roman | 学术论文标准 |
| Arial | 常用无衬线 |
| Calibri | Office 默认 |
| Cambria | 标题常用 |
| Courier New | 等宽字体 |
| Georgia | 衬线字体 |
| Verdana | 屏幕阅读友好 |

## WPS Font API 属性

| 属性 | 作用 |
|------|------|
| `Font.Name` | 主字体名（通常设中文字体） |
| `Font.NameFarEast` | 中文/CJK 字体 |
| `Font.NameAscii` | ASCII 字符字体（英文/数字） |
| `Font.NameOther` | 其他字符字体 |

## 常见排版组合

| 场景 | 中文字体 | 英文字体 | 字号 |
|------|---------|---------|------|
| 学术论文正文 | 宋体 | Times New Roman | 小四 (12pt) |
| 学术论文标题 | 黑体 | Arial | 三号 (16pt) |
| 公文正文 | 仿宋 | Times New Roman | 三号 (16pt) |
| 公文标题 | 方正小标宋 | — | 二号 (22pt) |
| 合同正文 | 宋体 | Times New Roman | 四号 (14pt) |
