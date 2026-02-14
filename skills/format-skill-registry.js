/**
 * format-skill-registry.js
 * 全技能注册表 + 参数 Schema
 * 每个技能包含: id, layer, params 定义, execute 函数引用
 */

/* global window */

(function () {
  'use strict';

  // 中文字号→磅值映射
  const CHINESE_FONT_SIZE_MAP = {
    '初号': 42, '小初': 36,
    '一号': 26, '小一': 24,
    '二号': 22, '小二': 18,
    '三号': 16, '小三': 15,
    '四号': 14, '小四': 12,
    '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5,
    '七号': 5.5, '八号': 5
  };

  // 字体名映射（中文→API名）
  const FONT_NAME_MAP = {
    '宋体': 'SimSun', '黑体': 'SimHei', '楷体': 'KaiTi', '仿宋': 'FangSong',
    '微软雅黑': 'Microsoft YaHei', '华文中宋': 'STZhongsong', '华文仿宋': 'STFangsong',
    '华文楷体': 'STKaiti', '华文宋体': 'STSong', '华文细黑': 'STXihei',
    '等线': 'DengXian', '新宋体': 'NSimSun', '方正小标宋': 'FZXiaoBiaoSong-B05',
    '方正仿宋': 'FZFangSong-Z02'
  };

  // 纸张大小映射
  const PAPER_SIZE_MAP = {
    'A3': { width: 841.9, height: 1190.5 },
    'A4': { width: 595.3, height: 841.9 },
    'A5': { width: 419.5, height: 595.3 },
    'B5': { width: 498.9, height: 708.7 },
    'Letter': { width: 612, height: 792 },
    'Legal': { width: 612, height: 1008 },
    '16K': { width: 546.7, height: 787.4 }
  };

  // 对齐方式映射
  const ALIGNMENT_MAP = {
    'left': 0,      // wdAlignParagraphLeft
    'center': 1,    // wdAlignParagraphCenter
    'right': 2,     // wdAlignParagraphRight
    'justify': 3    // wdAlignParagraphJustify
  };

  // 行距模式映射
  const LINE_SPACING_RULE_MAP = {
    'multiple': 5,   // wdLineSpaceMultiple
    'exact': 4,      // wdLineSpaceExactly
    'atLeast': 3     // wdLineSpaceAtLeast
  };

  // role → detect 函数映射
  const ROLE_DETECT_MAP = {
    'title': 'detect.title',
    'heading_1': 'detect.headings',
    'heading_2': 'detect.headings',
    'heading_3': 'detect.headings',
    'heading_4': 'detect.headings',
    'abstract': 'detect.abstract',
    'keywords': 'detect.keywords',
    'body': 'detect.body_paragraphs',
    'figure_caption': 'detect.figure_captions',
    'table_caption': 'detect.table_captions',
    'references': 'detect.references',
    'acknowledgment': 'detect.acknowledgment',
    'toc': 'detect.toc_area',
    'footnotes': 'detect.footnotes'
  };

  // 技能定义注册表
  const skills = {};

  function defineSkill(id, definition) {
    skills[id] = { id, ...definition };
  }

  // ── 第一层：结构识别 ──

  defineSkill('detect.title', {
    layer: 1, description: '识别主标题',
    params: {},
    execute: null // 由 layer1-structure-detect.js 注入
  });

  defineSkill('detect.headings', {
    layer: 1, description: '识别多级标题',
    params: { levels: { type: 'array', required: false } },
    execute: null
  });

  defineSkill('detect.abstract', {
    layer: 1, description: '识别摘要段落',
    params: {},
    execute: null
  });

  defineSkill('detect.keywords', {
    layer: 1, description: '识别关键词段落',
    params: {},
    execute: null
  });

  defineSkill('detect.toc_area', {
    layer: 1, description: '识别目录区域',
    params: {},
    execute: null
  });

  defineSkill('detect.body_paragraphs', {
    layer: 1, description: '识别正文段落',
    params: {},
    execute: null
  });

  defineSkill('detect.figure_captions', {
    layer: 1, description: '识别图标题',
    params: {},
    execute: null
  });

  defineSkill('detect.table_captions', {
    layer: 1, description: '识别表标题',
    params: {},
    execute: null
  });

  defineSkill('detect.images', {
    layer: 1, description: '识别图片段落',
    params: {},
    execute: null
  });

  defineSkill('detect.tables', {
    layer: 1, description: '识别表格区域',
    params: {},
    execute: null
  });

  defineSkill('detect.references', {
    layer: 1, description: '识别参考文献',
    params: {},
    execute: null
  });

  defineSkill('detect.acknowledgment', {
    layer: 1, description: '识别致谢/附录',
    params: {},
    execute: null
  });

  defineSkill('detect.footnotes', {
    layer: 1, description: '识别脚注',
    params: {},
    execute: null
  });

  // ── 第二层：字体操作 ──

  defineSkill('text.font.set', {
    layer: 2, description: '设置字体属性',
    params: {
      target: { type: 'object', required: false },
      zhFont: { type: 'string', required: false },
      enFont: { type: 'string', required: false },
      fontSize: { type: 'number', required: false },
      bold: { type: 'boolean', required: false },
      italic: { type: 'boolean', required: false },
      underline: { type: 'boolean', required: false },
      color: { type: 'string', required: false },
      strikethrough: { type: 'boolean', required: false }
    },
    execute: null
  });

  defineSkill('text.font.set_zh', {
    layer: 2, description: '设置中文字体',
    params: {
      target: { type: 'object', required: false },
      fontName: { type: 'string', required: true }
    },
    execute: null
  });

  defineSkill('text.font.set_en', {
    layer: 2, description: '设置英文字体',
    params: {
      target: { type: 'object', required: false },
      fontName: { type: 'string', required: true }
    },
    execute: null
  });

  defineSkill('text.font.set_size', {
    layer: 2, description: '设置字号',
    params: {
      target: { type: 'object', required: false },
      size: { type: 'number', required: true }
    },
    execute: null
  });

  defineSkill('text.font.set_bold', {
    layer: 2, description: '设置加粗',
    params: {
      target: { type: 'object', required: false },
      bold: { type: 'boolean', required: true }
    },
    execute: null
  });

  defineSkill('text.font.set_color', {
    layer: 2, description: '设置颜色',
    params: {
      target: { type: 'object', required: false },
      color: { type: 'string', required: true }
    },
    execute: null
  });

  defineSkill('text.clear_formatting', {
    layer: 2, description: '清除格式',
    params: { target: { type: 'object', required: false } },
    execute: null
  });

  defineSkill('text.superscript', {
    layer: 2, description: '上标',
    params: {
      target: { type: 'object', required: false },
      enabled: { type: 'boolean', required: true }
    },
    execute: null
  });

  defineSkill('text.subscript', {
    layer: 2, description: '下标',
    params: {
      target: { type: 'object', required: false },
      enabled: { type: 'boolean', required: true }
    },
    execute: null
  });

  defineSkill('text.highlight', {
    layer: 2, description: '文字高亮',
    params: {
      target: { type: 'object', required: false },
      color: { type: 'string', required: true }
    },
    execute: null
  });

  // ── 第二层：段落操作 ──

  defineSkill('paragraph.alignment.set', {
    layer: 2, description: '设置对齐方式',
    params: {
      target: { type: 'object', required: false },
      alignment: { type: 'enum', required: true, values: ['left', 'center', 'right', 'justify'] }
    },
    execute: null
  });

  defineSkill('paragraph.spacing.set', {
    layer: 2, description: '设置段前段后距',
    params: {
      target: { type: 'object', required: false },
      spaceBefore: { type: 'number', required: false },
      spaceAfter: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('paragraph.line_spacing.set', {
    layer: 2, description: '设置行距',
    params: {
      target: { type: 'object', required: false },
      mode: { type: 'enum', required: true, values: ['multiple', 'exact', 'atLeast'] },
      value: { type: 'number', required: true }
    },
    execute: null
  });

  defineSkill('paragraph.indent.set', {
    layer: 2, description: '设置缩进',
    params: {
      target: { type: 'object', required: false },
      firstLineChars: { type: 'number', required: false },
      firstLinePoints: { type: 'number', required: false },
      hanging: { type: 'number', required: false },
      left: { type: 'number', required: false },
      right: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('paragraph.columns.set', {
    layer: 2, description: '设置分栏',
    params: {
      target: { type: 'object', required: false },
      count: { type: 'number', required: true },
      spacing: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('paragraph.borders.set', {
    layer: 2, description: '设置段落边框',
    params: {
      target: { type: 'object', required: false },
      borderType: { type: 'string', required: true },
      lineWidth: { type: 'number', required: false },
      color: { type: 'string', required: false }
    },
    execute: null
  });

  defineSkill('paragraph.numbering.set', {
    layer: 2, description: '设置编号列表',
    params: {
      target: { type: 'object', required: false },
      format: { type: 'string', required: true },
      startAt: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('paragraph.bullets.set', {
    layer: 2, description: '设置项目符号',
    params: {
      target: { type: 'object', required: false },
      symbol: { type: 'string', required: false }
    },
    execute: null
  });

  // ── 第三层：页面设置 ──

  defineSkill('page.margins.set', {
    layer: 3, description: '设置页边距',
    params: {
      section: { type: 'number', required: false },
      top: { type: 'number', required: true },
      bottom: { type: 'number', required: true },
      left: { type: 'number', required: true },
      right: { type: 'number', required: true },
      gutter: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('page.size.set', {
    layer: 3, description: '设置纸张大小',
    params: {
      section: { type: 'number', required: false },
      paperSize: { type: 'enum', required: true, values: ['A3', 'A4', 'A5', 'B5', 'Letter', 'Legal', '16K'] },
      width: { type: 'number', required: false },
      height: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('page.orientation.set', {
    layer: 3, description: '设置纸张方向',
    params: {
      section: { type: 'number', required: false },
      orientation: { type: 'enum', required: true, values: ['portrait', 'landscape'] }
    },
    execute: null
  });

  // ── 第三层：页眉页脚 ──

  defineSkill('header.set', {
    layer: 3, description: '设置页眉',
    params: {
      section: { type: 'number', required: false },
      text: { type: 'string', required: true },
      alignment: { type: 'enum', required: false, values: ['left', 'center', 'right'] },
      fontName: { type: 'string', required: false },
      fontSize: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('header.clear', {
    layer: 3, description: '清除页眉',
    params: { section: { type: 'number', required: false } },
    execute: null
  });

  defineSkill('footer.set', {
    layer: 3, description: '设置页脚',
    params: {
      section: { type: 'number', required: false },
      text: { type: 'string', required: true },
      alignment: { type: 'enum', required: false, values: ['left', 'center', 'right'] },
      fontName: { type: 'string', required: false },
      fontSize: { type: 'number', required: false }
    },
    execute: null
  });

  defineSkill('footer.clear', {
    layer: 3, description: '清除页脚',
    params: { section: { type: 'number', required: false } },
    execute: null
  });

  // ── 第三层：页码 ──

  defineSkill('page_number.set', {
    layer: 3, description: '设置页码',
    params: {
      section: { type: 'number', required: false },
      format: { type: 'enum', required: true, values: ['arabic', 'roman_lower', 'roman_upper'] },
      position: { type: 'enum', required: true, values: ['bottom_left', 'bottom_center', 'bottom_right', 'top_left', 'top_center', 'top_right'] }
    },
    execute: null
  });

  defineSkill('page_number.restart', {
    layer: 3, description: '分节重新编号',
    params: {
      section: { type: 'number', required: true },
      startFrom: { type: 'number', required: true },
      format: { type: 'enum', required: false, values: ['arabic', 'roman_lower', 'roman_upper'] }
    },
    execute: null
  });

  // ── 第三层：分节/分页 ──

  defineSkill('section.break.insert', {
    layer: 3, description: '插入分节符',
    params: {
      target: { type: 'object', required: true },
      breakType: { type: 'enum', required: true, values: ['nextPage', 'continuous', 'evenPage', 'oddPage'] }
    },
    execute: null
  });

  defineSkill('section.page_break_before', {
    layer: 3, description: '段前分页',
    params: { target: { type: 'object', required: true } },
    execute: null
  });

  defineSkill('section.keep_with_next', {
    layer: 3, description: '与下段同页',
    params: { target: { type: 'object', required: true } },
    execute: null
  });

  // ── 第六层：自定义代码 ──

  defineSkill('custom.execute', {
    layer: 6, description: '执行自定义 WPS JSAPI 代码',
    params: {
      code: { type: 'string', required: true },
      description: { type: 'string', required: true },
      target: { type: 'object', required: false }
    },
    execute: null
  });

  // ── 导出 ──

  window.FormatSkillRegistry = {
    skills: skills,
    defineSkill: defineSkill,
    CHINESE_FONT_SIZE_MAP: CHINESE_FONT_SIZE_MAP,
    FONT_NAME_MAP: FONT_NAME_MAP,
    PAPER_SIZE_MAP: PAPER_SIZE_MAP,
    ALIGNMENT_MAP: ALIGNMENT_MAP,
    LINE_SPACING_RULE_MAP: LINE_SPACING_RULE_MAP,
    ROLE_DETECT_MAP: ROLE_DETECT_MAP
  };
})();
