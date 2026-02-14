/**
 * layer5-style-toc.js
 * 第五层（补充）：样式管理与目录操作
 * 提供样式应用/创建/修改、目录生成/更新、水印等技能。
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) return;

  var ALIGNMENT_MAP = registry.ALIGNMENT_MAP;
  var LINE_SPACING_RULE_MAP = registry.LINE_SPACING_RULE_MAP;

  // ── 样式技能 ──

  // 应用已有样式到目标
  registry.skills['style.apply'] = {
    id: 'style.apply',
    layer: 5,
    description: '应用已有样式',
    params: {
      target: { type: 'object', required: true },
      styleName: { type: 'string', required: true }
    },
    execute: function (params, rangeOrIndices) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      function applyStyle(range) {
        range.Style = params.styleName;
      }

      if (Array.isArray(rangeOrIndices)) {
        for (var i = 0; i < rangeOrIndices.length; i++) {
          var idx = rangeOrIndices[i];
          if (idx >= 1 && idx <= doc.Paragraphs.Count) {
            applyStyle(doc.Paragraphs.Item(idx).Range);
          }
        }
      } else if (rangeOrIndices) {
        applyStyle(rangeOrIndices);
      }
    }
  };

  // 修改已有样式的属性
  registry.skills['style.modify'] = {
    id: 'style.modify',
    layer: 5,
    description: '修改已有样式属性',
    params: {
      styleName: { type: 'string', required: true },
      zhFont: { type: 'string', required: false },
      enFont: { type: 'string', required: false },
      fontSize: { type: 'number', required: false },
      bold: { type: 'boolean', required: false },
      italic: { type: 'boolean', required: false },
      alignment: { type: 'enum', required: false, values: ['left', 'center', 'right', 'justify'] },
      lineSpacingMode: { type: 'enum', required: false, values: ['multiple', 'exact', 'atLeast'] },
      lineSpacingValue: { type: 'number', required: false },
      spaceBefore: { type: 'number', required: false },
      spaceAfter: { type: 'number', required: false },
      firstLineIndent: { type: 'number', required: false }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var style;
      try {
        style = doc.Styles.Item(params.styleName);
      } catch (e) {
        throw new Error('样式 "' + params.styleName + '" 不存在');
      }

      var font = style.Font;
      if (params.zhFont) {
        font.NameFarEast = params.zhFont;
        if (!params.enFont) font.Name = params.zhFont;
      }
      if (params.enFont) {
        font.NameAscii = params.enFont;
        font.NameOther = params.enFont;
      }
      if (params.fontSize !== undefined) font.Size = params.fontSize;
      if (params.bold !== undefined) font.Bold = params.bold ? -1 : 0;
      if (params.italic !== undefined) font.Italic = params.italic ? -1 : 0;

      var pf = style.ParagraphFormat;
      if (params.alignment) {
        var align = ALIGNMENT_MAP[params.alignment];
        if (align !== undefined) pf.Alignment = align;
      }
      if (params.lineSpacingMode && params.lineSpacingValue !== undefined) {
        var rule = LINE_SPACING_RULE_MAP[params.lineSpacingMode];
        if (rule !== undefined) {
          pf.LineSpacingRule = rule;
          if (params.lineSpacingMode === 'multiple') {
            pf.LineSpacing = params.lineSpacingValue * 12;
          } else {
            pf.LineSpacing = params.lineSpacingValue;
          }
        }
      }
      if (params.spaceBefore !== undefined) pf.SpaceBefore = params.spaceBefore;
      if (params.spaceAfter !== undefined) pf.SpaceAfter = params.spaceAfter;
      if (params.firstLineIndent !== undefined) pf.FirstLineIndent = params.firstLineIndent;
    }
  };

  // 创建新样式
  registry.skills['style.create'] = {
    id: 'style.create',
    layer: 5,
    description: '创建新样式',
    params: {
      styleName: { type: 'string', required: true },
      basedOn: { type: 'string', required: false },
      zhFont: { type: 'string', required: false },
      enFont: { type: 'string', required: false },
      fontSize: { type: 'number', required: false },
      bold: { type: 'boolean', required: false },
      alignment: { type: 'enum', required: false, values: ['left', 'center', 'right', 'justify'] }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;

      // 检查是否已存在
      try {
        doc.Styles.Item(params.styleName);
        // 已存在，改为修改
        registry.skills['style.modify'].execute(params);
        return;
      } catch (e) {
        // 不存在，创建新样式
      }

      // wdStyleTypeParagraph = 1
      var newStyle = doc.Styles.Add(params.styleName, 1);

      if (params.basedOn) {
        try { newStyle.BaseStyle = params.basedOn; } catch (e) { /* ignore */ }
      }

      var font = newStyle.Font;
      if (params.zhFont) {
        font.NameFarEast = params.zhFont;
        if (!params.enFont) font.Name = params.zhFont;
      }
      if (params.enFont) {
        font.NameAscii = params.enFont;
        font.NameOther = params.enFont;
      }
      if (params.fontSize !== undefined) font.Size = params.fontSize;
      if (params.bold !== undefined) font.Bold = params.bold ? -1 : 0;

      if (params.alignment) {
        var align = ALIGNMENT_MAP[params.alignment];
        if (align !== undefined) newStyle.ParagraphFormat.Alignment = align;
      }
    }
  };

  // ── 目录技能 ──

  // 插入/更新目录
  registry.skills['toc.insert'] = {
    id: 'toc.insert',
    layer: 5,
    description: '插入目录',
    params: {
      levels: { type: 'number', required: false },
      target: { type: 'object', required: false }
    },
    execute: function (params, rangeOrIndices) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      var levels = params.levels || 3;

      // 检查是否已有目录 → 直接更新
      if (doc.TablesOfContents.Count > 0) {
        doc.TablesOfContents.Item(1).Update();
        return;
      }

      // 确定插入点：必须用折叠的 Range（start===end），否则会替换内容
      var insertPos;
      if (Array.isArray(rangeOrIndices) && rangeOrIndices.length > 0) {
        insertPos = doc.Paragraphs.Item(rangeOrIndices[0]).Range.Start;
      } else {
        insertPos = 0;
      }
      var insertRange = doc.Range(insertPos, insertPos);

      // TablesOfContents.Add(Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel)
      doc.TablesOfContents.Add(insertRange, true, 1, levels);
    }
  };

  // 更新目录
  registry.skills['toc.update'] = {
    id: 'toc.update',
    layer: 5,
    description: '更新目录',
    params: {},
    execute: function () {
      var doc = wps.WpsApplication().ActiveDocument;
      if (doc.TablesOfContents.Count > 0) {
        for (var i = 1; i <= doc.TablesOfContents.Count; i++) {
          doc.TablesOfContents.Item(i).Update();
        }
      }
    }
  };

  // 删除目录
  registry.skills['toc.remove'] = {
    id: 'toc.remove',
    layer: 5,
    description: '删除目录',
    params: {},
    execute: function () {
      var doc = wps.WpsApplication().ActiveDocument;
      if (doc.TablesOfContents.Count > 0) {
        for (var i = doc.TablesOfContents.Count; i >= 1; i--) {
          doc.TablesOfContents.Item(i).Delete();
        }
      }
    }
  };

  // ── 水印技能 ──

  // 文字水印
  registry.skills['watermark.text.set'] = {
    id: 'watermark.text.set',
    layer: 5,
    description: '设置文字水印',
    params: {
      text: { type: 'string', required: true },
      fontName: { type: 'string', required: false },
      fontSize: { type: 'number', required: false },
      color: { type: 'string', required: false },
      layout: { type: 'enum', required: false, values: ['diagonal', 'horizontal'] }
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      var section = doc.Sections.Item(1);
      var header = section.Headers.Item(1);
      var headerRange = header.Range;

      // 在页眉中添加形状作为水印
      var shapes = header.Shapes;
      var text = params.text || '水印';
      var fontSize = params.fontSize || 54;
      var fontName = params.fontName || 'SimSun';

      // 添加文本框作为水印
      // AddTextEffect(PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top)
      try {
        var shape = shapes.AddTextEffect(
          0, // msoTextEffect1
          text,
          fontName,
          fontSize,
          0, // not bold
          0, // not italic
          0, // left
          0  // top
        );

        // 设为水印样式：在文字下方、半透明
        shape.WrapFormat.Type = 3; // wdWrapBehind (实际可能需要调整)
        shape.RelativeHorizontalPosition = 1; // wdRelativeHorizontalPositionPage
        shape.RelativeVerticalPosition = 1; // wdRelativeVerticalPositionPage

        if (params.layout === 'diagonal') {
          shape.Rotation = -45;
        }
      } catch (e) {
        // 如果 AddTextEffect 不可用，使用 AddTextbox 替代
        var textbox = shapes.AddTextbox(
          1,   // msoTextOrientationHorizontal
          100, // left
          200, // top
          400, // width
          100  // height
        );
        textbox.TextFrame.TextRange.Text = text;
        textbox.TextFrame.TextRange.Font.Size = fontSize;
        textbox.TextFrame.TextRange.Font.Name = fontName;
        if (params.color) {
          textbox.TextFrame.TextRange.Font.Color = hexToWpsColor(params.color);
        }
        textbox.Line.Visible = 0; // 无边框
        textbox.Fill.Visible = 0; // 无填充
        if (params.layout === 'diagonal') {
          textbox.Rotation = -45;
        }
      }
    }
  };

  // 清除水印
  registry.skills['watermark.clear'] = {
    id: 'watermark.clear',
    layer: 5,
    description: '清除水印',
    params: {},
    execute: function () {
      var doc = wps.WpsApplication().ActiveDocument;
      // 水印通常在页眉中的形状
      for (var s = 1; s <= doc.Sections.Count; s++) {
        var header = doc.Sections.Item(s).Headers.Item(1);
        var shapes = header.Shapes;
        for (var i = shapes.Count; i >= 1; i--) {
          try { shapes.Item(i).Delete(); } catch (e) { /* ignore */ }
        }
      }
    }
  };

  // 辅助：hexToWpsColor
  function hexToWpsColor(hex) {
    if (typeof hex === 'number') return hex;
    if (typeof hex !== 'string') return 0;
    hex = hex.replace(/^#/, '');
    if (hex.length === 3) {
      hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    var r = parseInt(hex.substring(0, 2), 16) || 0;
    var g = parseInt(hex.substring(2, 4), 16) || 0;
    var b = parseInt(hex.substring(4, 6), 16) || 0;
    return (b << 16) | (g << 8) | r;
  }
})();
