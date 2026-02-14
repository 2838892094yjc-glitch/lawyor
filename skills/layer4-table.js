/**
 * layer4-table.js
 * 第四层（补充）：表格样式操作
 * 提供表格边框、单元格对齐、底纹、行高、列宽、合并、三线表等技能。
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) return;

  var ALIGNMENT_MAP = registry.ALIGNMENT_MAP;

  // ── 辅助函数 ──

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

  function getTable(doc, tableIndex) {
    var idx = tableIndex || 1;
    if (idx >= 1 && idx <= doc.Tables.Count) {
      return doc.Tables.Item(idx);
    }
    return null;
  }

  // WPS 边框类型常量
  var WD_BORDER = {
    TOP: -1,        // wdBorderTop
    LEFT: -2,       // wdBorderLeft
    BOTTOM: -3,     // wdBorderBottom
    RIGHT: -4,      // wdBorderRight
    HORIZONTAL: -5, // wdBorderHorizontal
    VERTICAL: -6    // wdBorderVertical
  };

  // WPS 线型常量
  var WD_LINE_STYLE = {
    NONE: 0,         // wdLineStyleNone
    SINGLE: 1,       // wdLineStyleSingle
    DOUBLE: 7,       // wdLineStyleDouble
    DASHED: 3,       // wdLineStyleDashSmallGap（WPS 兼容值，失败时回退实线）
    THICK_THIN: 14,  // wdLineStyleThickThinSmallGap
    THIN_THICK: 15   // wdLineStyleThinThickSmallGap
  };

  // WPS 线宽常量
  var WD_LINE_WIDTH = {
    '0.25pt': 2,  // wdLineWidth025pt
    '0.5pt': 4,   // wdLineWidth050pt
    '0.75pt': 6,  // wdLineWidth075pt
    '1pt': 8,     // wdLineWidth100pt
    '1.5pt': 12,  // wdLineWidth150pt
    '2.25pt': 18, // wdLineWidth225pt
    '3pt': 24,    // wdLineWidth300pt
    '4.5pt': 36,  // wdLineWidth450pt
    '6pt': 48     // wdLineWidth600pt
  };

  function parseLineWidth(w) {
    if (typeof w === 'number') return w;
    if (typeof w === 'string' && WD_LINE_WIDTH[w]) return WD_LINE_WIDTH[w];
    return 8; // default 1pt
  }

  function setBorderLine(border, lineStyle, lineWidth, color) {
    function tryDashed(borderObj) {
      var candidates = [WD_LINE_STYLE.DASHED, 4, 2];
      for (var i = 0; i < candidates.length; i++) {
        var c = candidates[i];
        try {
          borderObj.LineStyle = c;
          if (borderObj.LineStyle === c) return true;
        } catch (e) { /* try next */ }
      }
      return false;
    }

    if (lineStyle === WD_LINE_STYLE.DASHED) {
      var dashedOk = tryDashed(border);
      if (!dashedOk) {
        try { border.LineStyle = WD_LINE_STYLE.SINGLE; } catch (e2) { /* ignore */ }
      }
    } else {
      try {
        border.LineStyle = lineStyle;
      } catch (e3) {
        if (lineStyle !== WD_LINE_STYLE.NONE) {
          try { border.LineStyle = WD_LINE_STYLE.SINGLE; } catch (e4) { /* ignore */ }
        }
      }
    }

    if (lineStyle !== WD_LINE_STYLE.NONE) {
      if (lineWidth !== undefined && lineWidth !== null) {
        try { border.LineWidth = lineWidth; } catch (e5) { /* ignore */ }
      }
      if (color) {
        try { border.Color = color; } catch (e6) { /* ignore */ }
      }
    }
  }

  // ── 技能注册 ──

  // 表格边框样式
  registry.skills['table.borders.set'] = {
    id: 'table.borders.set',
    layer: 4,
    description: '设置表格边框',
    params: {
      tableIndex: { type: 'number', required: true },
      style: { type: 'enum', required: false, values: ['all', 'box', 'none', 'three_line', 'dashed'] },
      lineWidth: { type: 'string', required: false },
      color: { type: 'string', required: false }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      var width = parseLineWidth(params.lineWidth);
      var color = params.color ? hexToWpsColor(params.color) : 0;
      var style = params.style || 'all';

      if (style === 'none') {
        // 清除所有边框
        var borders = [WD_BORDER.TOP, WD_BORDER.BOTTOM, WD_BORDER.LEFT, WD_BORDER.RIGHT, WD_BORDER.HORIZONTAL, WD_BORDER.VERTICAL];
        for (var b = 0; b < borders.length; b++) {
          table.Borders.Item(borders[b]).LineStyle = WD_LINE_STYLE.NONE;
        }
      } else if (style === 'dashed') {
        // 虚线边框：优先尝试 dashed，WPS 不支持时回退实线
        var dashedBorders = [WD_BORDER.TOP, WD_BORDER.BOTTOM, WD_BORDER.LEFT, WD_BORDER.RIGHT, WD_BORDER.HORIZONTAL, WD_BORDER.VERTICAL];
        for (var db = 0; db < dashedBorders.length; db++) {
          setBorderLine(table.Borders.Item(dashedBorders[db]), WD_LINE_STYLE.DASHED, width, color);
        }
      } else if (style === 'three_line') {
        // 三线表：顶部粗线、底部粗线、表头下细线、无竖线
        // 清除所有边框先
        var allBorders = [WD_BORDER.TOP, WD_BORDER.BOTTOM, WD_BORDER.LEFT, WD_BORDER.RIGHT, WD_BORDER.HORIZONTAL, WD_BORDER.VERTICAL];
        for (var c = 0; c < allBorders.length; c++) {
          table.Borders.Item(allBorders[c]).LineStyle = WD_LINE_STYLE.NONE;
        }
        // 顶部粗线 (1.5pt)
        table.Borders.Item(WD_BORDER.TOP).LineStyle = WD_LINE_STYLE.SINGLE;
        table.Borders.Item(WD_BORDER.TOP).LineWidth = WD_LINE_WIDTH['1.5pt'];
        if (color) table.Borders.Item(WD_BORDER.TOP).Color = color;
        // 底部粗线 (1.5pt)
        table.Borders.Item(WD_BORDER.BOTTOM).LineStyle = WD_LINE_STYLE.SINGLE;
        table.Borders.Item(WD_BORDER.BOTTOM).LineWidth = WD_LINE_WIDTH['1.5pt'];
        if (color) table.Borders.Item(WD_BORDER.BOTTOM).Color = color;
        // 表头下方细线 (0.75pt) — 第一行的下边框
        if (table.Rows.Count > 1) {
          var firstRow = table.Rows.Item(1);
          firstRow.Borders.Item(WD_BORDER.BOTTOM).LineStyle = WD_LINE_STYLE.SINGLE;
          firstRow.Borders.Item(WD_BORDER.BOTTOM).LineWidth = WD_LINE_WIDTH['0.75pt'];
          if (color) firstRow.Borders.Item(WD_BORDER.BOTTOM).Color = color;
        }
      } else if (style === 'box') {
        // 仅外边框
        var outerBorders = [WD_BORDER.TOP, WD_BORDER.BOTTOM, WD_BORDER.LEFT, WD_BORDER.RIGHT];
        // 清除内部线
        table.Borders.Item(WD_BORDER.HORIZONTAL).LineStyle = WD_LINE_STYLE.NONE;
        table.Borders.Item(WD_BORDER.VERTICAL).LineStyle = WD_LINE_STYLE.NONE;
        for (var d = 0; d < outerBorders.length; d++) {
          table.Borders.Item(outerBorders[d]).LineStyle = WD_LINE_STYLE.SINGLE;
          table.Borders.Item(outerBorders[d]).LineWidth = width;
          if (color) table.Borders.Item(outerBorders[d]).Color = color;
        }
      } else {
        // all: 全部边框
        var all = [WD_BORDER.TOP, WD_BORDER.BOTTOM, WD_BORDER.LEFT, WD_BORDER.RIGHT, WD_BORDER.HORIZONTAL, WD_BORDER.VERTICAL];
        for (var e = 0; e < all.length; e++) {
          table.Borders.Item(all[e]).LineStyle = WD_LINE_STYLE.SINGLE;
          table.Borders.Item(all[e]).LineWidth = width;
          if (color) table.Borders.Item(all[e]).Color = color;
        }
      }
    }
  };

  // 表格单元格对齐
  registry.skills['table.cell_alignment.set'] = {
    id: 'table.cell_alignment.set',
    layer: 4,
    description: '设置表格单元格对齐',
    params: {
      tableIndex: { type: 'number', required: true },
      row: { type: 'number', required: false },
      col: { type: 'number', required: false },
      horizontal: { type: 'enum', required: false, values: ['left', 'center', 'right', 'justify'] },
      vertical: { type: 'enum', required: false, values: ['top', 'center', 'bottom'] }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      var verticalMap = { 'top': 0, 'center': 1, 'bottom': 3 };

      function applyAlignment(cell) {
        if (params.horizontal) {
          var align = ALIGNMENT_MAP[params.horizontal];
          if (align !== undefined) cell.Range.ParagraphFormat.Alignment = align;
        }
        if (params.vertical) {
          var vAlign = verticalMap[params.vertical];
          if (vAlign !== undefined) cell.VerticalAlignment = vAlign;
        }
      }

      if (params.row && params.col) {
        // 单个单元格
        applyAlignment(table.Cell(params.row, params.col));
      } else if (params.row) {
        // 整行
        var row = table.Rows.Item(params.row);
        var colCount = table.Columns.Count;
        for (var c = 1; c <= colCount; c++) {
          try { applyAlignment(table.Cell(params.row, c)); } catch (e) { /* merged cell */ }
        }
      } else {
        // 整个表格
        var rowCount = table.Rows.Count;
        var cols = table.Columns.Count;
        for (var r = 1; r <= rowCount; r++) {
          for (var cc = 1; cc <= cols; cc++) {
            try { applyAlignment(table.Cell(r, cc)); } catch (e) { /* merged cell */ }
          }
        }
      }
    }
  };

  // 表格底纹/背景色
  registry.skills['table.shading.set'] = {
    id: 'table.shading.set',
    layer: 4,
    description: '设置表格底纹/背景色',
    params: {
      tableIndex: { type: 'number', required: true },
      row: { type: 'number', required: false },
      col: { type: 'number', required: false },
      color: { type: 'string', required: true }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      var wpsColor = hexToWpsColor(params.color);

      if (params.row && params.col) {
        table.Cell(params.row, params.col).Shading.BackgroundPatternColor = wpsColor;
      } else if (params.row) {
        var colCount = table.Columns.Count;
        for (var c = 1; c <= colCount; c++) {
          try { table.Cell(params.row, c).Shading.BackgroundPatternColor = wpsColor; } catch (e) { /* merged */ }
        }
      } else {
        // 整个表格
        var rowCount = table.Rows.Count;
        var cols = table.Columns.Count;
        for (var r = 1; r <= rowCount; r++) {
          for (var cc = 1; cc <= cols; cc++) {
            try { table.Cell(r, cc).Shading.BackgroundPatternColor = wpsColor; } catch (e) { /* merged */ }
          }
        }
      }
    }
  };

  // 表格行高
  registry.skills['table.row_height.set'] = {
    id: 'table.row_height.set',
    layer: 4,
    description: '设置表格行高',
    params: {
      tableIndex: { type: 'number', required: true },
      row: { type: 'number', required: false },
      height: { type: 'number', required: true },
      rule: { type: 'enum', required: false, values: ['auto', 'atLeast', 'exact'] }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      // wdRowHeightAuto=0, wdRowHeightAtLeast=1, wdRowHeightExactly=2
      var ruleMap = { 'auto': 0, 'atLeast': 1, 'exact': 2 };
      var rule = ruleMap[params.rule || 'atLeast'] || 1;

      if (params.row) {
        table.Rows.Item(params.row).Height = params.height;
        table.Rows.Item(params.row).HeightRule = rule;
      } else {
        for (var r = 1; r <= table.Rows.Count; r++) {
          table.Rows.Item(r).Height = params.height;
          table.Rows.Item(r).HeightRule = rule;
        }
      }
    }
  };

  // 表格列宽
  registry.skills['table.column_width.set'] = {
    id: 'table.column_width.set',
    layer: 4,
    description: '设置表格列宽',
    params: {
      tableIndex: { type: 'number', required: true },
      col: { type: 'number', required: false },
      width: { type: 'number', required: true }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      if (params.col) {
        table.Columns.Item(params.col).Width = params.width;
      } else {
        for (var c = 1; c <= table.Columns.Count; c++) {
          table.Columns.Item(c).Width = params.width;
        }
      }
    }
  };

  // 表格自动适应
  registry.skills['table.autofit.set'] = {
    id: 'table.autofit.set',
    layer: 4,
    description: '表格自动适应',
    params: {
      tableIndex: { type: 'number', required: true },
      mode: { type: 'enum', required: true, values: ['content', 'window', 'fixed'] }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      // wdAutoFitContent=1, wdAutoFitWindow=2, wdAutoFitFixed=0
      var modeMap = { 'content': 1, 'window': 2, 'fixed': 0 };
      table.AutoFitBehavior(modeMap[params.mode] || 2);
    }
  };

  // 表格字体设置
  registry.skills['table.font.set'] = {
    id: 'table.font.set',
    layer: 4,
    description: '设置表格字体',
    params: {
      tableIndex: { type: 'number', required: true },
      row: { type: 'number', required: false },
      zhFont: { type: 'string', required: false },
      enFont: { type: 'string', required: false },
      fontSize: { type: 'number', required: false },
      bold: { type: 'boolean', required: false }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      function applyFont(range) {
        if (params.zhFont) {
          range.Font.NameFarEast = params.zhFont;
          if (!params.enFont) range.Font.Name = params.zhFont;
        }
        if (params.enFont) {
          range.Font.NameAscii = params.enFont;
          range.Font.NameOther = params.enFont;
        }
        if (params.fontSize !== undefined) range.Font.Size = params.fontSize;
        if (params.bold !== undefined) range.Font.Bold = params.bold ? -1 : 0;
      }

      if (params.row) {
        var colCount = table.Columns.Count;
        for (var c = 1; c <= colCount; c++) {
          try { applyFont(table.Cell(params.row, c).Range); } catch (e) { /* merged */ }
        }
      } else {
        applyFont(table.Range);
      }
    }
  };

  // 表头重复
  registry.skills['table.header_row.set'] = {
    id: 'table.header_row.set',
    layer: 4,
    description: '设置表头行跨页重复',
    params: {
      tableIndex: { type: 'number', required: true },
      repeat: { type: 'boolean', required: true }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      table.Rows.Item(1).HeadingFormat = params.repeat ? -1 : 0;
    }
  };

  // 表格居中（表格在页面中的对齐）
  registry.skills['table.alignment.set'] = {
    id: 'table.alignment.set',
    layer: 4,
    description: '设置表格在页面中的对齐',
    params: {
      tableIndex: { type: 'number', required: true },
      alignment: { type: 'enum', required: true, values: ['left', 'center', 'right'] }
    },
    execute: function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var table = getTable(doc, params.tableIndex);
      if (!table) throw new Error('表格 ' + params.tableIndex + ' 不存在');

      // wdAlignRowLeft=0, Center=1, Right=2
      var alignMap = { 'left': 0, 'center': 1, 'right': 2 };
      for (var r = 1; r <= table.Rows.Count; r++) {
        table.Rows.Item(r).Alignment = alignMap[params.alignment] || 1;
      }
    }
  };
})();
