/**
 * layer3-page-section.js
 * 第三层：页面/页眉页脚/页码/分节
 * 操作粒度为 Section（节）。
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) return;

  var PAPER_SIZE_MAP = registry.PAPER_SIZE_MAP;
  var ALIGNMENT_MAP = registry.ALIGNMENT_MAP;

  // ── 辅助函数 ──

  function getPageSetup(doc, sectionIndex) {
    if (sectionIndex && sectionIndex >= 1 && sectionIndex <= doc.Sections.Count) {
      return doc.Sections.Item(sectionIndex).PageSetup;
    }
    return doc.PageSetup;
  }

  function getSection(doc, sectionIndex) {
    var idx = sectionIndex || 1;
    if (idx >= 1 && idx <= doc.Sections.Count) {
      return doc.Sections.Item(idx);
    }
    return doc.Sections.Item(1);
  }

  // ── 页面设置 ──

  registry.skills['page.margins.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var ps = getPageSetup(doc, params.section);
    if (params.top !== undefined) ps.TopMargin = params.top;
    if (params.bottom !== undefined) ps.BottomMargin = params.bottom;
    if (params.left !== undefined) ps.LeftMargin = params.left;
    if (params.right !== undefined) ps.RightMargin = params.right;
    if (params.gutter !== undefined) ps.Gutter = params.gutter;
  };

  registry.skills['page.size.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var ps = getPageSetup(doc, params.section);

    if (params.paperSize && PAPER_SIZE_MAP[params.paperSize]) {
      var size = PAPER_SIZE_MAP[params.paperSize];
      ps.PageWidth = size.width;
      ps.PageHeight = size.height;
    } else if (params.width && params.height) {
      ps.PageWidth = params.width;
      ps.PageHeight = params.height;
    }
  };

  registry.skills['page.orientation.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var ps = getPageSetup(doc, params.section);
    // wdOrientPortrait = 0, wdOrientLandscape = 1
    ps.Orientation = params.orientation === 'landscape' ? 1 : 0;
  };

  // ── 页眉页脚 ──

  registry.skills['header.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);
    // wdHeaderFooterPrimary = 1
    var header = section.Headers.Item(1);
    var range = header.Range;
    range.Text = params.text || '';

    if (params.fontName) {
      range.Font.Name = params.fontName;
      range.Font.NameFarEast = params.fontName;
      range.Font.NameAscii = params.fontName;
    }
    if (params.fontSize) {
      range.Font.Size = params.fontSize;
    }
    if (params.alignment) {
      var align = ALIGNMENT_MAP[params.alignment];
      if (align !== undefined) {
        range.ParagraphFormat.Alignment = align;
      }
    }
  };

  registry.skills['header.clear'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);
    var header = section.Headers.Item(1);
    header.Range.Text = '';
  };

  registry.skills['footer.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);
    var footer = section.Footers.Item(1);
    var range = footer.Range;
    range.Text = params.text || '';

    if (params.fontName) {
      range.Font.Name = params.fontName;
      range.Font.NameFarEast = params.fontName;
      range.Font.NameAscii = params.fontName;
    }
    if (params.fontSize) {
      range.Font.Size = params.fontSize;
    }
    if (params.alignment) {
      var align = ALIGNMENT_MAP[params.alignment];
      if (align !== undefined) {
        range.ParagraphFormat.Alignment = align;
      }
    }
  };

  registry.skills['footer.clear'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);
    var footer = section.Footers.Item(1);
    footer.Range.Text = '';
  };

  // ── 页码 ──

  registry.skills['page_number.set'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);

    // 页码样式
    var numberStyleMap = {
      'arabic': 0,       // wdPageNumberStyleArabic
      'roman_lower': 1,  // wdPageNumberStyleLowercaseRoman
      'roman_upper': 2   // wdPageNumberStyleUppercaseRoman
    };

    // 页码位置
    var isTop = (params.position || '').startsWith('top_');
    var headerFooter = isTop ? section.Headers.Item(1) : section.Footers.Item(1);

    // 设置页码样式
    var style = numberStyleMap[params.format] || 0;
    headerFooter.PageNumbers.NumberStyle = style;

    // 添加页码
    if (headerFooter.PageNumbers.Count === 0) {
      // wdAlignPageNumberLeft=0, Center=1, Right=2
      var alignMap = {
        'bottom_left': 0, 'bottom_center': 1, 'bottom_right': 2,
        'top_left': 0, 'top_center': 1, 'top_right': 2
      };
      var align = alignMap[params.position] || 1;
      headerFooter.PageNumbers.Add(align, true);
    }
  };

  registry.skills['page_number.restart'].execute = function (params) {
    var doc = wps.WpsApplication().ActiveDocument;
    var section = getSection(doc, params.section);

    // 断开与前一节的链接
    section.Headers.Item(1).LinkToPrevious = false;
    section.Footers.Item(1).LinkToPrevious = false;

    var footer = section.Footers.Item(1);
    footer.PageNumbers.RestartNumberingAtSection = true;
    footer.PageNumbers.StartingNumber = params.startFrom || 1;

    if (params.format) {
      var numberStyleMap = {
        'arabic': 0,
        'roman_lower': 1,
        'roman_upper': 2
      };
      footer.PageNumbers.NumberStyle = numberStyleMap[params.format] || 0;
    }
  };

  // ── 分节/分页 ──

  registry.skills['section.break.insert'].execute = function (params, rangeOrIndices) {
    var breakTypeMap = {
      'nextPage': 2,     // wdSectionBreakNextPage
      'continuous': 3,   // wdSectionBreakContinuous
      'evenPage': 4,     // wdSectionBreakEvenPage
      'oddPage': 5       // wdSectionBreakOddPage
    };
    var breakType = breakTypeMap[params.breakType] || 2;

    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;

    if (Array.isArray(rangeOrIndices)) {
      // 从后往前插入，避免索引偏移
      var sorted = rangeOrIndices.slice().sort(function (a, b) { return b - a; });
      for (var i = 0; i < sorted.length; i++) {
        var para = doc.Paragraphs.Item(sorted[i]);
        para.Range.InsertBreak(breakType);
      }
    } else if (rangeOrIndices) {
      rangeOrIndices.InsertBreak(breakType);
    }
  };

  registry.skills['section.page_break_before'].execute = function (params, rangeOrIndices) {
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;

    if (Array.isArray(rangeOrIndices)) {
      for (var i = 0; i < rangeOrIndices.length; i++) {
        var idx = rangeOrIndices[i];
        if (idx >= 1 && idx <= doc.Paragraphs.Count) {
          doc.Paragraphs.Item(idx).Format.PageBreakBefore = -1; // True
        }
      }
    } else if (rangeOrIndices) {
      var paraCount = rangeOrIndices.Paragraphs.Count;
      for (var j = 1; j <= paraCount; j++) {
        rangeOrIndices.Paragraphs.Item(j).Format.PageBreakBefore = -1;
      }
    }
  };

  registry.skills['section.keep_with_next'].execute = function (params, rangeOrIndices) {
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;

    if (Array.isArray(rangeOrIndices)) {
      for (var i = 0; i < rangeOrIndices.length; i++) {
        var idx = rangeOrIndices[i];
        if (idx >= 1 && idx <= doc.Paragraphs.Count) {
          doc.Paragraphs.Item(idx).Format.KeepWithNext = -1; // True
        }
      }
    } else if (rangeOrIndices) {
      var paraCount = rangeOrIndices.Paragraphs.Count;
      for (var j = 1; j <= paraCount; j++) {
        rangeOrIndices.Paragraphs.Item(j).Format.KeepWithNext = -1;
      }
    }
  };
})();
