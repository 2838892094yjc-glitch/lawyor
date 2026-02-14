/**
 * layer2-text-paragraph.js
 * 第二层：文字与段落格式操作
 * 对已识别的标签类别或指定范围应用格式规则。
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) return;

  var ALIGNMENT_MAP = registry.ALIGNMENT_MAP;
  var LINE_SPACING_RULE_MAP = registry.LINE_SPACING_RULE_MAP;

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
    // WPS uses BGR format: (B << 16) | (G << 8) | R
    return (b << 16) | (g << 8) | r;
  }

  function setFontOnRange(range, params) {
    var font = range.Font;
    if (params.zhFont) {
      font.NameFarEast = params.zhFont;
      if (!params.enFont) {
        font.Name = params.zhFont;
      }
    }
    if (params.enFont) {
      font.NameAscii = params.enFont;
      font.NameOther = params.enFont;
    }
    if (params.fontSize !== undefined && params.fontSize !== null) {
      font.Size = params.fontSize;
    }
    if (params.bold !== undefined) {
      font.Bold = params.bold ? -1 : 0;
    }
    if (params.italic !== undefined) {
      font.Italic = params.italic ? -1 : 0;
    }
    if (params.underline !== undefined) {
      font.Underline = params.underline ? 1 : 0; // wdUnderlineSingle = 1
    }
    if (params.color !== undefined) {
      font.Color = hexToWpsColor(params.color);
    }
    if (params.strikethrough !== undefined) {
      font.StrikeThrough = params.strikethrough ? -1 : 0;
    }
  }

  /**
   * 将连续段落索引合并为区间，减少 WPS API 调用次数。
   * [3,4,5,6, 10,11,12] → [{from:3,to:6}, {from:10,to:12}]
   */
  function mergeContiguousIndices(indices) {
    if (!indices || indices.length === 0) return [];
    var sorted = indices.slice().sort(function (a, b) { return a - b; });
    var groups = [];
    var start = sorted[0], end = sorted[0];
    for (var i = 1; i < sorted.length; i++) {
      if (sorted[i] === end + 1) {
        end = sorted[i];
      } else {
        groups.push({ from: start, to: end });
        start = sorted[i];
        end = sorted[i];
      }
    }
    groups.push({ from: start, to: end });
    return groups;
  }

  /**
   * 通用执行器：处理 range 或 indices。
   * indices 模式下合并连续段落为区间，批量操作 Range。
   */
  function executeOnTarget(params, rangeOrIndices, fn) {
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;

    if (Array.isArray(rangeOrIndices)) {
      var groups = mergeContiguousIndices(rangeOrIndices);
      for (var g = 0; g < groups.length; g++) {
        var startPos = doc.Paragraphs.Item(groups[g].from).Range.Start;
        var endPos = doc.Paragraphs.Item(groups[g].to).Range.End;
        fn(doc.Range(startPos, endPos));
      }
    } else if (rangeOrIndices) {
      fn(rangeOrIndices);
    } else {
      fn(doc.Content);
    }
  }

  // ── 字体技能注册 ──

  registry.skills['text.font.set'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      setFontOnRange(range, params);
    });
  };

  registry.skills['text.font.set_zh'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.NameFarEast = params.fontName;
      range.Font.Name = params.fontName;
    });
  };

  registry.skills['text.font.set_en'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.NameAscii = params.fontName;
      range.Font.NameOther = params.fontName;
    });
  };

  registry.skills['text.font.set_size'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Size = params.size;
    });
  };

  registry.skills['text.font.set_bold'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Bold = params.bold ? -1 : 0;
    });
  };

  registry.skills['text.font.set_color'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Color = hexToWpsColor(params.color);
    });
  };

  registry.skills['text.clear_formatting'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Reset();
      range.ParagraphFormat.Reset();
    });
  };

  registry.skills['text.superscript'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Superscript = params.enabled ? -1 : 0;
    });
  };

  registry.skills['text.subscript'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.Font.Subscript = params.enabled ? -1 : 0;
    });
  };

  registry.skills['text.highlight'].execute = function (params, rangeOrIndices) {
    var colorMap = {
      'yellow': 7, 'green': 4, 'cyan': 5, 'magenta': 6,
      'blue': 2, 'red': 3, 'darkBlue': 9, 'teal': 10,
      'none': 0
    };
    var wpsColor = colorMap[params.color] || 7; // default yellow
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.HighlightColorIndex = wpsColor;
    });
  };

  // ── 段落技能注册（全部使用 executeOnTarget 批量操作 Range）──

  registry.skills['paragraph.alignment.set'].execute = function (params, rangeOrIndices) {
    var alignment = ALIGNMENT_MAP[params.alignment];
    if (alignment === undefined) alignment = 3; // justify
    executeOnTarget(params, rangeOrIndices, function (range) {
      range.ParagraphFormat.Alignment = alignment;
    });
  };

  registry.skills['paragraph.spacing.set'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      var pf = range.ParagraphFormat;
      if (params.spaceBefore !== undefined) pf.SpaceBefore = params.spaceBefore;
      if (params.spaceAfter !== undefined) pf.SpaceAfter = params.spaceAfter;
    });
  };

  registry.skills['paragraph.line_spacing.set'].execute = function (params, rangeOrIndices) {
    var rule = LINE_SPACING_RULE_MAP[params.mode];
    if (rule === undefined) rule = 5; // multiple
    executeOnTarget(params, rangeOrIndices, function (range) {
      var pf = range.ParagraphFormat;
      pf.LineSpacingRule = rule;
      if (params.mode === 'multiple') {
        pf.LineSpacing = params.value * 12;
      } else {
        pf.LineSpacing = params.value;
      }
    });
  };

  registry.skills['paragraph.indent.set'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      var pf = range.ParagraphFormat;
      if (params.firstLineChars !== undefined) {
        pf.CharacterUnitFirstLineIndent = params.firstLineChars;
      }
      if (params.firstLinePoints !== undefined) {
        pf.FirstLineIndent = params.firstLinePoints;
      }
      if (params.hanging !== undefined) {
        pf.CharacterUnitFirstLineIndent = -Math.abs(params.hanging);
      }
      if (params.left !== undefined) {
        pf.LeftIndent = params.left;
      }
      if (params.right !== undefined) {
        pf.RightIndent = params.right;
      }
    });
  };

  registry.skills['paragraph.columns.set'].execute = function (params, rangeOrIndices) {
    // 分栏通常对选区或节操作
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;
    var range = Array.isArray(rangeOrIndices) ? doc.Content : (rangeOrIndices || doc.Content);

    range.PageSetup.TextColumns.SetCount(params.count);
    if (params.spacing !== undefined) {
      range.PageSetup.TextColumns.Spacing = params.spacing;
    }
  };

  registry.skills['paragraph.borders.set'].execute = function (params, rangeOrIndices) {
    // 段落边框 - 基础实现
    executeOnTarget(params, rangeOrIndices, function (range) {
      // borderType: "box", "top", "bottom", "left", "right"
      var borders = range.ParagraphFormat.Borders;
      if (params.borderType === 'box') {
        borders.Enable = true;
      }
    });
  };

  registry.skills['paragraph.numbering.set'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      var lf = range.ListFormat;
      lf.ApplyNumberDefault();
    });
  };

  registry.skills['paragraph.bullets.set'].execute = function (params, rangeOrIndices) {
    executeOnTarget(params, rangeOrIndices, function (range) {
      var lf = range.ListFormat;
      lf.ApplyBulletDefault();
    });
  };
})();
