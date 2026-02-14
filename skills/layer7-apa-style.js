/**
 * layer7-apa-style.js
 * 第七层：APA 格式专用技能
 * 提供 APA 论文格式的完整处理能力
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) {
    console.warn('[Layer7-APA] FormatSkillRegistry not found');
    return;
  }

  // 辅助函数：执行操作到目标段落
  function executeOnParagraphs(params, callback) {
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;
    var paras = doc.Content.Paragraphs;

    var startPara = params.startPara || 1;
    var endPara = params.endPara || paras.Count;

    for (var i = startPara; i <= endPara && i <= paras.Count; i++) {
      try {
        var para = paras(i);
        callback(para);
      } catch (e) {
        console.warn('[Layer7-APA] Error on paragraph ' + i + ': ' + e.message);
      }
    }
  }

  // 辅助函数：查找参考文献页面
  function findReferencesSection() {
    var app = wps.WpsApplication();
    var doc = app.ActiveDocument;
    var paras = doc.Content.Paragraphs;

    for (var i = 1; i <= paras.Count; i++) {
      try {
        var paraText = paras(i).Range.Text.trim().toLowerCase();
        if (paraText === 'references' || paraText === 'reference' || paraText === '参考文献') {
          return i;
        }
      } catch (e) {
        // continue
      }
    }
    return null;
  }

  // === APA 主技能 ===

  registry.skills['apa.apply'] = {
    id: 'apa.apply',
    layer: 7,
    description: '应用完整的 APA 格式（第七版）',
    params: {
      paperSize: { type: 'enum', values: ['A4', 'Letter'], default: 'A4', description: '纸张大小' },
      font: { type: 'enum', values: ['times', 'arial'], default: 'times', description: '正文字体' },
      lineSpacing: { type: 'number', default: 2.0, description: '行距（倍数）' },
      includeReferences: { type: 'boolean', default: true, description: '是否格式化参考文献' }
    },
    execute: function (params, rangeOrIndices) {
      params = params || {};
      var self = this;

      // 1. 设置页面
      var pageSizeCmd = {
        skill: 'page.size.set',
        params: { paperSize: params.paperSize || 'A4' }
      };
      registry.executeSkill(pageSizeCmd.skill, pageSizeCmd.params);

      var marginsCmd = {
        skill: 'page.margins.set',
        params: { top: 72, bottom: 72, left: 72, right: 72 } // 1英寸 = 72pt
      };
      registry.executeSkill(marginsCmd.skill, marginsCmd.params);

      // 2. 设置字体
      var fontName = params.font === 'arial' ? 'Arial' : 'Times New Roman';
      var fontCmd = {
        skill: 'text.font.set',
        params: { fontName: fontName, fontSize: 12 }
      };
      registry.executeSkill(fontCmd.skill, fontCmd.params);

      // 3. 设置行距
      var spacingCmd = {
        skill: 'paragraph.line_spacing.set',
        params: { mode: 'multiple', value: params.lineSpacing || 2.0 }
      };
      registry.executeSkill(spacingCmd.skill, spacingCmd.params);

      // 4. 设置首行缩进
      var indentCmd = {
        skill: 'paragraph.indent.set',
        params: { firstLineChars: 36 } // 0.5英寸 = 36字符
      };
      registry.executeSkill(indentCmd.skill, indentCmd.params);

      // 5. 格式化参考文献（如果需要）
      if (params.includeReferences) {
        var refCmd = {
          skill: 'apa.references.format',
          params: {}
        };
        registry.executeSkill(refCmd.skill, refCmd.params);
      }

      console.log('[Layer7-APA] APA format applied');
    }
  };

  // === APA 参考文献格式化 ===

  registry.skills['apa.references.format'] = {
    id: 'apa.references.format',
    layer: 7,
    description: '格式化参考文献页（APA 悬挂缩进）',
    params: {
      title: { type: 'string', default: 'References', description: '参考文献标题' },
      hangingIndent: { type: 'boolean', default: true, description: '是否使用悬挂缩进' },
      indentSize: { type: 'number', default: 36, description: '悬挂缩进大小（字符单位，0.5英寸≈36）' },
      titleBold: { type: 'boolean', default: true, description: '标题是否加粗' },
      titleCenter: { type: 'boolean', default: true, description: '标题是否居中' }
    },
    execute: function (params, rangeOrIndices) {
      params = params || {};

      // 1. 查找参考文献页面
      var refStart = findReferencesSection();
      if (!refStart) {
        console.warn('[Layer7-APA] References section not found, creating new page');
        // 创建参考文献页面
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        doc.Content.InsertAfter('\n\n' + (params.title || 'References'));
        refStart = doc.Content.Paragraphs.Count - 1;
      }

      // 2. 设置标题格式
      var titlePara = doc.Content.Paragraphs(refStart);
      if (params.titleBold) {
        titlePara.Range.Font.Bold = true;
      }
      if (params.titleCenter) {
        titlePara.ParagraphFormat.Alignment = 1; // center
      }

      // 3. 设置悬挂缩进（从第二段开始）
      if (params.hangingIndent) {
        var paras = doc.Content.Paragraphs;
        for (var i = refStart + 1; i <= paras.Count; i++) {
          try {
            var para = paras(i);
            var paraText = para.Range.Text.trim();
            // 跳过空段落
            if (paraText.length === 0) continue;
            // 跳过明显的非参考文献内容
            if (paraText.match(/^(摘要|abstract|关键词|keywords|引言|introduction)/i)) continue;

            para.ParagraphFormat.CharacterUnitFirstLineIndent = -params.indentSize;
            para.ParagraphFormat.FirstLineIndent = -params.indentSize;
          } catch (e) {
            console.warn('[Layer7-APA] Error formatting paragraph ' + i + ': ' + e.message);
          }
        }
      }

      console.log('[Layer7-APA] References formatted with hanging indent');
    }
  };

  // === APA 封面格式 ===

  registry.skills['apa.cover.format'] = {
    id: 'apa.cover.format',
    layer: 7,
    description: '格式化 APA 封面',
    params: {
      title: { type: 'string', description: '论文标题' },
      author: { type: 'string', description: '作者' },
      affiliation: { type: 'string', description: '单位' },
      course: { type: 'string', description: '课程' },
      instructor: { type: 'string', description: '指导教师' },
      date: { type: 'string', description: '日期' }
    },
    execute: function (params, rangeOrIndices) {
      params = params || {};

      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      // 在文档开头插入封面内容
      var coverContent = [];

      if (params.title) {
        coverContent.push(params.title + '\n\n');
      }

      var metaLines = [];
      if (params.affiliation) metaLines.push(params.affiliation);
      if (params.course) metaLines.push(params.course);
      if (params.instructor) metaLines.push('Instructor: ' + params.instructor);
      if (params.date) metaLines.push(params.date);

      if (metaLines.length > 0) {
        coverContent.push(metaLines.join('\n'));
      }

      // 插入封面（需要在文档最前面）
      if (coverContent.length > 0) {
        var range = doc.Range(0, 0);
        range.Text = coverContent.join('\n');

        // 设置标题居中加粗
        var titlePara = doc.Content.Paragraphs(1);
        titlePara.ParagraphFormat.Alignment = 1; // center
        titlePara.Range.Font.Bold = true;
        titlePara.Range.Font.Size = 12;

        // 设置元信息居中
        for (var i = 2; i <= doc.Content.Paragraphs.Count && i <= 5; i++) {
          doc.Content.Paragraphs(i).ParagraphFormat.Alignment = 1;
        }
      }

      console.log('[Layer7-APA] Cover page formatted');
    }
  };

  // === APA 摘要页格式 ===

  registry.skills['apa.abstract.format'] = {
    id: 'apa.abstract.format',
    layer: 7,
    description: '格式化 APA 摘要页',
    params: {
      includeKeywords: { type: 'boolean', default: true, description: '是否包含关键词' },
      keywords: { type: 'string', description: '关键词（用逗号分隔）' }
    },
    execute: function (params, rangeOrIndices) {
      params = params || {};

      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      // 查找摘要标题
      var abstractFound = false;
      var paras = doc.Content.Paragraphs;

      for (var i = 1; i <= paras.Count; i++) {
        try {
          var paraText = paras(i).Range.Text.trim().toLowerCase();
          if (paraText === 'abstract' || paraText === '摘要') {
            abstractFound = true;

            // 标题居中加粗
            paras(i).ParagraphFormat.Alignment = 1;
            paras(i).Range.Font.Bold = true;

            // 标题后空一行
            if (i < paras.Count) {
              paras(i + 1).Range.InsertBefore('\n');
            }

            // 设置缩进
            if (params.includeKeywords) {
              // 摘要正文缩进
              for (var j = i + 1; j <= paras.Count; j++) {
                var pText = paras(j).Range.Text.trim();
                if (pText.toLowerCase().startsWith('keywords') || pText.toLowerCase().startsWith('关键词')) {
                  break;
                }
                if (pText.length > 0) {
                  paras(j).ParagraphFormat.FirstLineIndent = 36; // 0.5英寸
                }
              }
            }
            break;
          }
        } catch (e) {
          // continue
        }
      }

      if (!abstractFound) {
        console.warn('[Layer7-APA] Abstract section not found');
      } else {
        console.log('[Layer7-APA] Abstract page formatted');
      }
    }
  };

  console.log('[Layer7] APA style skills registered');
})();
