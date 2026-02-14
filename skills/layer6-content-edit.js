/**
 * layer6-content-edit.js
 * 第六层：基础文本编辑技能
 * 提供插入、删除、替换、追加等基础文本编辑能力
 */

/* global window, wps */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  if (!registry) return;

  // ── 安全阈值配置 ──
  var CONTENT_LIMITS = {
    MAX_DELETE_CHARS: 500,           // 单次删除最大字符数
    MAX_DELETE_RATIO: 0.05,          // 删除不超过文档总长 5%
    MAX_INSERT_CHARS: 2000,          // 单次插入最大字符数
    MAX_REPLACE_COUNT: 10,           // 单次替换最大次数
    KEY_AREA_PROTECT: 100            // 文档开头保护字符数
  };

  // ── 辅助函数 ──

  /**
   * 获取文档总字符数
   */
  function getDocumentCharCount(doc) {
    try {
      return doc.Content.Text.length;
    } catch (e) {
      return 0;
    }
  }

  /**
   * 检查删除是否超出安全阈值
   */
  function validateDelete(start, end, doc) {
    var deleteChars = end - start;
    var totalChars = getDocumentCharCount(doc);

    // 检查绝对上限
    if (deleteChars > CONTENT_LIMITS.MAX_DELETE_CHARS) {
      throw new Error('删除内容过长：' + deleteChars + ' 字符，超过上限 ' + CONTENT_LIMITS.MAX_DELETE_CHARS + ' 字符');
    }

    // 检查相对比例
    if (totalChars > 0 && deleteChars / totalChars > CONTENT_LIMITS.MAX_DELETE_RATIO) {
      throw new Error('删除比例过高：' + (deleteChars / totalChars * 100).toFixed(1) + '%，超过上限 ' + (CONTENT_LIMITS.MAX_DELETE_RATIO * 100) + '%');
    }

    // 检查是否在保护区域（文档开头）
    if (start < CONTENT_LIMITS.KEY_AREA_PROTECT) {
      throw new Error('文档开头 ' + CONTENT_LIMITS.KEY_AREA_PROTECT + ' 字符为保护区域，不可删除');
    }

    return true;
  }

  /**
   * 检查插入是否超出安全阈值
   */
  function validateInsert(text) {
    if (!text || text.length === 0) {
      throw new Error('插入内容不能为空');
    }
    if (text.length > CONTENT_LIMITS.MAX_INSERT_CHARS) {
      throw new Error('插入内容过长：' + text.length + ' 字符，超过上限 ' + CONTENT_LIMITS.MAX_INSERT_CHARS + ' 字符');
    }
    return true;
  }

  // ── content.insert ──
  // 在指定位置插入文本
  registry.skills['content.insert'] = {
    id: 'content.insert',
    layer: 6,
    description: '在指定位置插入文本',
    params: {
      position: { type: 'enum', required: true, values: ['start', 'end', 'cursor', 'before_target', 'after_target'] },
      text: { type: 'string', required: true },
      targetText: { type: 'string', required: false },  // 用于 before_target/after_target
      offset: { type: 'number', required: false }        // 偏移量
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      // 验证插入内容
      validateInsert(params.text);

      var range;
      var position = params.position || 'end';

      if (position === 'start') {
        // 文档开头插入
        range = doc.Content;
        range.Collapse(1); // wdCollapseStart
        range.Text = params.text;
      } else if (position === 'end') {
        // 文档末尾追加
        range = doc.Content;
        range.Collapse(0); // wdCollapseEnd
        range.Text = params.text;
      } else if (position === 'cursor') {
        // 当前位置插入（光标位置）
        range = app.Selection.Range;
        range.Text = params.text;
      } else if (position === 'before_target' || position === 'after_target') {
        // 在目标文本前后插入
        var targetText = params.targetText;
        if (!targetText) {
          throw new Error('before_target/after_target 模式需要指定 targetText');
        }
        // 查找目标文本
        var findResult = doc.Content.Find;
        findResult.Text = targetText;
        findResult.Forward = true;
        findResult.Wrap = 1; // wdFindStop

        if (findResult.Execute()) {
          var foundRange = findResult.Found;
          if (position === 'before_target') {
            range = foundRange.Duplicate;
            range.Collapse(1);
            range.Text = params.text;
          } else {
            range = foundRange.Duplicate;
            range.Collapse(0);
            range.Text = params.text;
          }
        } else {
          throw new Error('未找到目标文本: ' + targetText);
        }
      } else {
        throw new Error('无效的插入位置: ' + position);
      }

      console.log('[content.insert] Inserted ' + params.text.length + ' characters at ' + position);
      return { success: true, charsInserted: params.text.length };
    }
  };

  // ── content.delete ──
  // 删除指定范围的文本
  registry.skills['content.delete'] = {
    id: 'content.delete',
    layer: 6,
    description: '删除指定范围的文本',
    params: {
      start: { type: 'number', required: false },
      end: { type: 'number', required: false },
      targetText: { type: 'string', required: false },  // 删除匹配的文本
      count: { type: 'number', required: false }        // 删除字符数（从当前位置）
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      var range;
      var deletedChars = 0;

      if (params.targetText) {
        // 删除匹配的目标文本
        var findResult = doc.Content.Find;
        findResult.Text = params.targetText;
        findResult.Forward = true;
        findResult.Wrap = 1; // wdFindStop

        var deleteCount = 0;
        while (findResult.Execute() && deleteCount < CONTENT_LIMITS.MAX_REPLACE_COUNT) {
          var foundRange = findResult.Found;
          var startPos = foundRange.Start;
          var endPos = foundRange.End;

          // 验证删除安全性
          validateDelete(startPos, endPos, doc);

          foundRange.Text = '';
          deletedChars += (endPos - startPos);
          deleteCount++;
        }

        if (deleteCount === 0) {
          throw new Error('未找到可删除的文本: ' + params.targetText);
        }
        console.log('[content.delete] Deleted ' + deletedChars + ' characters, ' + deleteCount + ' matches');
      } else if (typeof params.start === 'number' && typeof params.end === 'number') {
        // 删除指定字符范围
        validateDelete(params.start, params.end, doc);

        range = doc.Range(params.start, params.end);
        deletedChars = params.end - params.start;
        range.Text = '';

        console.log('[content.delete] Deleted characters ' + params.start + ' to ' + params.end);
      } else if (typeof params.count === 'number') {
        // 从文档开头删除指定字符数
        var totalChars = getDocumentCharCount(doc);
        var deleteFrom = Math.min(CONTENT_LIMITS.KEY_AREA_PROTECT, totalChars);
        var deleteTo = Math.min(deleteFrom + params.count, totalChars);

        validateDelete(deleteFrom, deleteTo, doc);

        range = doc.Range(deleteFrom, deleteTo);
        deletedChars = deleteTo - deleteFrom;
        range.Text = '';

        console.log('[content.delete] Deleted ' + deletedChars + ' characters from position ' + deleteFrom);
      } else {
        throw new Error('请指定删除范围：start+end, targetText, 或 count');
      }

      return { success: true, charsDeleted: deletedChars };
    }
  };

  // ── content.replace ──
  // 替换文本
  registry.skills['content.replace'] = {
    id: 'content.replace',
    layer: 6,
    description: '替换文本内容',
    params: {
      oldText: { type: 'string', required: true },
      newText: { type: 'string', required: true },
      replaceAll: { type: 'boolean', required: false, default: true },
      caseSensitive: { type: 'boolean', required: false, default: false }
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      if (!params.oldText || params.oldText.length === 0) {
        throw new Error('oldText 不能为空');
      }

      // 验证新文本长度
      validateInsert(params.newText);

      var findResult = doc.Content.Find;
      findResult.Text = params.oldText;
      findResult.Forward = true;
      findResult.Wrap = 1; // wdFindStop
      findResult.MatchCase = params.caseSensitive || false;
      findResult.MatchWholeWord = false;

      var replaceCount = 0;
      var maxReplace = params.replaceAll ? CONTENT_LIMITS.MAX_REPLACE_COUNT : 1;

      while (findResult.Execute() && replaceCount < maxReplace) {
        findResult.Found.Text = params.newText;
        replaceCount++;
      }

      if (replaceCount === 0) {
        throw new Error('未找到可替换的文本: ' + params.oldText);
      }

      console.log('[content.replace] Replaced ' + replaceCount + ' occurrences');
      return { success: true, replaceCount: replaceCount };
    }
  };

  // ── content.append ──
  // 文档末尾追加文本
  registry.skills['content.append'] = {
    id: 'content.append',
    layer: 6,
    description: '在文档末尾追加文本',
    params: {
      text: { type: 'string', required: true },
      newLine: { type: 'boolean', required: false, default: true }  // 是否换行
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      // 验证插入内容
      validateInsert(params.text);

      var range = doc.Content;
      range.Collapse(0); // wdCollapseEnd

      // 如果需要换行，先插入换行符
      if (params.newLine !== false) {
        // 检查文档是否为空
        if (range.Text.length > 0) {
          range.Text = '\n' + params.text;
        } else {
          range.Text = params.text;
        }
      } else {
        range.Text = params.text;
      }

      console.log('[content.append] Appended ' + params.text.length + ' characters');
      return { success: true, charsAppended: params.text.length };
    }
  };

  // ── content.find ──
  // 查找文本（不修改，仅定位）
  registry.skills['content.find'] = {
    id: 'content.find',
    layer: 6,
    description: '查找文本并返回位置信息',
    params: {
      text: { type: 'string', required: true },
      caseSensitive: { type: 'boolean', required: false, default: false }
    },
    execute: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;

      if (!params.text || params.text.length === 0) {
        throw new Error('text 不能为空');
      }

      var findResult = doc.Content.Find;
      findResult.Text = params.text;
      findResult.Forward = true;
      findResult.Wrap = 1; // wdFindStop
      findResult.MatchCase = params.caseSensitive || false;

      var foundPositions = [];
      var count = 0;
      var maxFind = 20; // 最多返回20个位置

      while (findResult.Execute() && count < maxFind) {
        foundPositions.push({
          start: findResult.Found.Start,
          end: findResult.Found.End,
          text: findResult.Found.Text
        });
        count++;
      }

      console.log('[content.find] Found ' + foundPositions.length + ' occurrences of "' + params.text + '"');
      return { success: true, foundCount: foundPositions.length, positions: foundPositions };
    }
  };

  console.log('[Layer6] Content edit skills registered');
})();
