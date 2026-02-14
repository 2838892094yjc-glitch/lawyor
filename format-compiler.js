/**
 * format-compiler.js
 * 编译器：将 AI 返回的 JSON 字符串编译为可执行的内部计划
 */

/* global window */

(function () {
  'use strict';

  var registry = window.FormatSkillRegistry;
  var CHINESE_FONT_SIZE_MAP = registry ? registry.CHINESE_FONT_SIZE_MAP : {};
  var FONT_NAME_MAP = registry ? registry.FONT_NAME_MAP : {};

  window.FormatCompiler = {
    /**
     * 编译 AI 输出为执行计划
     * @param {string} jsonString - AI 返回的 JSON 字符串
     * @returns {{ plan: string, commands: Array, errors: Array }}
     */
    compile: function (jsonString) {
      // 1. JSON 解析
      var parsed;
      try {
        parsed = JSON.parse(jsonString);
      } catch (e) {
        // 尝试修复常见 JSON 问题
        parsed = this._attemptJsonFix(jsonString);
        if (!parsed) {
          return { plan: null, commands: [], errors: [{ code: 'PARSE_ERROR', message: 'JSON 解析失败: ' + e.message }] };
        }
      }

      // 2. 规范化
      var commands = (parsed.commands || []).map(function (cmd, i) {
        return {
          id: cmd.id || ('cmd_' + (i + 1)),
          skill: cmd.skill,
          params: cmd.params || {},
          description: cmd.description || cmd.skill,
          _index: i
        };
      });

      // 3. 单位规范化（确保所有值为 points）
      for (var i = 0; i < commands.length; i++) {
        this._normalizeUnits(commands[i]);
        this._normalizeTarget(commands[i]);
      }

      return {
        plan: parsed.plan || '',
        commands: commands,
        errors: []
      };
    },

    /**
     * 尝试修复常见 JSON 格式问题
     */
    _attemptJsonFix: function (str) {
      if (!str || typeof str !== 'string') return null;

      // 去掉 markdown 代码块标记
      var cleaned = str.replace(/^```json\s*\n?/m, '').replace(/\n?```\s*$/m, '');
      try { return JSON.parse(cleaned); } catch (e) { /* continue */ }

      // 去掉前后非 JSON 文本
      var match = cleaned.match(/\{[\s\S]*\}/);
      if (match) {
        try { return JSON.parse(match[0]); } catch (e) { /* continue */ }
      }

      // 尝试更深层的提取
      for (var start = 0; start < cleaned.length; start++) {
        if (cleaned[start] !== '{') continue;
        var depth = 0;
        var inString = false;
        var escaped = false;

        for (var i = start; i < cleaned.length; i++) {
          var ch = cleaned[i];
          if (inString) {
            if (escaped) { escaped = false; }
            else if (ch === '\\') { escaped = true; }
            else if (ch === '"') { inString = false; }
            continue;
          }
          if (ch === '"') { inString = true; continue; }
          if (ch === '{') depth++;
          if (ch === '}') depth--;
          if (depth === 0) {
            try { return JSON.parse(cleaned.slice(start, i + 1)); } catch (e) { break; }
          }
        }
      }

      return null;
    },

    /**
     * 确保度量值都是 points，字体名是 API 名
     */
    _normalizeUnits: function (cmd) {
      var p = cmd.params;

      // 字号：中文名转磅值
      if (typeof p.fontSize === 'string') {
        p.fontSize = CHINESE_FONT_SIZE_MAP[p.fontSize] || parseFloat(p.fontSize) || 12;
      }
      if (typeof p.size === 'string') {
        p.size = CHINESE_FONT_SIZE_MAP[p.size] || parseFloat(p.size) || 12;
      }

      // 中文字体名转 API 名
      if (p.zhFont && FONT_NAME_MAP[p.zhFont]) {
        p.zhFont = FONT_NAME_MAP[p.zhFont];
      }
      if (p.enFont && FONT_NAME_MAP[p.enFont]) {
        p.enFont = FONT_NAME_MAP[p.enFont];
      }
      if (p.fontName && FONT_NAME_MAP[p.fontName]) {
        p.fontName = FONT_NAME_MAP[p.fontName];
      }
    },

    /**
     * 为缺省 target 的指令补充默认值
     */
    _normalizeTarget: function (cmd) {
      var skill = cmd.skill || '';
      // custom.execute 不需要 target 规范化
      if (skill === 'custom.execute') return;
      // 页面级操作不需要 target
      if (skill.startsWith('page.') || skill.startsWith('header.') ||
          skill.startsWith('footer.') || skill.startsWith('page_number.') ||
          skill.startsWith('detect.')) {
        return;
      }
      // 分节操作需要 target
      if (skill.startsWith('section.') && cmd.params.target) {
        return;
      }
      // 其他操作如果没有 target，默认 document
      if (!cmd.params.target) {
        cmd.params.target = { type: 'document' };
      }
    }
  };
})();
