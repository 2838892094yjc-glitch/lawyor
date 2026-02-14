/**
 * format-validator.js
 * 验证器：3阶段管道 (结构 → 语义 → 可行性)
 */

/* global window, wps */

(function () {
  'use strict';

  window.FormatValidator = {
    /**
     * 完整验证管道
     * @param {object} compiled - FormatCompiler.compile() 的输出
     * @returns {{ valid: boolean, errors: Array, warnings: Array }}
     */
    validate: function (compiled) {
      var errors = [];
      var warnings = [];

      // 如果编译阶段就有错误
      if (compiled.errors && compiled.errors.length > 0) {
        return { valid: false, errors: compiled.errors, warnings: warnings };
      }

      // Phase 1: 结构验证
      this._validateStructure(compiled.commands, errors);
      if (errors.length > 0) return { valid: false, errors: errors, warnings: warnings };

      // Phase 2: 语义验证
      this._validateSemantics(compiled.commands, errors, warnings);
      if (errors.length > 0) return { valid: false, errors: errors, warnings: warnings };

      // Phase 2.5: 自定义代码安全验证
      this._validateCustomCodeSafety(compiled.commands, errors, warnings);
      if (errors.length > 0) return { valid: false, errors: errors, warnings: warnings };

      // Phase 2.6: custom.execute 兜底策略拦截
      this._validateCustomFallbackPolicy(compiled.commands, errors, warnings);
      if (errors.length > 0) return { valid: false, errors: errors, warnings: warnings };

      // Phase 3: 可行性验证（需要文档上下文）
      this._validateFeasibility(compiled.commands, errors, warnings);

      return { valid: errors.length === 0, errors: errors, warnings: warnings };
    },

    /**
     * Phase 1: 结构验证
     */
    _validateStructure: function (commands, errors) {
      if (!Array.isArray(commands) || commands.length === 0) {
        // 空指令列表不一定是错误（可能是"无法执行"的情况）
        return;
      }

      for (var i = 0; i < commands.length; i++) {
        var cmd = commands[i];
        if (!cmd.skill || typeof cmd.skill !== 'string') {
          errors.push({ level: 'error', code: 'MISSING_SKILL', index: i,
            message: '指令 ' + (i + 1) + ': 缺少 skill 字段' });
        }
        if (!cmd.params || typeof cmd.params !== 'object') {
          errors.push({ level: 'error', code: 'MISSING_PARAMS', index: i,
            message: '指令 ' + (i + 1) + ': 缺少 params 字段' });
        }
      }
    },

    /**
     * Phase 2: 语义验证
     */
    _validateSemantics: function (commands, errors, warnings) {
      var registry = window.FormatSkillRegistry;
      if (!registry || !registry.skills) return;

      var skillDefs = registry.skills;

      for (var i = 0; i < commands.length; i++) {
        var cmd = commands[i];

        // 技能存在性
        var skillDef = skillDefs[cmd.skill];
        if (!skillDef) {
          errors.push({ level: 'error', code: 'UNKNOWN_SKILL', index: i,
            message: '指令 ' + (i + 1) + ': 未知技能 "' + cmd.skill + '"' });
          continue;
        }

        // 必填参数
        if (skillDef.params) {
          var paramKeys = Object.keys(skillDef.params);
          for (var p = 0; p < paramKeys.length; p++) {
            var key = paramKeys[p];
            var def = skillDef.params[key];
            if (def.required && !(key in cmd.params)) {
              errors.push({ level: 'error', code: 'MISSING_PARAM', index: i,
                message: '指令 ' + (i + 1) + ': 缺少必填参数 "' + key + '"' });
            }
          }
        }

        // 枚举值验证
        if (skillDef.params && cmd.params) {
          var cmdParamKeys = Object.keys(cmd.params);
          for (var q = 0; q < cmdParamKeys.length; q++) {
            var pKey = cmdParamKeys[q];
            var val = cmd.params[pKey];
            var paramDef = skillDef.params[pKey];
            if (paramDef && paramDef.type === 'enum' && paramDef.values) {
              if (!paramDef.values.includes(val)) {
                errors.push({ level: 'error', code: 'INVALID_ENUM', index: i,
                  message: '指令 ' + (i + 1) + ': "' + pKey + '" 值 "' + val + '" 不在合法范围 [' + paramDef.values.join(', ') + ']' });
              }
            }
          }
        }

        // target 结构验证
        if (cmd.params && cmd.params.target) {
          this._validateTarget(cmd.params.target, i, errors);
        }

        // 数值范围警告
        if (cmd.params) {
          if (cmd.params.fontSize && (cmd.params.fontSize < 5 || cmd.params.fontSize > 96)) {
            warnings.push({ level: 'warning', code: 'UNUSUAL_VALUE', index: i,
              message: '指令 ' + (i + 1) + ': 字号 ' + cmd.params.fontSize + 'pt 超出常规范围 (5-96)' });
          }
          if (cmd.params.size && (cmd.params.size < 5 || cmd.params.size > 96)) {
            warnings.push({ level: 'warning', code: 'UNUSUAL_VALUE', index: i,
              message: '指令 ' + (i + 1) + ': 字号 ' + cmd.params.size + 'pt 超出常规范围 (5-96)' });
          }
        }
      }
    },

    /**
     * Phase 3: 可行性验证（需要文档上下文）
     */
    _validateFeasibility: function (commands, errors, warnings) {
      if (typeof wps === 'undefined') return;

      var app;
      try { app = wps.WpsApplication(); } catch (e) { return; }
      var doc = app.ActiveDocument;
      if (!doc) {
        errors.push({ level: 'error', code: 'NO_DOCUMENT', message: '没有打开的文档' });
        return;
      }

      var paraCount, tableCount;
      try { paraCount = doc.Paragraphs.Count; } catch (e) { paraCount = 0; }
      try { tableCount = doc.Tables.Count; } catch (e) { tableCount = 0; }

      for (var i = 0; i < commands.length; i++) {
        var cmd = commands[i];
        var target = cmd.params ? cmd.params.target : null;
        if (!target) continue;

        if (target.type === 'paragraph_index' && target.index > paraCount) {
          errors.push({ level: 'error', code: 'TARGET_OUT_OF_RANGE', index: i,
            message: '指令 ' + (i + 1) + ': 段落 ' + target.index + ' 不存在（共 ' + paraCount + ' 段）' });
        }
        if (target.type === 'paragraph_range' && target.to > paraCount) {
          errors.push({ level: 'error', code: 'TARGET_OUT_OF_RANGE', index: i,
            message: '指令 ' + (i + 1) + ': 段落范围超出文档（to=' + target.to + ', 共 ' + paraCount + ' 段）' });
        }
        if (target.type === 'table_index' && target.index > tableCount) {
          errors.push({ level: 'error', code: 'TARGET_OUT_OF_RANGE', index: i,
            message: '指令 ' + (i + 1) + ': 表格 ' + target.index + ' 不存在（共 ' + tableCount + ' 个）' });
        }
        if (target.type === 'selection') {
          try {
            if (app.Selection.Start === app.Selection.End) {
              warnings.push({ level: 'warning', code: 'EMPTY_SELECTION', index: i,
                message: '指令 ' + (i + 1) + ': 当前没有选区，将作用于光标位置' });
            }
          } catch (e) { /* ignore */ }
        }

        // 性能警告
        if (target.type === 'all_paragraphs' && paraCount > 200) {
          warnings.push({ level: 'warning', code: 'LARGE_RANGE', index: i,
            message: '指令 ' + (i + 1) + ': 将影响 ' + paraCount + ' 个段落，执行可能较慢' });
        }
      }
    },

    /**
     * 验证 target 结构
     */
    _validateTarget: function (target, cmdIndex, errors) {
      var validTypes = ['selection', 'document', 'all_paragraphs', 'paragraph_index',
        'paragraph_range', 'search', 'heading_level', 'role', 'section_index', 'table_index'];

      if (!target.type || !validTypes.includes(target.type)) {
        errors.push({ level: 'error', code: 'INVALID_TARGET', index: cmdIndex,
          message: '指令 ' + (cmdIndex + 1) + ': 无效的目标类型 "' + target.type + '"' });
        return;
      }

      // 特定类型需要额外参数
      if (target.type === 'paragraph_index' && !target.index) {
        errors.push({ level: 'error', code: 'MISSING_TARGET_PARAM', index: cmdIndex,
          message: '指令 ' + (cmdIndex + 1) + ': paragraph_index 目标需要 index 参数' });
      }
      if (target.type === 'paragraph_range' && (!target.from || !target.to)) {
        errors.push({ level: 'error', code: 'MISSING_TARGET_PARAM', index: cmdIndex,
          message: '指令 ' + (cmdIndex + 1) + ': paragraph_range 目标需要 from 和 to 参数' });
      }
      if (target.type === 'search' && !target.text) {
        errors.push({ level: 'error', code: 'MISSING_TARGET_PARAM', index: cmdIndex,
          message: '指令 ' + (cmdIndex + 1) + ': search 目标需要 text 参数' });
      }
      if (target.type === 'role' && !target.role) {
        errors.push({ level: 'error', code: 'MISSING_TARGET_PARAM', index: cmdIndex,
          message: '指令 ' + (cmdIndex + 1) + ': role 目标需要 role 参数' });
      }
    },

    /**
     * Phase 2.5: 自定义代码安全验证
     */
    _validateCustomCodeSafety: function (commands, errors, warnings) {
      var BANNED_APIS = [
        'fetch', 'XMLHttpRequest', 'eval', 'require', 'import',
        'process', 'setTimeout', 'setInterval', 'Function(',
        'WebSocket', 'ActiveXObject', 'localStorage', 'document.cookie'
      ];
      var MAX_CODE_LENGTH = 2000;

      for (var i = 0; i < commands.length; i++) {
        var cmd = commands[i];
        if (cmd.skill !== 'custom.execute') continue;

        var code = cmd.params && cmd.params.code;
        if (!code || typeof code !== 'string') {
          errors.push({ level: 'error', code: 'MISSING_CODE', index: i,
            message: '指令 ' + (i + 1) + ': custom.execute 缺少 code 参数' });
          continue;
        }

        // 代码长度检查
        if (code.length > MAX_CODE_LENGTH) {
          errors.push({ level: 'error', code: 'CODE_TOO_LONG', index: i,
            message: '指令 ' + (i + 1) + ': 代码长度 ' + code.length + ' 超出上限 ' + MAX_CODE_LENGTH + ' 字符' });
          continue;
        }

        // 禁止 API 检查（skills.xxx 为白名单，不参与匹配）
        var codeForCheck = code.replace(/skills\.\w+/g, 'SKILLS_CALL');
        for (var b = 0; b < BANNED_APIS.length; b++) {
          var api = BANNED_APIS[b];
          if (codeForCheck.indexOf(api) !== -1) {
            errors.push({ level: 'error', code: 'UNSAFE_CODE', index: i,
              message: '指令 ' + (i + 1) + ': 代码包含禁止的 API "' + api + '"',
              bannedApi: api });
          }
        }

        // 每条 custom.execute 产生一个 warning 提醒用户检查
        warnings.push({ level: 'warning', code: 'CUSTOM_CODE', index: i,
          message: '指令 ' + (i + 1) + ': 包含自定义代码，请检查安全性' });
      }
    },

    /**
     * Phase 2.6: custom.execute 兜底策略拦截
     * 可由前端重试机制据此要求模型改用内置技能。
     */
    _hasExplicitFallbackReason: function (cmd) {
      var desc = cmd && cmd.params && typeof cmd.params.description === 'string'
        ? cmd.params.description
        : '';
      if (!desc) return false;

      var hasBuiltinRef = /(内置技能|table\.|text\.|paragraph\.|page\.|style\.|toc\.|watermark\.)/i.test(desc);
      var hasLimitation = /(不支持|无法|不能|受限|缺少|暂无|覆盖不了|无法表达)/.test(desc);
      return hasBuiltinRef && hasLimitation;
    },

    _validateCustomFallbackPolicy: function (commands, errors, warnings) {
      var policyChecks = [
        {
          code: 'CUSTOM_TABLE_STYLE',
          pattern: /(table\.|Tables?\.|Borders?\.|Cell\(|Rows?\.|Columns?\.|AutoFit|Shading|表格|边框|单元格|三线表)/i,
          message: '涉及表格样式时必须使用内置 table.* 技能，不得使用 custom.execute',
          hint: '表格样式请改用 table.*（如 table.borders.set / table.cell_alignment.set）'
        },
        {
          code: 'CUSTOM_TEXT_STYLE',
          pattern: /(Font\.|fontSize|bold\b|italic\b|underline\b|字体|字号|加粗|斜体|下划线)/i,
          message: '涉及字体样式时必须使用内置 text.* 技能，不得使用 custom.execute',
          hint: '字体样式请改用 text.font.* 内置技能'
        },
        {
          code: 'CUSTOM_PARAGRAPH_STYLE',
          pattern: /(ParagraphFormat|LineSpacing|SpaceBefore|SpaceAfter|FirstLineIndent|Alignment|行距|缩进|段前|段后|两端对齐|居中|右对齐)/i,
          message: '涉及段落样式时必须使用内置 paragraph.* 技能，不得使用 custom.execute',
          hint: '段落样式请改用 paragraph.*（alignment/spacing/line_spacing/indent）'
        },
        {
          code: 'CUSTOM_PAGE_STYLE',
          pattern: /(PageSetup|页边距|纸张|横向|纵向|页眉|页脚|页码|分节|目录|水印|样式)/i,
          message: '涉及页面/样式类操作时必须使用内置技能，不得使用 custom.execute',
          hint: '页面与样式请改用 page.* / header.* / footer.* / page_number.* / style.* / toc.* / watermark.*'
        }
      ];

      for (var i = 0; i < commands.length; i++) {
        var cmd = commands[i];
        if (cmd.skill !== 'custom.execute') continue;

        var hasFallbackReason = this._hasExplicitFallbackReason(cmd);
        var pieces = [
          cmd.description || '',
          (cmd.params && cmd.params.description) || '',
          (cmd.params && cmd.params.code) || ''
        ];
        var scopeText = pieces.join('\n');

        for (var p = 0; p < policyChecks.length; p++) {
          var rule = policyChecks[p];
          if (rule.pattern.test(scopeText)) {
            // 表格样式已有内置 table.* 能力，禁止使用 custom.execute。
            if (rule.code === 'CUSTOM_TABLE_STYLE' || !hasFallbackReason) {
              errors.push({
                level: 'error',
                code: 'CUSTOM_NOT_FALLBACK',
                policyCode: rule.code,
                index: i,
                message: '指令 ' + (i + 1) + ': ' + rule.message,
                policyHint: rule.hint
              });
            } else {
              warnings.push({
                level: 'warning',
                code: 'CUSTOM_FALLBACK_ACCEPTED',
                index: i,
                message: '指令 ' + (i + 1) + ': 检测到已说明内置技能限制，允许 custom.execute 兜底执行'
              });
            }
            break;
          }
        }

        if (cmd.params && typeof cmd.params.description === 'string') {
          var desc = cmd.params.description;
          if (desc.indexOf('内置技能') === -1 && desc.indexOf('无法') === -1 && desc.indexOf('兜底') === -1) {
            warnings.push({
              level: 'warning',
              code: 'CUSTOM_REASON_MISSING',
              index: i,
              message: '指令 ' + (i + 1) + ': custom.execute 建议在 description 中写明内置技能无法覆盖的原因'
            });
          }
        }
      }
    }
  };
})();
