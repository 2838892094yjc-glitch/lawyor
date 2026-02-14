/**
 * format-execution-engine.js
 * 执行引擎：逐步执行已编译+验证的指令序列，配合 comparison log
 *
 * 性能关键点：
 * - captureState 同步执行，不走 wpsActionQueue（避免额外排队延迟）
 * - role/heading_level 目标的 detect 在进入队列前完成（避免嵌套死锁）
 * - cooldown 降到 20ms（WPS JSAPI 不需要长间隔）
 */

/* global window, wps, wpsActionQueue */

(function () {
  'use strict';

  var COOLDOWN_MS = 20;
  var _currentRunId = null; // 当前执行的 runId

  /**
   * 生成全局唯一 runId
   * 格式：{timestamp}-{uuidv4}，例如 20260213-550e8400-e29b-41d4-a716-446655440000
   */
  function generateRunId() {
    var now = new Date();
    var timestamp = now.getFullYear() +
      String(now.getMonth() + 1).padStart(2, '0') +
      String(now.getDate()).padStart(2, '0');
    // 简单 UUID v4 生成
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0;
      var v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
    return timestamp + '-' + uuid;
  }

  /**
   * 获取当前 runId
   */
  function getCurrentRunId() {
    return _currentRunId;
  }

  function log(msg) {
    var rid = _currentRunId || 'no-runId';
    console.log('[FormatEngine][' + rid + '] ' + msg);
  }
  function warn(msg) {
    var rid = _currentRunId || 'no-runId';
    console.warn('[FormatEngine][' + rid + '] ' + msg);
  }

  window.FormatExecutionEngine = {
    /**
     * 生成 runId（供外部调用）
     */
    generateRunId: generateRunId,

    run: async function (compiled, callbacks) {
      callbacks = callbacks || {};
      var onStepStart = callbacks.onStepStart;
      var onStepComplete = callbacks.onStepComplete;
      var onStepError = callbacks.onStepError;
      var onComplete = callbacks.onComplete;

      // 生成全局唯一 runId 并设置到模块级变量
      _currentRunId = generateRunId();
      var results = [];
      var totalCmds = compiled.commands.length;
      log('run() started, commands=' + totalCmds);

      // 执行阶段统一关闭屏幕刷新，减少 custom.execute 的重绘开销
      var screenUpdatingWasOn = true;
      var screenUpdatingChanged = false;
      try {
        var app = wps.WpsApplication();
        screenUpdatingWasOn = app.ScreenUpdating;
        if (totalCmds > 0 && screenUpdatingWasOn) {
          app.ScreenUpdating = false;
          screenUpdatingChanged = true;
          log('ScreenUpdating disabled (commands=' + totalCmds + ')');
        }
      } catch (e) {
        warn('ScreenUpdating toggle failed: ' + e.message);
      }

      try {
        // 预填充缓存：如有 role/heading_level 目标，提前触发统一扫描
        var needsRoleDetect = compiled.commands.some(function (c) {
          var t = c.params && c.params.target;
          return t && (t.type === 'role' || t.type === 'heading_level');
        });
        if (needsRoleDetect) {
          log('pre-filling structure cache...');
          await window.DocumentStructureCache.detect('body');
          log('structure cache pre-filled');
        }

        for (var i = 0; i < totalCmds; i++) {
          var cmd = compiled.commands[i];
          var startTime = Date.now();

          log('step ' + (i + 1) + '/' + totalCmds + ': ' + cmd.skill + ' - ' + (cmd.description || ''));
          if (onStepStart) onStepStart(i, cmd);

          try {
            // 1. 捕获执行前状态（同步，不排队）
            var captureBeforeStart = Date.now();
            var beforeState = this._captureStateDirect(cmd);
            var captureBeforeMs = Date.now() - captureBeforeStart;

            // 2. 执行技能
            var executeStart = Date.now();
            var executionMeta = await this._executeOneCommand(cmd);
            var executeMs = Date.now() - executeStart;

            // 3. 捕获执行后状态（同步，不排队）
            var captureAfterStart = Date.now();
            var afterState = this._captureStateDirect(cmd);
            var captureAfterMs = Date.now() - captureAfterStart;

            var compareStart = Date.now();
            var comparisonEntry = window.ComparisonLog.createEntry(
              cmd, beforeState, afterState, Date.now() - startTime
            );
            var compareMs = Date.now() - compareStart;

            // 添加 runId 和 stepId 用于全链路追踪
            comparisonEntry.runId = _currentRunId;
            comparisonEntry.stepId = i + 1;

            comparisonEntry.performance = {
              captureBeforeMs: captureBeforeMs,
              executeMs: executeMs,
              captureAfterMs: captureAfterMs,
              compareMs: compareMs,
              queueWaitMs: executionMeta && executionMeta.queueWaitMs ? executionMeta.queueWaitMs : 0,
              queueTaskMs: executionMeta && executionMeta.queueTaskMs ? executionMeta.queueTaskMs : 0,
              totalMs: Date.now() - startTime
            };

            if (executionMeta && executionMeta.customEffects) {
              comparisonEntry.customEffects = executionMeta.customEffects;
            }

            var effectCount = this._estimateEffectCount(comparisonEntry, executionMeta);
            comparisonEntry.effectCount = effectCount;
            var measurementAvailable = Object.keys(beforeState || {}).length > 0 ||
              Object.keys(afterState || {}).length > 0 ||
              !!(executionMeta && executionMeta.customEffects);
            if (measurementAvailable && effectCount === 0) {
              comparisonEntry.warnings = comparisonEntry.warnings || [];
              comparisonEntry.warnings.push('SUCCESS_ZERO_EFFECT');
              warn('step ' + (i + 1) + ' SUCCESS but zero-effect: ' + JSON.stringify({
                skill: cmd.skill,
                description: cmd.description || cmd.skill,
                performance: comparisonEntry.performance,
                customEffects: comparisonEntry.customEffects || null
              }));
            }

            window.ComparisonLog.addEntry(comparisonEntry);

            log('step ' + (i + 1) + ' perf: ' + JSON.stringify(comparisonEntry.performance));
            if (comparisonEntry.customEffects) {
              log('step ' + (i + 1) + ' custom-effects: ' + JSON.stringify(comparisonEntry.customEffects));

              // 任务D: 收集 custom.execute 成功样本用于技能进化
              if (cmd.skill === 'custom.execute' && window.SkillEvolutionManager) {
                try {
                  var stepResult = {
                    skill: cmd.skill,
                    params: cmd.params,
                    executionMeta: executionMeta || {},
                    comparison: comparisonEntry
                  };
                  var candidate = window.SkillEvolutionManager.collectCandidateFromExecution(stepResult);
                  if (candidate) {
                    log('Skill evolution: candidate collected ' + candidate.id);
                  }
                } catch (evolutionErr) {
                  warn('Skill evolution collect failed: ' + evolutionErr.message);
                }
              }
            }

            log('step ' + (i + 1) + ' OK (' + (Date.now() - startTime) + 'ms) changes=' + comparisonEntry.changes.length + ' effect=' + effectCount);
            results.push({ index: i, success: true, skill: cmd.skill, comparison: comparisonEntry, executionMeta: executionMeta || {} });
            if (onStepComplete) onStepComplete(i, cmd, comparisonEntry);

          } catch (err) {
            warn('step ' + (i + 1) + ' FAILED: ' + (err.message || String(err)));
            var errorEntry = {
              timestamp: new Date().toISOString(),
              commandId: cmd.id,
              skill: cmd.skill,
              description: cmd.description || cmd.skill,
              changes: [],
              status: 'failed',
              error: err.message || String(err),
              duration: (Date.now() - startTime) + 'ms'
            };
            window.ComparisonLog.addEntry(errorEntry);

            results.push({ index: i, success: false, skill: cmd.skill, error: err.message || String(err) });
            if (onStepError) onStepError(i, cmd, err);
          }
        }
      } finally {
        try {
          if (screenUpdatingChanged) {
            wps.WpsApplication().ScreenUpdating = screenUpdatingWasOn;
            log('ScreenUpdating restored');
          }
        } catch (e) { /* ignore */ }
      }

      log('run() finished, success=' + results.filter(function (r) { return r.success; }).length +
          ' failed=' + results.filter(function (r) { return !r.success; }).length);
      if (onComplete) onComplete(results);
      // 清空 runId
      var finishedRunId = _currentRunId;
      _currentRunId = null;
      // 返回结果时带上 runId 方便外部追踪
      return { runId: finishedRunId, results: results };
    },

    /**
     * 执行单条指令
     * role/heading_level 的 detect 在队列外完成，避免嵌套死锁
     */
    _executeOneCommand: async function (cmd) {
      var skillDef = window.FormatSkillRegistry.skills[cmd.skill];
      if (!skillDef) {
        throw new Error('技能 "' + cmd.skill + '" 未注册');
      }
      // custom.execute 不需要预注册的 execute 函数
      if (cmd.skill !== 'custom.execute' && !skillDef.execute) {
        throw new Error('技能 "' + cmd.skill + '" 未实现');
      }

      var self = this;

      // detect 技能直接进队列
      if (cmd.skill.startsWith('detect.')) {
        var detectMeta = await this._runInQueue(cmd.skill, function () {
          return skillDef.execute(cmd.params);
        });
        return {
          queueWaitMs: detectMeta.queueWaitMs,
          queueTaskMs: detectMeta.queueTaskMs
        };
      }

      // custom.execute 自定义代码执行（注入 skills helper 供内部调用内置技能）
      if (cmd.skill === 'custom.execute') {
        var customMeta = await this._runInQueue(cmd.skill, function () {
          var app = wps.WpsApplication();
          var doc = app.ActiveDocument;
          var targetRange = cmd.params.target ? self._resolveTarget(cmd.params) : null;
          var registry = window.FormatSkillRegistry;
          var count = doc.Paragraphs.Count;
          var customBefore = self._captureDocStats(doc);
          var helperTracker = {
            totalCalls: 0,
            byMethod: { font: 0, alignment: 0, spacing: 0, lineSpacing: 0, indent: 0 },
            touchedParagraphs: {}
          };

          function trackHelperCall(method, paragraphIndex) {
            helperTracker.totalCalls += 1;
            helperTracker.byMethod[method] = (helperTracker.byMethod[method] || 0) + 1;
            helperTracker.touchedParagraphs[paragraphIndex] = true;
          }

          function checkIndex(index, method) {
            var n = Number(index);
            if (n !== n || n < 1 || n > count) {
              throw new Error('skills.' + method + ': index 必须是 1 到 ' + count + ' 的整数');
            }
            return n;
          }

          var skillsHelper = {
            font: function (index, params) {
              var i = checkIndex(index, 'font');
              trackHelperCall('font', i);
              var p = params || {};
              registry.skills['text.font.set'].execute({
                zhFont: p.zhFont,
                enFont: p.enFont,
                fontSize: p.fontSize,
                bold: p.bold,
                italic: p.italic,
                underline: p.underline,
                color: p.color
              }, [i]);
            },
            alignment: function (index, align) {
              var i = checkIndex(index, 'alignment');
              trackHelperCall('alignment', i);
              var allowed = ['left', 'center', 'right', 'justify'];
              if (allowed.indexOf(align) === -1) {
                throw new Error('skills.alignment: alignment 必须是 left/center/right/justify');
              }
              registry.skills['paragraph.alignment.set'].execute({ alignment: align }, [i]);
            },
            spacing: function (index, params) {
              var i = checkIndex(index, 'spacing');
              trackHelperCall('spacing', i);
              var p = params || {};
              registry.skills['paragraph.spacing.set'].execute({
                spaceBefore: p.spaceBefore,
                spaceAfter: p.spaceAfter
              }, [i]);
            },
            lineSpacing: function (index, mode, value) {
              var i = checkIndex(index, 'lineSpacing');
              trackHelperCall('lineSpacing', i);
              var allowed = ['multiple', 'exact', 'atLeast'];
              if (allowed.indexOf(mode) === -1) {
                throw new Error('skills.lineSpacing: mode 必须是 multiple/exact/atLeast');
              }
              var v = Number(value);
              if (v !== v) throw new Error('skills.lineSpacing: value 必须是数字');
              registry.skills['paragraph.line_spacing.set'].execute({ mode: mode, value: v }, [i]);
            },
            indent: function (index, params) {
              var i = checkIndex(index, 'indent');
              trackHelperCall('indent', i);
              var p = params || {};
              registry.skills['paragraph.indent.set'].execute({
                firstLineChars: p.firstLineChars,
                firstLinePoints: p.firstLinePoints,
                hanging: p.hanging,
                left: p.left,
                right: p.right
              }, [i]);
            }
          };

          var userEffects = {};
          var customCode = String((cmd.params && cmd.params.code) || '');
          log('  custom.execute begin: codeLength=' + customCode.length + ' desc=' + (cmd.description || cmd.skill));
          var wrappedCode = 'try {\n' + customCode + '\n} catch(e) { throw new Error("自定义代码执行失败: " + e.message); }';
          var fn = new Function('app', 'doc', 'wps', 'range', 'skills', 'effects', wrappedCode);
          fn(app, doc, wps, targetRange, skillsHelper, userEffects);

          var customAfter = self._captureDocStats(doc);
          return {
            customEffects: self._summarizeCustomEffects(customBefore, customAfter, helperTracker, userEffects),
            codeLength: customCode.length
          };
        });

        return {
          queueWaitMs: customMeta.queueWaitMs,
          queueTaskMs: customMeta.queueTaskMs,
          customEffects: customMeta.result ? customMeta.result.customEffects : null,
          codeLength: customMeta.result ? customMeta.result.codeLength : 0
        };
      }

      // 在队列外预解析 role / heading_level 目标
      var preResolvedIndices = null;

      if (cmd.params.target && cmd.params.target.type === 'role') {
        log('  pre-resolve role: ' + cmd.params.target.role);
        preResolvedIndices = await self._resolveRoleTarget(cmd.params.target);
        log('  role resolved: ' + preResolvedIndices.length + ' paragraphs');
      } else if (cmd.params.target && cmd.params.target.type === 'heading_level') {
        log('  pre-resolve heading_level: ' + cmd.params.target.level);
        preResolvedIndices = await self._resolveHeadingTarget(cmd.params.target);
        log('  heading resolved: ' + preResolvedIndices.length + ' paragraphs');
      }

      // 进入队列执行 WPS API（无嵌套）
      var queueMeta = await this._runInQueue(cmd.skill, function () {
        log('  queue executing: ' + cmd.skill);

        if (preResolvedIndices) {
          return skillDef.execute(cmd.params, preResolvedIndices);
        }

        if (cmd.params.target && cmd.params.target.type === 'search') {
          var searchRanges = self._resolveSearchTarget(cmd.params.target);
          log('  search found ' + searchRanges.length + ' matches');
          for (var i = 0; i < searchRanges.length; i++) {
            skillDef.execute(cmd.params, searchRanges[i]);
          }
          return;
        }

        var range = self._resolveTarget(cmd.params);
        return skillDef.execute(cmd.params, range);
      });

      return {
        queueWaitMs: queueMeta.queueWaitMs,
        queueTaskMs: queueMeta.queueTaskMs
      };
    },

    _runInQueue: function (label, task) {
      var enqueueAt = Date.now();
      return wpsActionQueue.add(async function () {
        var startAt = Date.now();
        var queueWaitMs = startAt - enqueueAt;
        log('  queue start: ' + label + ' wait=' + queueWaitMs + 'ms');
        try {
          var result = await task();
          var queueTaskMs = Date.now() - startAt;
          log('  queue done: ' + label + ' task=' + queueTaskMs + 'ms');
          return { queueWaitMs: queueWaitMs, queueTaskMs: queueTaskMs, result: result };
        } catch (err) {
          var failedTaskMs = Date.now() - startAt;
          warn('  queue failed: ' + label + ' task=' + failedTaskMs + 'ms err=' + (err.message || String(err)));
          throw err;
        }
      }, { cooldownMs: COOLDOWN_MS });
    },

    _captureDocStats: function (doc) {
      var stats = {
        paragraphCount: 0,
        tableCount: 0,
        contentLength: 0,
        tableBorderSignature: ''
      };
      if (!doc) return stats;
      try { stats.paragraphCount = doc.Paragraphs.Count; } catch (e) { /* ignore */ }
      try { stats.tableCount = doc.Tables.Count; } catch (e2) { /* ignore */ }
      try {
        var content = doc.Content;
        stats.contentLength = Math.max(0, Number(content.End || 0) - Number(content.Start || 0));
      } catch (e3) { /* ignore */ }

      // 低成本边框签名：用于识别 custom.execute 的表格边框类改动。
      try {
        var borderIds = [-1, -2, -3, -4, -5, -6];
        var tablesToCheck = Math.min(stats.tableCount || 0, 20);
        var parts = [];
        for (var i = 1; i <= tablesToCheck; i++) {
          var table = doc.Tables.Item(i);
          var row = [];
          for (var b = 0; b < borderIds.length; b++) {
            var border = table.Borders.Item(borderIds[b]);
            row.push(String(border.LineStyle) + '/' + String(border.LineWidth) + '/' + String(border.Color));
          }
          parts.push(row.join(','));
        }
        stats.tableBorderSignature = parts.join('|');
      } catch (e4) { /* ignore */ }

      return stats;
    },

    _summarizeCustomEffects: function (before, after, helperTracker, userEffects) {
      before = before || {};
      after = after || {};
      helperTracker = helperTracker || { totalCalls: 0, byMethod: {}, touchedParagraphs: {} };
      userEffects = userEffects || {};

      var touchedParagraphs = Object.keys(helperTracker.touchedParagraphs || {}).length;
      var manualEffectCount = 0;
      var userKeys = Object.keys(userEffects);
      for (var i = 0; i < userKeys.length; i++) {
        var v = userEffects[userKeys[i]];
        if (typeof v === 'number') {
          manualEffectCount += Math.abs(v);
        } else if (v === true) {
          manualEffectCount += 1;
        }
      }

      var paragraphDelta = (after.paragraphCount || 0) - (before.paragraphCount || 0);
      var tableDelta = (after.tableCount || 0) - (before.tableCount || 0);
      var contentLengthDelta = (after.contentLength || 0) - (before.contentLength || 0);
      var tableBorderChanged = before.tableBorderSignature && after.tableBorderSignature &&
        before.tableBorderSignature !== after.tableBorderSignature;
      var estimatedChanged = Math.abs(paragraphDelta) +
        Math.abs(tableDelta) +
        Math.abs(contentLengthDelta) +
        (tableBorderChanged ? 1 : 0) +
        touchedParagraphs +
        manualEffectCount;

      return {
        helperCalls: helperTracker.totalCalls || 0,
        helperCallBreakdown: helperTracker.byMethod || {},
        touchedParagraphs: touchedParagraphs,
        documentDelta: {
          paragraphs: paragraphDelta,
          tables: tableDelta,
          contentLength: contentLengthDelta,
          tableBorderChanged: tableBorderChanged
        },
        userReportedEffects: userEffects,
        manualEffectCount: manualEffectCount,
        estimatedChanged: estimatedChanged
      };
    },

    _estimateEffectCount: function (comparisonEntry, executionMeta) {
      var changes = comparisonEntry && comparisonEntry.changes ? comparisonEntry.changes.length : 0;
      if (changes > 0) return changes;

      var customEffects = executionMeta && executionMeta.customEffects;
      if (!customEffects) return 0;

      var delta = customEffects.documentDelta || {};
      var deltaCount = Math.abs(delta.paragraphs || 0) +
        Math.abs(delta.tables || 0) +
        Math.abs(delta.contentLength || 0) +
        (delta.tableBorderChanged ? 1 : 0);
      return deltaCount +
        Math.abs(customEffects.helperCalls || 0) +
        Math.abs(customEffects.touchedParagraphs || 0) +
        Math.abs(customEffects.manualEffectCount || 0);
    },

    /**
     * 同步捕获状态，不走队列（避免排队延迟）
     */
    _captureStateDirect: function (cmd) {
      try {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        if (!doc) return {};

        var properties = this._getRelevantProperties(cmd.skill);
        if (properties.length === 0) return {};

        // 页面级
        if (cmd.skill.startsWith('page.')) {
          return window.ComparisonLog.captureState(doc.Content, properties);
        }

        // role 目标：取第一个段落
        if (cmd.params.target && cmd.params.target.type === 'role') {
          var cache = window.DocumentStructureCache;
          var indices = cache.getIndices(cmd.params.target.role);
          if (indices && indices.length > 0) {
            var paraRange = doc.Paragraphs.Item(indices[0]).Range;
            return window.ComparisonLog.captureState(paraRange, properties);
          }
        }

        var range = this._resolveTarget(cmd.params);
        if (!range) return {};
        return window.ComparisonLog.captureState(range, properties);
      } catch (e) {
        warn('captureStateDirect error: ' + e.message);
        return {};
      }
    },

    _resolveTarget: function (params) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      var target = params.target;

      if (!target || target.type === 'document') return doc.Content;

      switch (target.type) {
        case 'selection': return app.Selection.Range;
        case 'all_paragraphs': return doc.Content;
        case 'paragraph_index': return doc.Paragraphs.Item(target.index).Range;
        case 'paragraph_range': {
          var start = doc.Paragraphs.Item(target.from).Range.Start;
          var end = doc.Paragraphs.Item(target.to).Range.End;
          return doc.Range(start, end);
        }
        case 'section_index': return doc.Sections.Item(target.index).Range;
        case 'table_index': return doc.Tables.Item(target.index).Range;
        default: return doc.Content;
      }
    },

    _resolveSearchTarget: function (target) {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      var ranges = [];
      var occurrence = target.occurrence || 0;

      var searchRange = doc.Content;
      var find = searchRange.Find;
      find.ClearFormatting();
      find.Text = target.text;
      find.Forward = true;
      find.Wrap = 0;
      find.MatchCase = false;
      find.MatchWholeWord = false;

      var count = 0;
      while (find.Execute()) {
        count++;
        ranges.push(searchRange.Duplicate);
        if (occurrence > 0 && count >= occurrence) break;
        searchRange.Start = searchRange.End;
        searchRange.End = doc.Content.End;
        find = searchRange.Find;
        find.ClearFormatting();
        find.Text = target.text;
        find.Forward = true;
        find.Wrap = 0;
      }

      if (ranges.length === 0) {
        throw new Error('未找到文本: "' + target.text + '"');
      }
      return ranges;
    },

    /**
     * 解析 role 目标（在队列外调用，内部会用 wpsActionQueue）
     */
    _resolveRoleTarget: async function (target) {
      var cache = window.DocumentStructureCache;
      var indices = cache.getIndices(target.role);
      if (indices) {
        log('  role cache hit: ' + target.role + ' count=' + indices.length);
        return indices;
      }
      log('  role cache miss, detecting: ' + target.role);
      await cache.detect(target.role);
      indices = cache.getIndices(target.role);
      if (!indices || indices.length === 0) {
        throw new Error('未检测到角色 "' + target.role + '" 对应的段落');
      }
      log('  role detect done: ' + target.role + ' count=' + indices.length);
      return indices;
    },

    _resolveHeadingTarget: async function (target) {
      return this._resolveRoleTarget({ role: 'heading_' + target.level });
    },

    _getRelevantProperties: function (skill) {
      if (skill === 'custom.execute') {
        return ['Font.Name', 'Font.Size', 'ParagraphFormat.Alignment'];
      }
      if (skill.startsWith('text.font') || skill.startsWith('text.')) {
        return ['Font.Name', 'Font.NameFarEast', 'Font.NameAscii', 'Font.Size', 'Font.Bold', 'Font.Color'];
      }
      if (skill.startsWith('paragraph.')) {
        return ['ParagraphFormat.Alignment', 'ParagraphFormat.LineSpacing',
                'ParagraphFormat.SpaceBefore', 'ParagraphFormat.SpaceAfter',
                'ParagraphFormat.FirstLineIndent'];
      }
      if (skill.startsWith('page.')) {
        return ['PageSetup.TopMargin', 'PageSetup.BottomMargin',
                'PageSetup.LeftMargin', 'PageSetup.RightMargin'];
      }
      return [];
    }
  };
})();
