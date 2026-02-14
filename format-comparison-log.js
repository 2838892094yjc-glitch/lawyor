/**
 * format-comparison-log.js
 * 对比日志系统：记录每次操作的前后状态变化
 */

/* global window */

(function () {
  'use strict';

  window.ComparisonLog = {
    entries: [],

    /**
     * 捕获范围的属性状态
     */
    captureState: function (range, properties) {
      var state = {};
      if (!range) return state;

      for (var i = 0; i < properties.length; i++) {
        var prop = properties[i];
        try {
          state[prop] = this._readProperty(range, prop);
        } catch (e) {
          state[prop] = '(无法读取)';
        }
      }
      return state;
    },

    /**
     * 创建对比日志条目
     */
    createEntry: function (cmd, beforeState, afterState, durationMs) {
      var changes = [];

      var props = Object.keys(beforeState);
      for (var i = 0; i < props.length; i++) {
        var prop = props[i];
        var before = String(beforeState[prop] === undefined ? '' : beforeState[prop]);
        var after = String(afterState[prop] === undefined ? '' : afterState[prop]);
        if (before !== after) {
          changes.push({
            property: prop,
            before: before,
            after: after
          });
        }
      }

      return {
        timestamp: new Date().toISOString(),
        commandId: cmd.id,
        skill: cmd.skill,
        description: cmd.description || cmd.skill,
        changes: changes,
        status: 'success',
        duration: durationMs + 'ms'
      };
    },

    /**
     * 添加日志条目
     */
    addEntry: function (entry) {
      this.entries.push(entry);
    },

    /**
     * 生成人类可读的摘要
     */
    formatSummary: function (entry) {
      if (!entry) return '';
      if (entry.status === 'failed') {
        return entry.skill + ': 失败 - ' + (entry.error || '未知错误');
      }
      if (!entry.changes || entry.changes.length === 0) {
        return entry.description + ': 无变化';
      }
      var lines = [];
      for (var i = 0; i < entry.changes.length; i++) {
        var c = entry.changes[i];
        lines.push(c.property + ': ' + c.before + ' -> ' + c.after);
      }
      return lines.join('\n');
    },

    /**
     * 格式化为 Chat UI 展示的树形文本
     */
    formatForChat: function (entry) {
      if (!entry) return '';
      if (entry.status === 'failed') {
        return entry.error || '执行失败';
      }
      if (!entry.changes || entry.changes.length === 0) {
        return '无变化';
      }
      var lines = [];
      for (var i = 0; i < entry.changes.length; i++) {
        var c = entry.changes[i];
        var prefix = (i === entry.changes.length - 1) ? '  ' : '  ';
        lines.push(prefix + c.property + ': ' + c.before + ' -> ' + c.after);
      }
      return lines.join('\n');
    },

    /**
     * 导出完整日志
     */
    exportLog: function () {
      return JSON.stringify(this.entries, null, 2);
    },

    /**
     * 导出结构化日志（带性能指标）
     * 格式：{runId, stepId, queueWaitMs, queueTaskMs, effectCount, warnings}
     */
    exportStructuredMetrics: function () {
      var metrics = [];
      for (var i = 0; i < this.entries.length; i++) {
        var entry = this.entries[i];
        if (!entry) continue;
        var metric = {
          runId: entry.runId || null,
          stepId: entry.stepId || null,
          timestamp: entry.timestamp,
          skill: entry.skill,
          status: entry.status
        };
        // 添加性能指标
        if (entry.performance) {
          metric.queueWaitMs = entry.performance.queueWaitMs || 0;
          metric.queueTaskMs = entry.performance.queueTaskMs || 0;
          metric.totalMs = entry.performance.totalMs || 0;
          metric.captureBeforeMs = entry.performance.captureBeforeMs || 0;
          metric.executeMs = entry.performance.executeMs || 0;
          metric.captureAfterMs = entry.performance.captureAfterMs || 0;
          metric.compareMs = entry.performance.compareMs || 0;
        }
        // 添加 effect 统计
        metric.effectCount = entry.effectCount || 0;
        // 添加 custom effects
        if (entry.customEffects) {
          metric.customEffects = entry.customEffects;
        }
        // 添加 warnings
        if (entry.warnings && entry.warnings.length > 0) {
          metric.warnings = entry.warnings;
        }
        // 添加错误信息
        if (entry.error) {
          metric.error = entry.error;
        }
        metrics.push(metric);
      }
      return metrics;
    },

    /**
     * 写入结构化指标日志（供外部调用）
     */
    logStructuredMetrics: function (runId, metrics) {
      this.entries.push({
        timestamp: new Date().toISOString(),
        runId: runId,
        type: 'metrics',
        metrics: metrics
      });
    },

    /**
     * 获取统计摘要
     */
    getSummary: function () {
      var summary = {
        totalEntries: this.entries.length,
        successCount: 0,
        failedCount: 0,
        zeroEffectCount: 0,
        totalQueueWaitMs: 0,
        totalQueueTaskMs: 0,
        totalEffectCount: 0,
        runs: {}
      };
      for (var i = 0; i < this.entries.length; i++) {
        var entry = this.entries[i];
        if (!entry) continue;
        if (entry.status === 'success') {
          summary.successCount++;
        } else if (entry.status === 'failed') {
          summary.failedCount++;
        }
        if (entry.warnings && entry.warnings.indexOf('SUCCESS_ZERO_EFFECT') !== -1) {
          summary.zeroEffectCount++;
        }
        if (entry.performance) {
          summary.totalQueueWaitMs += entry.performance.queueWaitMs || 0;
          summary.totalQueueTaskMs += entry.performance.queueTaskMs || 0;
        }
        summary.totalEffectCount += entry.effectCount || 0;
        // 按 runId 分组
        if (entry.runId) {
          if (!summary.runs[entry.runId]) {
            summary.runs[entry.runId] = { steps: 0, success: 0, failed: 0 };
          }
          summary.runs[entry.runId].steps++;
          if (entry.status === 'success') {
            summary.runs[entry.runId].success++;
          } else if (entry.status === 'failed') {
            summary.runs[entry.runId].failed++;
          }
        }
      }
      return summary;
    },

    /**
     * 清空日志
     */
    clear: function () {
      this.entries = [];
    },

    /**
     * 读取范围的单个属性
     */
    _readProperty: function (range, prop) {
      if (!range) return undefined;

      switch (prop) {
        case 'Font.Name': return range.Font.Name;
        case 'Font.NameFarEast': return range.Font.NameFarEast;
        case 'Font.NameAscii': return range.Font.NameAscii;
        case 'Font.Size': return range.Font.Size + 'pt';
        case 'Font.Bold': return range.Font.Bold ? '是' : '否';
        case 'Font.Italic': return range.Font.Italic ? '是' : '否';
        case 'Font.Color': return range.Font.Color;
        case 'ParagraphFormat.Alignment': {
          var alignNames = { 0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐' };
          return alignNames[range.ParagraphFormat.Alignment] || String(range.ParagraphFormat.Alignment);
        }
        case 'ParagraphFormat.LineSpacing': return range.ParagraphFormat.LineSpacing + 'pt';
        case 'ParagraphFormat.SpaceBefore': return range.ParagraphFormat.SpaceBefore + 'pt';
        case 'ParagraphFormat.SpaceAfter': return range.ParagraphFormat.SpaceAfter + 'pt';
        case 'ParagraphFormat.FirstLineIndent': return range.ParagraphFormat.FirstLineIndent + 'pt';
        case 'PageSetup.TopMargin': return range.PageSetup.TopMargin + 'pt';
        case 'PageSetup.BottomMargin': return range.PageSetup.BottomMargin + 'pt';
        case 'PageSetup.LeftMargin': return range.PageSetup.LeftMargin + 'pt';
        case 'PageSetup.RightMargin': return range.PageSetup.RightMargin + 'pt';
        default: return undefined;
      }
    }
  };
})();
