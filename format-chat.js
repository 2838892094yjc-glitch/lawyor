/**
 * format-chat.js
 * 聊天 UI 逻辑：管理消息、调用 AI、执行计划
 */

/* global window, wps, wpsActionQueue, document */

(function () {
  'use strict';

  function log(msg) {
    console.log('[FormatChat] ' + msg);
  }
  function warn(msg) {
    console.warn('[FormatChat] ' + msg);
  }

  // 会话上下文配置
  var SESSION_CONFIG = {
    maxTurns: 30              // 保留最近 N 轮对话（现代模型上下文窗口大，30轮足够）
  };

  window.FormatChat = {
    messages: [],
    currentPlan: null,
    _msgIdCounter: 0,
    _lastUserText: '',
    // 会话上下文存储
    _sessionContext: {
      turns: [],              // [{role, content, timestamp, skill?, result?}]
      lastIntent: null,       // 上次意图识别结果
      lastPlanSummary: null    // 上次计划摘要
    },

    /**
     * 追加会话轮次
     * @param {string} role - 'user' | 'assistant' | 'system'
     * @param {string} content - 对话内容
     * @param {object} meta - 额外元数据（skill, result, error等）
     */
    _appendSessionTurn: function (role, content, meta) {
      var turn = {
        role: role,
        content: content,
        timestamp: new Date().toISOString()
      };
      if (meta) {
        if (meta.skill) turn.skill = meta.skill;
        if (meta.result) turn.result = meta.result;
        if (meta.error) turn.error = meta.error;
        if (meta.intent) turn.intent = meta.intent;
      }
      this._sessionContext.turns.push(turn);
      // 超过最大轮次时移除最早的轮次
      if (this._sessionContext.turns.length > SESSION_CONFIG.maxTurns) {
        this._sessionContext.turns.shift();
      }
      log('_appendSessionTurn: ' + role + ', total turns=' + this._sessionContext.turns.length);
    },

    /**
     * 构建发送给 AI 的会话上下文
     * @returns {Array} - messages 数组
     */
    _buildSessionContextForAI: function () {
      var contextMessages = [];
      var turns = this._sessionContext.turns;
      // 保留最近 N 轮
      var recentTurns = turns.slice(-SESSION_CONFIG.maxTurns);
      for (var i = 0; i < recentTurns.length; i++) {
        var turn = recentTurns[i];
        contextMessages.push({
          role: turn.role,
          content: turn.content
        });
      }
      log('_buildSessionContextForAI: ' + contextMessages.length + ' messages from ' + recentTurns.length + ' turns');
      return contextMessages;
    },

    /**
     * 检查是否需要压缩上下文（token 过多时）
     * 清理会话上下文
     */
    _clearSessionContext: function () {
      this._sessionContext.turns = [];
      this._sessionContext.lastIntent = null;
      this._sessionContext.lastPlanSummary = null;
      log('_clearSessionContext: cleared');
    },

    /**
     * 初始化 Chat UI
     */
    init: function () {
      var self = this;
      var input = document.getElementById('format-chat-input');
      var sendBtn = document.getElementById('format-chat-send');

      if (sendBtn) {
        sendBtn.addEventListener('click', function () {
          var text = input ? input.value.trim() : '';
          if (text) {
            if (input) input.value = '';
            self.sendUserMessage(text);
          }
        });
      }

      if (input) {
        input.addEventListener('keydown', function (e) {
          if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            var text = input.value.trim();
            if (text) {
              input.value = '';
              self.sendUserMessage(text);
            }
          }
        });
      }

      // 快捷操作按钮
      var quickBtns = document.querySelectorAll('.fmt-quick-btn');
      for (var i = 0; i < quickBtns.length; i++) {
        quickBtns[i].addEventListener('click', function () {
          var text = this.getAttribute('data-prompt');
          if (text) self.sendUserMessage(text);
        });
      }

      // 清除聊天按钮
      var clearBtn = document.getElementById('format-chat-clear');
      if (clearBtn) {
        clearBtn.addEventListener('click', function () {
          self.clearChat();
        });
      }

      // 导出日志按钮
      var exportBtn = document.getElementById('format-chat-export');
      if (exportBtn) {
        exportBtn.addEventListener('click', function () {
          self.exportLog();
        });
      }
    },

    /**
     * 发送用户消息并处理
     */
    sendUserMessage: async function (text, _retryContext) {
      log('sendUserMessage: "' + text.substring(0, 50) + '"');
      this._lastUserText = text;

      // 记录用户输入到会话上下文（非重试时）
      if (!_retryContext) {
        this._appendSessionTurn('user', text);
        this.addMessage('user', text);
      }
      var loadingId = this.addMessage('loading', _retryContext ? '正在重试（安全修正）...' : '正在分析...');

      try {
        // 1. 获取文档上下文
        log('getDocumentContext...');
        var context = await this.getDocumentContext();
        log('getDocumentContext done: ' + JSON.stringify(context));

        // 2. 意图识别 - 调用后端 /agent/dispatch（带会话历史）
        var routingDecision = null;
        if (!_retryContext) {
          try {
            var sessionContext = this._buildSessionContextForAI();
            var dispatchResponse = await fetch('/agent/dispatch', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                text: text,
                chatContext: sessionContext
              })
            });
            routingDecision = await dispatchResponse.json();
            // 记录意图到会话上下文
            this._sessionContext.lastIntent = routingDecision.intent;
            log('Intent classified: ' + routingDecision.intent + ' (confidence: ' + routingDecision.confidence + ', reason: ' + routingDecision.reason + ')');
          } catch (intentErr) {
            warn('Intent classification failed: ' + intentErr.message + ', defaulting to FORMAT');
            // 出错时默认走 FORMAT
            routingDecision = { intent: 'FORMAT', confidence: 0.3, reason: '服务异常，默认FORMAT' };
          }
        }

        // 3. 调用 AI（重试时追加修正提示前缀）- 不带会话历史
        var aiText = text;
        var retryPrefixes = [];
        if (_retryContext && _retryContext.bannedApis && _retryContext.bannedApis.length > 0) {
          retryPrefixes.push('【安全修正】上次生成的代码被拦截，禁止使用以下 API: ' +
            _retryContext.bannedApis.join(', ') + '。请重新生成不含这些 API 的代码。');
        }
        if (_retryContext && _retryContext.forceBuiltinSkills) {
          var policyHints = _retryContext.policyHints || [];
          retryPrefixes.push('【策略修正】custom.execute 仅可兜底，当前需求必须优先使用内置技能。禁止把字体/段落/页面/表格/样式类操作写成 custom.execute。' +
            (policyHints.length > 0 ? '\n提示: ' + policyHints.join('；') : ''));
        }
        if (retryPrefixes.length > 0) {
          aiText = retryPrefixes.join('\n') + '\n\n' + text;
        }
        log('callAIFormat...');
        // 执行指令不需要会话历史，传 null
        var rawJson = await this.callAIFormat(aiText, context, null);
        log('callAIFormat done, response length=' + rawJson.length);
        log('AI raw response: ' + rawJson.substring(0, 300));

        // 记录 AI 响应到会话上下文
        this._appendSessionTurn('assistant', rawJson.substring(0, 500), { skill: 'format' });

        // 3. Compile
        log('FormatCompiler.compile...');
        var compiled = window.FormatCompiler.compile(rawJson);
        log('compile done, commands=' + compiled.commands.length + ' errors=' + compiled.errors.length);
        if (compiled.errors.length > 0) {
          this.removeMessage(loadingId);
          this.addMessage('error', '编译失败: ' + compiled.errors.map(function (e) { return e.message; }).join('; '));
          return;
        }

        // 如果是"无法执行"
        if (compiled.commands.length === 0 && compiled.plan) {
          this.removeMessage(loadingId);
          this.addMessage('error', compiled.plan);
          return;
        }

        // 4. Validate
        log('FormatValidator.validate...');
        var validation = window.FormatValidator.validate(compiled);
        log('validate done, valid=' + validation.valid + ' errors=' + validation.errors.length + ' warnings=' + validation.warnings.length);
        if (!validation.valid) {
          var retryContext = _retryContext || {};
          var retryPatch = {};

          // 检查是否有 UNSAFE_CODE 错误，自动重试要求模型避开危险 API
          var unsafeErrors = validation.errors.filter(function (e) { return e.code === 'UNSAFE_CODE'; });
          if (unsafeErrors.length > 0 && !(retryContext.bannedApis && retryContext.bannedApis.length > 0)) {
            retryPatch.bannedApis = unsafeErrors.map(function (e) { return e.bannedApi || 'unknown'; });
          }

          // 检查是否触发 custom.execute 兜底策略，自动重试要求模型改用内置技能
          var fallbackErrors = validation.errors.filter(function (e) { return e.code === 'CUSTOM_NOT_FALLBACK'; });
          if (fallbackErrors.length > 0 && !retryContext.forceBuiltinSkills) {
            retryPatch.forceBuiltinSkills = true;
            retryPatch.policyHints = fallbackErrors
              .map(function (e) { return e.policyHint || ''; })
              .filter(function (h, idx, arr) { return h && arr.indexOf(h) === idx; });
          }

          if (Object.keys(retryPatch).length > 0) {
            log('validation intercepted, auto-retrying with patch=' + JSON.stringify(retryPatch));
            this.removeMessage(loadingId);
            if (retryPatch.bannedApis) {
              this.addMessage('warning', '检测到禁止的 API (' + retryPatch.bannedApis.join(', ') + ')，自动重试中...');
            }
            if (retryPatch.forceBuiltinSkills) {
              this.addMessage('warning', '检测到 custom.execute 非兜底使用，自动重试并强制优先内置技能...');
            }
            return this.sendUserMessage(text, Object.assign({}, retryContext, retryPatch));
          }

          this.removeMessage(loadingId);
          this.addMessage('error', '验证失败:\n' + validation.errors.map(function (e) { return e.message; }).join('\n'));
          return;
        }

        // 显示警告
        if (validation.warnings.length > 0) {
          for (var w = 0; w < validation.warnings.length; w++) {
            this.addMessage('warning', validation.warnings[w].message);
          }
        }

        // 5. 显示计划，等待确认
        this.removeMessage(loadingId);
        this.currentPlan = compiled;
        log('showing plan to user, commands=' + compiled.commands.length);
        this.addMessage('plan', compiled);

      } catch (err) {
        warn('sendUserMessage error: ' + (err.message || String(err)));
        this.removeMessage(loadingId);
        this.addMessage('error', err.message || String(err));
      }
    },

    /**
     * 执行当前计划
     */
    executePlan: async function (saveFirst) {
      if (!this.currentPlan) return;
      var compiled = this.currentPlan;
      this.currentPlan = null;
      log('executePlan: ' + compiled.commands.length + ' commands, saveFirst=' + !!saveFirst);

      // 隐藏执行/取消按钮
      var planActions = document.querySelectorAll('.format-plan-actions');
      for (var a = 0; a < planActions.length; a++) {
        planActions[a].style.display = 'none';
      }

      // 用户选择「保存并执行」时才保存
      if (saveFirst) {
        try {
          var _doc = wps.WpsApplication().ActiveDocument;
          if (_doc && _doc.Path) {
            _doc.Save();
            log('document saved (user requested)');
          }
        } catch (e) {
          warn('pre-save failed: ' + e.message);
        }
      }

      // 清除 comparison log
      window.ComparisonLog.clear();

      // 保存格式快照（用于回滚）
      if (window.FormatSnapshot) {
        try {
          log('beginUndoGroup...');
          window.FormatSnapshot.beginUndoGroup('AI 排版: ' + (compiled.plan || '').substring(0, 30));
          log('beginUndoGroup done');
        } catch (e) {
          warn('snapshot begin failed: ' + e.message);
        }
      }

      var self = this;
      var stepMsgIds = [];

      log('FormatExecutionEngine.run() starting...');
      await window.FormatExecutionEngine.run(compiled, {
        onStepStart: function (i, cmd) {
          log('UI onStepStart: ' + (i + 1) + ' ' + cmd.skill);
          var msgId = self.addMessage('step-running', {
            index: i,
            description: cmd.description || cmd.skill
          });
          stepMsgIds[i] = msgId;
        },
        onStepComplete: function (i, cmd, comparison) {
          log('UI onStepComplete: ' + (i + 1) + ' ' + cmd.skill + ' changes=' + (comparison ? comparison.changes.length : 0));
          self.updateStepMessage(stepMsgIds[i], 'done', comparison);
          if (comparison && comparison.warnings && comparison.warnings.indexOf('SUCCESS_ZERO_EFFECT') !== -1) {
            // 增强 zero-effect 告警，提供更详细的建议
            var skillName = cmd.skill || 'unknown';
            var suggestions = [];
            if (skillName.includes('detect')) {
              suggestions.push('可能是文档中没有符合条件的元素');
            } else if (skillName.includes('table')) {
              suggestions.push('表格格式可能已经是目标样式');
            } else if (skillName.includes('paragraph') || skillName.includes('text')) {
              suggestions.push('段落文本可能已经是目标格式');
            } else if (skillName.includes('custom')) {
              suggestions.push('自定义代码可能没有实际修改文档');
            } else {
              suggestions.push('请检查目标元素是否正确选择');
            }
            // 添加性能信息帮助诊断
            var perfInfo = '';
            if (comparison.performance) {
              perfInfo = ' (执行' + (comparison.performance.totalMs || 0) + 'ms)';
            }
            self.addMessage('warning', '⚠️ 步骤 ' + (i + 1) + ' 执行成功但未检测到有效变更' + perfInfo + '\n可能原因：' + suggestions.join('；') + '。\n建议：请尝试更明确地描述格式要求，或检查目标元素。');
          }
        },
        onStepError: function (i, cmd, err) {
          warn('UI onStepError: ' + (i + 1) + ' ' + cmd.skill + ' err=' + (err.message || String(err)));
          self.updateStepMessage(stepMsgIds[i], 'error', null, err.message || String(err));
        },
        onComplete: function (results) {
          var success = results.filter(function (r) { return r.success; }).length;
          var failed = results.filter(function (r) { return !r.success; }).length;
          log('UI onComplete: success=' + success + ' failed=' + failed);
          self.addMessage('summary', { success: success, failed: failed, total: results.length });
        }
      });

      // 结束 Undo 组
      if (window.FormatSnapshot) {
        try {
          log('endUndoGroup...');
          await window.FormatSnapshot.endUndoGroup();
          log('endUndoGroup done');
        } catch (e) {
          warn('snapshot end failed: ' + e.message);
        }
      }
      log('executePlan finished');
    },

    /**
     * 取消当前计划
     */
    cancelPlan: function () {
      this.currentPlan = null;
      var planActions = document.querySelectorAll('.format-plan-actions');
      for (var a = 0; a < planActions.length; a++) {
        planActions[a].style.display = 'none';
      }
      this.addMessage('error', '已取消执行');
    },

    /**
     * 重试上一次请求
     */
    retry: function () {
      if (this._lastUserText) {
        this.sendUserMessage(this._lastUserText);
      }
    },

    /**
     * 清除聊天记录
     */
    clearChat: function () {
      var container = document.getElementById('format-chat-messages');
      if (container) {
        container.innerHTML = '<div class="format-msg">' +
          '<div class="fmt-msg-bubble fmt-msg-system-bubble">' +
          '描述你想要的排版格式，例如：<br>' +
          '"宋体小四，1.5倍行距"<br>' +
          '"标题居中加粗黑体三号"<br>' +
          '"按毕业论文格式排版"' +
          '</div></div>';
      }
      this.messages = [];
      this.currentPlan = null;
      this._msgIdCounter = 0;
      // 清理会话上下文
      this._clearSessionContext();
      window.ComparisonLog.clear();
      window.DocumentStructureCache.invalidate();
      if (window.FormatSnapshot) window.FormatSnapshot.clear();
    },

    /**
     * 一键回滚上一次排版操作
     */
    rollback: async function () {
      if (!window.FormatSnapshot) {
        this.addMessage('error', '回滚功能不可用');
        return;
      }
      var loadingId = this.addMessage('loading', '正在回滚...');
      try {
        var result = await window.FormatSnapshot.rollback();
        this.removeMessage(loadingId);
        this.addMessage('summary', { success: 1, failed: 0, total: 1 });
        // 清除结构缓存，因为文档已变化
        window.DocumentStructureCache.invalidate();
      } catch (err) {
        this.removeMessage(loadingId);
        this.addMessage('error', '回滚失败: ' + (err.message || String(err)));
      }
    },

    /**
     * 导出对比日志
     */
    exportLog: function () {
      var log = window.ComparisonLog.exportLog();
      if (!log || log === '[]') {
        this.addMessage('error', '暂无执行日志');
        return;
      }
      // 创建 Blob 下载
      try {
        var blob = new Blob([log], { type: 'application/json' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'format-log-' + new Date().toISOString().slice(0, 10) + '.json';
        a.click();
        URL.revokeObjectURL(url);
        this.addMessage('summary', { success: 1, failed: 0, total: 1 });
      } catch (e) {
        // 如果无法下载，显示在聊天中
        this.addMessage('error', '导出失败: ' + e.message);
      }
    },

    /**
     * 获取文档上下文
     */
    getDocumentContext: function () {
      try {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        if (!doc) return { hasDocument: false };
        var selText = '';
        try { selText = app.Selection.Text ? app.Selection.Text.substring(0, 200) : ''; } catch (e) { /* ignore */ }
        return {
          hasDocument: true,
          paragraphCount: doc.Paragraphs.Count,
          tableCount: doc.Tables.Count,
          sectionCount: doc.Sections.Count,
          hasSelection: app.Selection.Start !== app.Selection.End,
          selectionText: selText
        };
      } catch (e) {
        warn('getDocumentContext error: ' + e.message);
        return { hasDocument: false };
      }
    },

    /**
     * 调用 AI 排版接口（含超时和自动重试）
     * @param {string} text - 用户输入
     * @param {object} context - 文档上下文
     * @param {Array} chatContext - 会话历史上下文（可选）
     */
    callAIFormat: async function (text, context, chatContext) {
      var maxRetries = 2;
      var timeoutMs = 90000; // 90 秒超时（kimi-k2.5 较慢）
      var lastError = null;

      for (var attempt = 0; attempt <= maxRetries; attempt++) {
        var controller = new AbortController();
        var timeoutId = setTimeout(function () { controller.abort(); }, timeoutMs);

        // 构建请求 payload
        var payload = { text: text, context: context };
        // 添加会话上下文（如果有）
        if (chatContext && chatContext.length > 0) {
          payload.chatContext = chatContext;
        }

        try {
          var formatUrl = (window.PEVCRuntimeConfig && window.PEVCRuntimeConfig.buildApiUrl)
            ? window.PEVCRuntimeConfig.buildApiUrl('/format')
            : 'http://127.0.0.1:8765/format';
          var response = await fetch(formatUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
            signal: controller.signal
          });
          clearTimeout(timeoutId);

          if (!response.ok) {
            var errText = '';
            try { errText = await response.text(); } catch (e) { /* ignore */ }
            // 5xx 错误可重试
            if (response.status >= 500 && attempt < maxRetries) {
              lastError = new Error('AI 服务错误 (' + response.status + ')');
              continue;
            }
            throw new Error('AI 服务错误 (' + response.status + '): ' + errText);
          }
          return await response.text();
        } catch (err) {
          clearTimeout(timeoutId);
          if (err.name === 'AbortError') {
            lastError = new Error('AI 服务超时 (' + (timeoutMs / 1000) + '秒)，请稍后重试');
            if (attempt < maxRetries) continue;
            throw lastError;
          }
          // 网络错误可重试
          if (attempt < maxRetries && (err.message || '').indexOf('fetch') !== -1) {
            lastError = err;
            continue;
          }
          throw err;
        }
      }
      throw lastError || new Error('AI 服务不可用');
    },

    // ── UI 方法 ──

    addMessage: function (type, content) {
      var container = document.getElementById('format-chat-messages');
      if (!container) return null;

      var msgId = 'fmt-msg-' + (++this._msgIdCounter);
      var div = document.createElement('div');
      div.id = msgId;
      div.className = 'format-msg format-msg-' + type;

      switch (type) {
        case 'user':
          div.innerHTML = '<div class="fmt-msg-bubble fmt-msg-user-bubble">' + this._escapeHtml(content) + '</div>';
          break;

        case 'loading':
          div.innerHTML = '<div class="fmt-msg-bubble fmt-msg-system-bubble">' +
            '<span class="fmt-spinner"></span> ' + this._escapeHtml(content) + '</div>';
          break;

        case 'error':
          div.innerHTML = '<div class="fmt-msg-bubble fmt-msg-error-bubble">' + this._escapeHtml(content) + '</div>';
          break;

        case 'warning':
          div.innerHTML = '<div class="fmt-msg-bubble fmt-msg-warning-bubble">' + this._escapeHtml(content) + '</div>';
          break;

        case 'plan':
          var planHtml = '<div class="fmt-msg-bubble fmt-msg-plan-bubble">';
          planHtml += '<div class="fmt-plan-title">' + this._escapeHtml(content.plan) + '</div>';
          planHtml += '<ol class="fmt-plan-steps">';
          for (var i = 0; i < content.commands.length; i++) {
            planHtml += '<li>' + this._escapeHtml(content.commands[i].description || content.commands[i].skill) + '</li>';
          }
          planHtml += '</ol>';
          planHtml += '<div class="format-plan-actions">';
          planHtml += '<button class="fmt-btn fmt-btn-execute" onclick="window.FormatChat.executePlan(false)">&#9654; 执行</button>';
          planHtml += '<button class="fmt-btn fmt-btn-execute-save" onclick="window.FormatChat.executePlan(true)">&#128190; 保存并执行</button>';
          planHtml += '<button class="fmt-btn fmt-btn-cancel" onclick="window.FormatChat.cancelPlan()">&#10005; 取消</button>';
          planHtml += '</div></div>';
          div.innerHTML = planHtml;
          break;

        case 'step-running':
          div.innerHTML = '<div class="fmt-msg-bubble fmt-msg-step-bubble fmt-step-running">' +
            '<span class="fmt-spinner"></span> ' +
            this._escapeHtml((content.index + 1) + '. ' + content.description) + '</div>';
          break;

        case 'summary':
          var statusText = content.success + ' 成功';
          if (content.failed > 0) statusText += ' / ' + content.failed + ' 失败';
          var statusClass = content.failed > 0 ? 'fmt-msg-warning-bubble' : 'fmt-msg-success-bubble';
          div.innerHTML = '<div class="fmt-msg-bubble ' + statusClass + '">执行完成: ' + statusText + '</div>';
          break;

        default:
          div.innerHTML = '<div class="fmt-msg-bubble">' + this._escapeHtml(String(content)) + '</div>';
      }

      container.appendChild(div);
      container.scrollTop = container.scrollHeight;

      this.messages.push({ id: msgId, type: type, content: content });
      return msgId;
    },

    removeMessage: function (msgId) {
      if (!msgId) return;
      var el = document.getElementById(msgId);
      if (el) el.remove();
    },

    updateStepMessage: function (msgId, status, comparison, errorMsg) {
      if (!msgId) return;
      var el = document.getElementById(msgId);
      if (!el) return;

      var bubble = el.querySelector('.fmt-msg-step-bubble');
      if (!bubble) return;

      if (status === 'done') {
        bubble.className = 'fmt-msg-bubble fmt-msg-step-bubble fmt-step-done';
        var icon = '<span class="fmt-step-icon-done">&#10003;</span> ';
        var originalText = bubble.textContent.replace(/^\s*/, '');
        var changeHtml = '';
        if (comparison && comparison.changes && comparison.changes.length > 0) {
          changeHtml = '<div class="fmt-step-changes">';
          for (var i = 0; i < comparison.changes.length; i++) {
            var c = comparison.changes[i];
            changeHtml += '<div class="fmt-step-change">' +
              this._escapeHtml(c.property) + ': ' +
              this._escapeHtml(c.before) + ' -> ' +
              this._escapeHtml(c.after) + '</div>';
          }
          changeHtml += '</div>';
        }
        bubble.innerHTML = icon + this._escapeHtml(originalText) +
          (comparison ? ' (' + comparison.duration + ')' : '') + changeHtml;
      } else if (status === 'error') {
        bubble.className = 'fmt-msg-bubble fmt-msg-step-bubble fmt-step-error';
        var errIcon = '<span class="fmt-step-icon-error">&#10007;</span> ';
        var origText = bubble.textContent.replace(/^\s*/, '');
        bubble.innerHTML = errIcon + this._escapeHtml(origText) +
          '<div class="fmt-step-error-msg">' + this._escapeHtml(errorMsg || '未知错误') + '</div>';
      }
    },

    _escapeHtml: function (str) {
      if (!str) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/\n/g, '<br>');
    }
  };
})();
