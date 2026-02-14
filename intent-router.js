/**
 * intent-router.js
 * 意图路由模块：自动识别用户意图并路由到对应执行链路
 *
 * 路由类型：
 * - FORMAT: 格式修改（现有主链路）
 * - CONTENT_EDIT: 基础文本编辑（新增）
 * - GENERAL_QA: 问答回复（不执行文档写操作）
 */

/* global window */

(function () {
  'use strict';

  // 意图类型常量
  var INTENT_FORMAT = 'FORMAT';
  var INTENT_CONTENT_EDIT = 'CONTENT_EDIT';
  var INTENT_GENERAL_QA = 'GENERAL_QA';

  // 置信度阈值配置
  var CONFIDENCE_THRESHOLD = {
    HIGH: 0.8,
    LOW: 0.5
  };

  // 意图关键词规则
  var INTENT_KEYWORDS = {
    [INTENT_FORMAT]: [
      // 格式相关
      '格式', '排版', '样式', '字体', '字号', '加粗', '斜体', '下划线',
      '对齐', '居中', '左对齐', '右对齐', '两端对齐',
      '行距', '行间距', '段落', '缩进', '首行缩进',
      '页边距', '页眉', '页脚', '页码',
      '标题', '正文', '目录', '大纲',
      '表格', '边框', '虚线', '底纹',
      '颜色', '背景', '高亮',
      '宋体', '黑体', '楷体', ' Times ', ' Arial ',
      '三号', '四号', '五号', '小五',
      '1.5', '2倍', '单倍', '固定值',
      '毕业论文', '论文格式', '合同格式', '排版'
    ],
    [INTENT_CONTENT_EDIT]: [
      // 内容编辑相关
      '插入', '添加', '新增',
      '删除', '去掉', '移除', '清除',
      '替换', '修改', '更改', '改',
      '追加', '补充', '在最后',
      '在开头', '在前面', '在后面',
      '在第', '第几段', '第几行',
      '这句话', '这个词', '这段话', '这部分',
      '改成', '改为', '变成', '改成'
    ],
    [INTENT_GENERAL_QA]: [
      // 问答相关
      '是什么', '什么是', '怎么', '如何',
      '为什么', '原因', '原理',
      '介绍一下', '说明', '解释',
      '有没有', '是否可以', '能否',
      '帮助', '使用', '用法', '教程'
    ]
  };

  /**
   * 关键词匹配
   */
  function matchKeywords(text, keywords) {
    var matched = [];
    var lowerText = text.toLowerCase();
    for (var i = 0; i < keywords.length; i++) {
      var keyword = keywords[i].toLowerCase();
      if (lowerText.indexOf(keyword) !== -1) {
        matched.push(keywords[i]);
      }
    }
    return matched;
  }

  /**
   * 基于关键词的意图分类
   */
  function classifyByKeywords(text) {
    var scores = {};

    // 计算每种意图的匹配分数
    for (var intent in INTENT_KEYWORDS) {
      if (INTENT_KEYWORDS.hasOwnProperty(intent)) {
        var matched = matchKeywords(text, INTENT_KEYWORDS[intent]);
        scores[intent] = matched.length;
      }
    }

    // 找出最高分
    var maxScore = 0;
    var topIntent = INTENT_FORMAT; // 默认格式意图
    for (var intent in scores) {
      if (scores.hasOwnProperty(intent)) {
        if (scores[intent] > maxScore) {
          maxScore = scores[intent];
          topIntent = intent;
        }
      }
    }

    // 计算置信度（基于匹配数）
    var totalMatches = 0;
    for (var intent in scores) {
      if (scores.hasOwnProperty(intent)) {
        totalMatches += scores[intent];
      }
    }
    var confidence = totalMatches > 0 ? (maxScore / totalMatches) * 0.8 : 0.5;

    // 如果有匹配，返回结果
    if (maxScore > 0) {
      return {
        intent: topIntent,
        confidence: confidence,
        reason: '关键词匹配: ' + matchKeywords(text, INTENT_KEYWORDS[topIntent]).join(', '),
        allMatches: scores
      };
    }

    // 无匹配，返回默认
    return {
      intent: INTENT_FORMAT,
      confidence: 0.3,
      reason: '无明确关键词匹配，默认走格式链路',
      allMatches: scores
    };
  }

  /**
   * 兜底策略（基于关键词启发式）
   */
  function fallbackHeuristic(text) {
    // 优先检查内容编辑
    var contentKeywords = INTENT_KEYWORDS[INTENT_CONTENT_EDIT];
    var contentMatches = matchKeywords(text, contentKeywords);
    if (contentMatches.length >= 2) {
      return {
        intent: INTENT_CONTENT_EDIT,
        confidence: 0.7,
        reason: '兜底策略: 检测到多个内容编辑关键词'
      };
    }

    // 检查问答
    var qaKeywords = INTENT_KEYWORDS[INTENT_GENERAL_QA];
    var qaMatches = matchKeywords(text, qaKeywords);
    if (qaMatches.length >= 2) {
      return {
        intent: INTENT_GENERAL_QA,
        confidence: 0.7,
        reason: '兜底策略: 检测到多个问答关键词'
      };
    }

    // 默认格式
    return {
      intent: INTENT_FORMAT,
      confidence: 0.5,
      reason: '兜底策略: 默认走格式链路'
    };
  }

  window.IntentRouter = {
    /**
     * 意图分类主函数
     * @param {Object} params - 分类参数
     * @param {string} params.text - 用户输入
     * @param {Array} params.chatContext - 会话上下文（可选）
     * @param {Object} params.docContext - 文档上下文（可选）
     * @returns {Object} { intent, confidence, reason }
     */
    classifyIntent: function (params) {
      var text = params.text || '';
      if (!text.trim()) {
        return {
          intent: INTENT_FORMAT,
          confidence: 0,
          reason: '空输入'
        };
      }

      // 1. 关键词分类
      var keywordResult = classifyByKeywords(text);

      // 2. 如果置信度高，直接返回
      if (keywordResult.confidence >= CONFIDENCE_THRESHOLD.HIGH) {
        console.log('[IntentRouter] High confidence:', keywordResult);
        return keywordResult;
      }

      // 3. 如果置信度中等，记录日志后返回
      if (keywordResult.confidence >= CONFIDENCE_THRESHOLD.LOW) {
        console.log('[IntentRouter] Medium confidence:', keywordResult);
        return keywordResult;
      }

      // 4. 置信度低，使用兜底策略
      console.log('[IntentRouter] Low confidence, using fallback');
      return fallbackHeuristic(text);
    },

    /**
     * 记录路由决策（用于优化）
     */
    logRoutingDecision: function (input, result) {
      console.log('[IntentRouter] Routing decision:', {
        input: input.substring(0, 100),
        intent: result.intent,
        confidence: result.confidence,
        reason: result.reason,
        timestamp: new Date().toISOString()
      });
      // 可以发送到服务端存储用于分析
    },

    /**
     * 获取意图类型常量
     */
    getIntentTypes: function () {
      return {
        FORMAT: INTENT_FORMAT,
        CONTENT_EDIT: INTENT_CONTENT_EDIT,
        GENERAL_QA: INTENT_GENERAL_QA
      };
    }
  };

  console.log('[IntentRouter] Initialized');
})();
