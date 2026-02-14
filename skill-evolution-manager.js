/**
 * skill-evolution-manager.js
 * Skill 自进化管理器
 * 从 custom.execute 成功样本沉淀为可复用内置技能候选
 */

/* global window */

(function () {
  'use strict';

  // 配置
  var CONFIG = {
    MAX_CANDIDATE_AGE_DAYS: 90,          // 候选有效期90天
    MIN_VALIDATION_SAMPLES: 10,          // 最少验证样本数
    VALIDATION_SUCCESS_RATE: 0.8,        // 验证通过率阈值
    MAX_CUSTOM_PER_REQUEST: 3,           // 单次请求 custom 上限
    DEPRECATION_THRESHOLD: 3              // 连续验证失败次数阈值
  };

  // 候选状态
  var CANDIDATE_STATUS = {
    DRAFT: 'draft',           // 草稿
    VALIDATING: 'validating', // 验证中
    APPROVED: 'approved',     // 待发布
    PUBLISHED: 'published',   // 已发布
    DEPRECATED: 'deprecated'  // 已废弃
  };

  // 候选存储（内存缓存）
  var candidates = {};
  var candidateIndex = [];

  // 生成候选ID
  function generateCandidateId() {
    var now = new Date();
    var dateStr = now.getFullYear() +
      String(now.getMonth() + 1).padStart(2, '0') +
      String(now.getDate()).padStart(2, '0');
    var random = Math.random().toString(36).substring(2, 8);
    return 'candidate-' + dateStr + '-' + random;
  }

  // 从 custom.execute 结果收集候选
  function collectCandidateFromExecution(stepResult) {
    // stepResult 包含: skill, params, executionMeta, comparison
    if (!stepResult || stepResult.skill !== 'custom.execute') {
      return null;
    }

    var executionMeta = stepResult.executionMeta || {};
    var customCode = executionMeta.customCode;
    var customResult = executionMeta.result;

    // 检查是否执行成功
    if (!customResult || !customResult.success) {
      console.log('[SkillEvolution] Skip: custom.execute not successful');
      return null;
    }

    // 检查 effectCount
    var effectCount = executionMeta.effectCount || 0;
    if (effectCount === 0) {
      console.log('[SkillEvolution] Skip: zero effect');
      return null;
    }

    // 生成候选
    var candidate = {
      id: generateCandidateId(),
      createdAt: new Date().toISOString(),
      status: CANDIDATE_STATUS.DRAFT,
      source: {
        customCode: customCode,
        effectCount: effectCount,
        params: stepResult.params
      },
      spec: null,
      samples: [],
      validationReports: [],
      validationFailures: 0,
      publishedAt: null
    };

    // 保存到存储
    candidates[candidate.id] = candidate;
    candidateIndex.push(candidate.id);

    console.log('[SkillEvolution] Candidate collected: ' + candidate.id + ', effects=' + effectCount);
    return candidate;
  }

  // 从 custom 代码提取可模版化的逻辑
  function extractTemplateFromCustom(code, effectMeta, promptText) {
    // 这是一个简化版本，实际需要更复杂的代码分析
    // 目前只是提取关键信息供人工审核

    var template = {
      description: '',
      skillName: '',
      params: [],
      codeAnalysis: {
        operations: [],
        targetTypes: []
      }
    };

    // 简单分析：提取常见的操作模式
    var operations = [];
    if (code.indexOf('.Font.') !== -1) {
      operations.push('font');
    }
    if (code.indexOf('.ParagraphFormat.') !== -1) {
      operations.push('paragraph');
    }
    if (code.indexOf('.Borders') !== -1) {
      operations.push('border');
    }
    if (code.indexOf('.Shading') !== -1) {
      operations.push('shading');
    }

    template.codeAnalysis.operations = operations;

    // 生成建议的 skill 名称
    if (operations.length > 0) {
      template.skillName = 'custom.' + operations.join('_');
    } else {
      template.skillName = 'custom.general';
    }

    // 从 prompt 提取描述
    if (promptText) {
      template.description = promptText.substring(0, 200);
    }

    return template;
  }

  // 生成候选技能规范
  function generateCandidateSkillSpec(templateInput) {
    var spec = {
      id: templateInput.candidateId,
      skillName: templateInput.skillName || 'custom.general',
      description: templateInput.description || '',
      layer: 6,
      params: templateInput.params || [],
      execute: '// TODO: 实现具体逻辑\nthrow new Error("Not implemented");',
      source: templateInput.source
    };
    return spec;
  }

  // 验证候选技能
  function validateCandidateSkill(candidateId) {
    var candidate = candidates[candidateId];
    if (!candidate) {
      throw new Error('Candidate not found: ' + candidateId);
    }

    // 更新状态
    candidate.status = CANDIDATE_STATUS.VALIDATING;

    // 模拟验证（实际需要执行测试）
    // 这里应该：运行验证样本，检查成功率

    var validationReport = {
      validatedAt: new Date().toISOString(),
      successRate: 0, // 实际计算
      samplesTested: 0,
      passed: false,
      errors: []
    };

    // TODO: 实际验证逻辑
    // 1. 加载验证样本
    // 2. 执行技能
    // 3. 比较结果
    // 4. 生成报告

    candidate.validationReports.push(validationReport);

    // 检查是否通过
    if (validationReport.successRate >= CONFIG.VALIDATION_SUCCESS_RATE) {
      candidate.status = CANDIDATE_STATUS.APPROVED;
      return { passed: true, report: validationReport };
    } else {
      candidate.validationFailures++;
      if (candidate.validationFailures >= CONFIG.DEPRECATION_THRESHOLD) {
        candidate.status = CANDIDATE_STATUS.DEPRECATED;
      } else {
        candidate.status = CANDIDATE_STATUS.DRAFT;
      }
      return { passed: false, report: validationReport };
    }
  }

  // 发布候选技能
  function publishCandidateSkill(candidateId, reviewer) {
    var candidate = candidates[candidateId];
    if (!candidate) {
      throw new Error('Candidate not found: ' + candidateId);
    }

    if (candidate.status !== CANDIDATE_STATUS.APPROVED) {
      throw new Error('Candidate not approved for publish: ' + candidate.status);
    }

    // 更新状态
    candidate.status = CANDIDATE_STATUS.PUBLISHED;
    candidate.publishedAt = new Date().toISOString();
    candidate.publishedBy = reviewer;

    console.log('[SkillEvolution] Candidate published: ' + candidateId);

    return {
      success: true,
      skillId: candidate.spec.skillName,
      message: '技能已发布到 skills/ 目录'
    };
  }

  // 废弃候选
  function deprecateCandidate(candidateId, reason) {
    var candidate = candidates[candidateId];
    if (!candidate) {
      throw new Error('Candidate not found: ' + candidateId);
    }

    candidate.status = CANDIDATE_STATUS.DEPRECATED;
    candidate.deprecatedAt = new Date().toISOString();
    candidate.deprecationReason = reason;

    console.log('[SkillEvolution] Candidate deprecated: ' + candidateId + ', reason: ' + reason);
  }

  // 获取候选列表
  function getCandidates(filter) {
    var result = [];
    for (var i = 0; i < candidateIndex.length; i++) {
      var id = candidateIndex[i];
      var candidate = candidates[id];
      if (!candidate) continue;

      if (filter) {
        if (filter.status && candidate.status !== filter.status) continue;
      }
      result.push(candidate);
    }
    return result;
  }

  // 获取候选详情
  function getCandidate(candidateId) {
    return candidates[candidateId] || null;
  }

  // 删除候选
  function deleteCandidate(candidateId) {
    if (candidates[candidateId]) {
      delete candidates[candidateId];
      var idx = candidateIndex.indexOf(candidateId);
      if (idx !== -1) {
        candidateIndex.splice(idx, 1);
      }
      console.log('[SkillEvolution] Candidate deleted: ' + candidateId);
      return true;
    }
    return false;
  }

  // 导出到全局
  window.SkillEvolutionManager = {
    // 配置
    CONFIG: CONFIG,
    STATUS: CANDIDATE_STATUS,

    // 核心方法
    collectCandidateFromExecution: collectCandidateFromExecution,
    extractTemplateFromCustom: extractTemplateFromCustom,
    generateCandidateSkillSpec: generateCandidateSkillSpec,
    validateCandidateSkill: validateCandidateSkill,
    publishCandidateSkill: publishCandidateSkill,
    deprecateCandidate: deprecateCandidate,

    // 查询方法
    getCandidates: getCandidates,
    getCandidate: getCandidate,
    deleteCandidate: deleteCandidate,

    // 统计
    getStats: function () {
      var stats = {
        total: candidateIndex.length,
        byStatus: {}
      };
      for (var i = 0; i < candidateIndex.length; i++) {
        var c = candidates[candidateIndex[i]];
        if (c) {
          stats.byStatus[c.status] = (stats.byStatus[c.status] || 0) + 1;
        }
      }
      return stats;
    }
  };

  console.log('[SkillEvolutionManager] Initialized');
})();
