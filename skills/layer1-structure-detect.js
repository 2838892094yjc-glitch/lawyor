/**
 * layer1-structure-detect.js
 * 第一层：文档结构识别
 * 识别文档各区域的语义角色，不做格式修改。
 */

/* global window, wps, wpsActionQueue */

(function () {
  'use strict';

  function log(msg) {
    console.log('[StructDetect] ' + msg);
  }

  // ── 结构识别缓存 ──

  var DocumentStructureCache = {
    _cache: {},
    _docKey: null,

    _getCurrentDocKey: function () {
      try {
        var doc = wps.WpsApplication().ActiveDocument;
        return doc ? (doc.FullName || doc.Name || 'unknown') : null;
      } catch (e) {
        return null;
      }
    },

    _checkDocChanged: function () {
      var key = this._getCurrentDocKey();
      if (key !== this._docKey) {
        log('doc changed, invalidating cache. old=' + this._docKey + ' new=' + key);
        this._cache = {};
        this._docKey = key;
      }
    },

    async detect(role) {
      log('detect() called, role=' + role);
      this._checkDocChanged();
      if (this._cache[role]) {
        log('detect() cache hit for ' + role);
        return this._cache[role];
      }

      // 首次 miss → 统一扫描，一次填充所有角色
      if (Object.keys(this._cache).length === 0) {
        log('detect() first miss, running _detectAllInOnePass...');
        var allResults = await _detectAllInOnePass();
        if (allResults) {
          for (var r in allResults) {
            if (allResults.hasOwnProperty(r)) this._cache[r] = allResults[r];
          }
        }
        if (this._cache[role]) return this._cache[role];
      }

      // Fallback：个别检测
      log('detect() cache miss, running _runDetect for ' + role + '...');
      var result = await _runDetect(role);
      if (result) {
        this._cache[role] = result;
        log('detect() done, role=' + role + ' indices=' + (result.paragraphIndices || []).length);
      } else {
        log('detect() returned null for ' + role);
      }
      return result;
    },

    getIndices: function (role) {
      this._checkDocChanged();
      var cached = this._cache[role];
      if (!cached) return null;
      return cached.paragraphIndices || null;
    },

    invalidate: function () {
      log('cache invalidated');
      this._cache = {};
    }
  };

  // ── 统一扫描：一次遍历识别所有角色 ──

  async function _detectAllInOnePass() {
    log('_detectAllInOnePass() starting...');
    return wpsActionQueue.add(async function () {
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      if (!doc) return null;

      var count = doc.Paragraphs.Count;
      log('_detectAllInOnePass() paragraphs=' + count);

      // 收集容器
      var titleIndices = [];
      var h1Indices = [], h2Indices = [], h3Indices = [], h4Indices = [];
      var abstractIndices = [], keywordsIndices = [], bodyIndices = [];
      var figCaptionIndices = [], tblCaptionIndices = [];
      var refIndices = [], ackIndices = [], tocIndices = [];

      // 标题模式
      var chapterPatterns = {
        1: /^第[一二三四五六七八九十百]+[章节]/,
        2: /^第[一二三四五六七八九十百]+[条款]|^\d+\.\d+\s/,
        3: /^\d+\.\d+\.\d+\s/,
        4: /^\d+\.\d+\.\d+\.\d+\s/
      };
      var headingStylePatterns = {
        1: { local: /标题\s*1/i, en: /Heading\s*1/i },
        2: { local: /标题\s*2/i, en: /Heading\s*2/i },
        3: { local: /标题\s*3/i, en: /Heading\s*3/i },
        4: { local: /标题\s*4/i, en: /Heading\s*4/i }
      };
      var sizeThresholds = { 1: 16, 2: 14, 3: 12, 4: 11 };

      // 状态标志
      var titleFound = false;
      var abstractStart = -1, inAbstract = false;
      var inReferences = false, inAcknowledgment = false;

      for (var i = 1; i <= count; i++) {
        if (i % 40 === 0) await _yield();

        var para = doc.Paragraphs.Item(i);
        var text = '';
        try { text = (para.Range.Text || '').replace(/[\r\n]/g, '').trim(); } catch (e) { /* ignore */ }
        var styleName = '';
        try { styleName = para.Style.NameLocal || para.Style.Name || ''; } catch (e) { /* ignore */ }

        var isExcluded = false;

        // --- 标题检测（仅前10段） ---
        if (!titleFound && i <= 10 && text) {
          if (/标题|Title/i.test(styleName) && !/标题\s*[2-9]|Heading\s*[2-9]/i.test(styleName)) {
            titleIndices.push(i);
            titleFound = true;
            isExcluded = true;
          } else {
            try {
              var tf = para.Range.Font;
              var tpf = para.Format;
              if (tf.Size >= 18 && tf.Bold && tpf.Alignment === 1) {
                titleIndices.push(i);
                titleFound = true;
                isExcluded = true;
              }
            } catch (e) { /* ignore */ }
          }
        }

        // --- TOC 检测（前50段） ---
        if (i <= 50) {
          if (/目录|TOC/i.test(styleName)) {
            tocIndices.push(i);
            isExcluded = true;
          }
          if (/^目\s*录$/.test(text) && !tocIndices.includes(i)) {
            tocIndices.push(i);
            isExcluded = true;
          }
        }

        // --- 摘要检测（前30段） ---
        if (i <= 30) {
          if (!inAbstract && /^摘\s*要|^Abstract/i.test(text)) {
            inAbstract = true;
            if (text.length < 10) {
              abstractStart = i + 1;
            } else {
              abstractIndices.push(i);
              abstractStart = i + 1;
            }
            isExcluded = true;
          } else if (inAbstract && i >= abstractStart) {
            if (!text || /^关键词|^Keywords|^Key\s*words/i.test(text) || /标题|Heading/i.test(styleName)) {
              inAbstract = false;
            } else {
              abstractIndices.push(i);
              isExcluded = true;
            }
          }
        }

        // --- 关键词检测（前30段） ---
        if (i <= 30 && /^关键词|^Keywords|^Key\s*words/i.test(text)) {
          keywordsIndices.push(i);
          isExcluded = true;
        }

        // --- 参考文献检测 ---
        if (/^参考文献|^References/i.test(text)) {
          inReferences = true;
          inAcknowledgment = false;
          refIndices.push(i);
          isExcluded = true;
        } else if (inReferences) {
          if (/标题\s*1|Heading\s*1/i.test(styleName) || /^致\s*谢|^附\s*录|^Acknowledgment|^Appendix/i.test(text)) {
            inReferences = false;
          } else if (text) {
            refIndices.push(i);
            isExcluded = true;
          }
        }

        // --- 致谢检测 ---
        if (/^致\s*谢|^Acknowledgment/i.test(text)) {
          inAcknowledgment = true;
          inReferences = false;
          ackIndices.push(i);
          isExcluded = true;
        } else if (inAcknowledgment) {
          if (/标题\s*1|Heading\s*1/i.test(styleName) || /^附\s*录|^Appendix/i.test(text)) {
            inAcknowledgment = false;
          } else if (text) {
            ackIndices.push(i);
            isExcluded = true;
          }
        }

        // --- 图表标题检测 ---
        if (/^图\s*\d|^Figure\s*\d/i.test(text)) {
          figCaptionIndices.push(i);
          isExcluded = true;
        }
        if (/^表\s*\d|^Table\s*\d/i.test(text)) {
          tblCaptionIndices.push(i);
          isExcluded = true;
        }

        // --- 多级标题检测 ---
        var isHeading = false;
        for (var lvl = 1; lvl <= 4; lvl++) {
          var patterns = headingStylePatterns[lvl];
          if (patterns.local.test(styleName) || patterns.en.test(styleName)) {
            var hArr = lvl === 1 ? h1Indices : lvl === 2 ? h2Indices : lvl === 3 ? h3Indices : h4Indices;
            hArr.push(i);
            isHeading = true;
            break;
          }
        }
        if (!isHeading && text) {
          // 格式启发式（level 1-2）
          if (text.length < 60) {
            try {
              var hf = para.Range.Font;
              if (hf.Bold) {
                if (hf.Size >= sizeThresholds[1]) { h1Indices.push(i); isHeading = true; }
                else if (hf.Size >= sizeThresholds[2]) { h2Indices.push(i); isHeading = true; }
              }
            } catch (e) { /* ignore */ }
          }
          // 文本模式匹配
          if (!isHeading) {
            for (var pl = 1; pl <= 4; pl++) {
              if (chapterPatterns[pl] && chapterPatterns[pl].test(text)) {
                var phArr = pl === 1 ? h1Indices : pl === 2 ? h2Indices : pl === 3 ? h3Indices : h4Indices;
                phArr.push(i);
                isHeading = true;
                break;
              }
            }
          }
        }
        if (isHeading) isExcluded = true;

        // --- 正文段落：排除以上所有角色 ---
        if (!isExcluded && text && !/标题|Heading|Title|TOC|目录/i.test(styleName)) {
          bodyIndices.push(i);
        }
      }

      // 标题 fallback：如果没找到，取第一个非空段
      if (titleIndices.length === 0) {
        for (var j = 1; j <= Math.min(count, 3); j++) {
          var ft = (doc.Paragraphs.Item(j).Range.Text || '').replace(/[\r\n]/g, '').trim();
          if (ft) { titleIndices.push(j); break; }
        }
      }

      log('_detectAllInOnePass() done: title=' + titleIndices.length +
          ' h1=' + h1Indices.length + ' h2=' + h2Indices.length +
          ' h3=' + h3Indices.length + ' h4=' + h4Indices.length +
          ' abstract=' + abstractIndices.length + ' keywords=' + keywordsIndices.length +
          ' body=' + bodyIndices.length + ' refs=' + refIndices.length);

      return {
        'title': { role: 'title', paragraphIndices: titleIndices, confidence: titleIndices.length > 0 ? 'medium' : 'low', evidence: 'unified' },
        'heading_1': { role: 'heading_1', paragraphIndices: h1Indices, confidence: 'medium', evidence: 'unified' },
        'heading_2': { role: 'heading_2', paragraphIndices: h2Indices, confidence: 'medium', evidence: 'unified' },
        'heading_3': { role: 'heading_3', paragraphIndices: h3Indices, confidence: 'medium', evidence: 'unified' },
        'heading_4': { role: 'heading_4', paragraphIndices: h4Indices, confidence: 'medium', evidence: 'unified' },
        'abstract': { role: 'abstract', paragraphIndices: abstractIndices, confidence: abstractIndices.length > 0 ? 'high' : 'low', evidence: 'unified' },
        'keywords': { role: 'keywords', paragraphIndices: keywordsIndices, confidence: keywordsIndices.length > 0 ? 'high' : 'low', evidence: 'unified' },
        'body': { role: 'body', paragraphIndices: bodyIndices, confidence: 'medium', evidence: 'unified' },
        'figure_caption': { role: 'figure_caption', paragraphIndices: figCaptionIndices, confidence: 'high', evidence: 'unified' },
        'table_caption': { role: 'table_caption', paragraphIndices: tblCaptionIndices, confidence: 'high', evidence: 'unified' },
        'references': { role: 'references', paragraphIndices: refIndices, confidence: inReferences || refIndices.length > 0 ? 'high' : 'low', evidence: 'unified' },
        'acknowledgment': { role: 'acknowledgment', paragraphIndices: ackIndices, confidence: inAcknowledgment || ackIndices.length > 0 ? 'high' : 'low', evidence: 'unified' },
        'toc': { role: 'toc', paragraphIndices: tocIndices, confidence: 'medium', evidence: 'unified' },
        'footnotes': { role: 'footnotes', paragraphIndices: [], confidence: 'low', evidence: 'not applicable' }
      };
    }, { cooldownMs: 0 });
  }

  // ── 根据 role 调用对应 detect ──

  async function _runDetect(role) {
    log('_runDetect() entering wpsActionQueue for role=' + role);
    return wpsActionQueue.add(async function () {
      log('_runDetect() executing in queue for role=' + role);
      var app = wps.WpsApplication();
      var doc = app.ActiveDocument;
      if (!doc) {
        log('_runDetect() no active document');
        return null;
      }

      log('_runDetect() doc paragraphs=' + doc.Paragraphs.Count);
      var result;
      switch (role) {
        case 'title': result = _detect_title(doc); break;
        case 'heading_1': result = await _detect_headings_level(doc, 1); break;
        case 'heading_2': result = await _detect_headings_level(doc, 2); break;
        case 'heading_3': result = await _detect_headings_level(doc, 3); break;
        case 'heading_4': result = await _detect_headings_level(doc, 4); break;
        case 'abstract': result = _detect_abstract(doc); break;
        case 'keywords': result = _detect_keywords(doc); break;
        case 'body': result = await _detect_body_paragraphs(doc); break;
        case 'figure_caption': result = await _detect_simple_text_pattern(doc, 'figure_caption', /^图\s*\d|^Figure\s*\d/i); break;
        case 'table_caption': result = await _detect_simple_text_pattern(doc, 'table_caption', /^表\s*\d|^Table\s*\d/i); break;
        case 'references': result = _detect_references(doc); break;
        case 'acknowledgment': result = _detect_acknowledgment(doc); break;
        case 'toc': result = _detect_toc_area(doc); break;
        case 'footnotes': result = _detect_footnotes(doc); break;
        default: result = null;
      }
      log('_runDetect() done, role=' + role + ' indices=' + (result && result.paragraphIndices ? result.paragraphIndices.length : 0));
      return result;
    }, { cooldownMs: 0 });
  }

  // 让出 UI 线程，防止 WPS 卡死
  function _yield() {
    return new Promise(function (resolve) { setTimeout(resolve, 0); });
  }

  // ── detect 实现 ──

  function _detect_title(doc) {
    var indices = [];
    var count = Math.min(doc.Paragraphs.Count, 10); // 标题通常在前10段

    for (var i = 1; i <= count; i++) {
      var para = doc.Paragraphs.Item(i);
      var text = (para.Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (!text) continue;

      var styleName = '';
      try { styleName = para.Style.NameLocal || para.Style.Name || ''; } catch (e) { /* ignore */ }

      // 策略1：Style 匹配
      if (/标题|Title/i.test(styleName) && !/标题\s*[2-9]|Heading\s*[2-9]/i.test(styleName)) {
        indices.push(i);
        break;
      }

      // 策略2：格式启发式 - 大字号+居中+加粗
      try {
        var font = para.Range.Font;
        var pf = para.Format;
        if (font.Size >= 18 && font.Bold && pf.Alignment === 1) {
          indices.push(i);
          break;
        }
      } catch (e) { /* ignore */ }
    }

    // 如果没找到，假设第1段非空即为标题
    if (indices.length === 0) {
      for (var j = 1; j <= Math.min(doc.Paragraphs.Count, 3); j++) {
        var t = (doc.Paragraphs.Item(j).Range.Text || '').replace(/[\r\n]/g, '').trim();
        if (t) { indices.push(j); break; }
      }
    }

    return { role: 'title', paragraphIndices: indices, confidence: indices.length > 0 ? 'medium' : 'low', evidence: 'heuristic' };
  }

  async function _detect_headings_level(doc, level) {
    var indices = [];
    var count = doc.Paragraphs.Count;
    var stylePatternLocal = new RegExp('标题\\s*' + level, 'i');
    var stylePatternEn = new RegExp('Heading\\s*' + level, 'i');

    var sizeThresholds = { 1: 16, 2: 14, 3: 12, 4: 11 };
    var threshold = sizeThresholds[level] || 12;

    var chineseChapterPatterns = {
      1: /^第[一二三四五六七八九十百]+[章节]/,
      2: /^第[一二三四五六七八九十百]+[条款]|^\d+\.\d+\s/,
      3: /^\d+\.\d+\.\d+\s/,
      4: /^\d+\.\d+\.\d+\.\d+\s/
    };

    for (var i = 1; i <= count; i++) {
      // 每 40 段让出 UI 线程，防止 WPS 卡死
      if (i % 40 === 0) await _yield();

      var para = doc.Paragraphs.Item(i);

      // 策略1：Style 匹配（1次 COM 调用，最快路径）
      var styleName = '';
      try { styleName = para.Style.NameLocal || para.Style.Name || ''; } catch (e) { /* ignore */ }
      if (stylePatternLocal.test(styleName) || stylePatternEn.test(styleName)) {
        indices.push(i);
        continue;
      }

      // 策略2+3 需要读文本，只读一次
      var text = '';
      try { text = (para.Range.Text || '').replace(/[\r\n]/g, '').trim(); } catch (e) { /* ignore */ }
      if (!text) continue;

      // 策略2：格式启发式（仅 level 1-2，且文本不太长才值得检查字号）
      if (level <= 2 && text.length < 60) {
        try {
          var font = para.Range.Font;
          if (font.Size >= threshold && font.Bold) {
            indices.push(i);
            continue;
          }
        } catch (e) { /* ignore */ }
      }

      // 策略3：文本模式匹配
      if (chineseChapterPatterns[level] && chineseChapterPatterns[level].test(text)) {
        indices.push(i);
      }
    }

    return { role: 'heading_' + level, paragraphIndices: indices, confidence: 'medium', evidence: 'style+heuristic' };
  }

  function _detect_abstract(doc) {
    var indices = [];
    var count = Math.min(doc.Paragraphs.Count, 30);

    var abstractStart = -1;
    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (/^摘\s*要|^Abstract/i.test(text)) {
        abstractStart = i;
        // 如果"摘要"独占一段，从下一段开始算内容
        if (text.length < 10) {
          abstractStart = i + 1;
        } else {
          indices.push(i);
        }
        break;
      }
    }

    if (abstractStart > 0) {
      for (var j = abstractStart; j <= Math.min(count, abstractStart + 5); j++) {
        if (indices.includes(j)) continue;
        var t = (doc.Paragraphs.Item(j).Range.Text || '').replace(/[\r\n]/g, '').trim();
        if (!t) break;
        if (/^关键词|^Keywords|^Key\s*words/i.test(t)) break;
        var sn = '';
        try { sn = doc.Paragraphs.Item(j).Style.NameLocal || ''; } catch (e) { /* ignore */ }
        if (/标题|Heading/i.test(sn)) break;
        indices.push(j);
      }
    }

    return { role: 'abstract', paragraphIndices: indices, confidence: indices.length > 0 ? 'high' : 'low', evidence: 'text pattern' };
  }

  function _detect_keywords(doc) {
    var indices = [];
    var count = Math.min(doc.Paragraphs.Count, 30);

    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (/^关键词|^Keywords|^Key\s*words/i.test(text)) {
        indices.push(i);
        break;
      }
    }

    return { role: 'keywords', paragraphIndices: indices, confidence: indices.length > 0 ? 'high' : 'low', evidence: 'text pattern' };
  }

  function _detect_toc_area(doc) {
    var indices = [];
    var count = Math.min(doc.Paragraphs.Count, 50);

    for (var i = 1; i <= count; i++) {
      var para = doc.Paragraphs.Item(i);
      var styleName = '';
      try { styleName = para.Style.NameLocal || para.Style.Name || ''; } catch (e) { /* ignore */ }
      if (/目录|TOC/i.test(styleName)) {
        indices.push(i);
      }
      var text = (para.Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (/^目\s*录$/.test(text)) {
        if (!indices.includes(i)) indices.push(i);
      }
    }

    return { role: 'toc', paragraphIndices: indices, confidence: 'medium', evidence: 'style+text' };
  }

  async function _detect_body_paragraphs(doc) {
    var excludeRoles = ['title', 'heading_1', 'heading_2', 'heading_3', 'heading_4',
                        'abstract', 'keywords', 'figure_caption', 'table_caption',
                        'references', 'acknowledgment', 'toc'];

    var excludedIndices = {};
    for (var r = 0; r < excludeRoles.length; r++) {
      var cached = DocumentStructureCache.getIndices(excludeRoles[r]);
      if (cached) {
        for (var c = 0; c < cached.length; c++) {
          excludedIndices[cached[c]] = true;
        }
      }
    }

    var indices = [];
    var count = doc.Paragraphs.Count;

    for (var i = 1; i <= count; i++) {
      if (i % 40 === 0) await _yield();
      if (excludedIndices[i]) continue;

      var para = doc.Paragraphs.Item(i);
      var text = '';
      try { text = (para.Range.Text || '').replace(/[\r\n]/g, '').trim(); } catch (e) { /* ignore */ }
      if (!text) continue;

      var styleName = '';
      try { styleName = para.Style.NameLocal || para.Style.Name || ''; } catch (e) { /* ignore */ }

      if (/标题|Heading|Title|TOC|目录/i.test(styleName)) continue;
      if (/^图\s*\d|^表\s*\d|^Figure\s*\d|^Table\s*\d/i.test(text)) continue;
      if (/^参考文献|^References/i.test(text)) continue;
      if (/^致\s*谢|^Acknowledgment/i.test(text)) continue;

      indices.push(i);
    }

    return { role: 'body', paragraphIndices: indices, confidence: 'medium', evidence: 'exclusion' };
  }

  async function _detect_simple_text_pattern(doc, role, pattern) {
    var indices = [];
    var count = doc.Paragraphs.Count;

    for (var i = 1; i <= count; i++) {
      if (i % 40 === 0) await _yield();
      var text = '';
      try { text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim(); } catch (e) { /* ignore */ }
      if (pattern.test(text)) {
        indices.push(i);
      }
    }

    return { role: role, paragraphIndices: indices, confidence: 'high', evidence: 'text pattern' };
  }

  // 保留旧名字用于 skill registry 注册
  function _detect_figure_captions(doc) {
    var indices = [];
    var count = doc.Paragraphs.Count;

    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (/^图\s*\d|^Figure\s*\d/i.test(text)) {
        indices.push(i);
      }
    }

    return { role: 'figure_caption', paragraphIndices: indices, confidence: 'high', evidence: 'text pattern' };
  }

  function _detect_table_captions(doc) {
    var indices = [];
    var count = doc.Paragraphs.Count;

    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();
      if (/^表\s*\d|^Table\s*\d/i.test(text)) {
        indices.push(i);
      }
    }

    return { role: 'table_caption', paragraphIndices: indices, confidence: 'high', evidence: 'text pattern' };
  }

  function _detect_references(doc) {
    var indices = [];
    var count = doc.Paragraphs.Count;
    var inReferences = false;

    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();

      if (/^参考文献|^References/i.test(text)) {
        inReferences = true;
        indices.push(i);
        continue;
      }

      if (inReferences) {
        if (!text) continue; // 空行跳过
        // 碰到新的大标题则结束
        var sn = '';
        try { sn = doc.Paragraphs.Item(i).Style.NameLocal || ''; } catch (e) { /* ignore */ }
        if (/标题\s*1|Heading\s*1/i.test(sn)) break;
        if (/^致\s*谢|^附\s*录|^Acknowledgment|^Appendix/i.test(text)) break;
        indices.push(i);
      }
    }

    return { role: 'references', paragraphIndices: indices, confidence: inReferences ? 'high' : 'low', evidence: 'text pattern' };
  }

  function _detect_acknowledgment(doc) {
    var indices = [];
    var count = doc.Paragraphs.Count;
    var inSection = false;

    for (var i = 1; i <= count; i++) {
      var text = (doc.Paragraphs.Item(i).Range.Text || '').replace(/[\r\n]/g, '').trim();

      if (/^致\s*谢|^Acknowledgment/i.test(text)) {
        inSection = true;
        indices.push(i);
        continue;
      }

      if (inSection) {
        if (!text) continue;
        var sn = '';
        try { sn = doc.Paragraphs.Item(i).Style.NameLocal || ''; } catch (e) { /* ignore */ }
        if (/标题\s*1|Heading\s*1/i.test(sn)) break;
        if (/^附\s*录|^Appendix/i.test(text)) break;
        indices.push(i);
      }
    }

    return { role: 'acknowledgment', paragraphIndices: indices, confidence: inSection ? 'high' : 'low', evidence: 'text pattern' };
  }

  function _detect_footnotes(doc) {
    // WPS 脚注通过 doc.Footnotes 访问，段落索引无直接对应
    return { role: 'footnotes', paragraphIndices: [], confidence: 'low', evidence: 'not applicable' };
  }

  // ── 注册 detect 技能的 execute 函数 ──

  var registry = window.FormatSkillRegistry;
  if (registry && registry.skills) {
    registry.skills['detect.title'].execute = function () { return _detect_title(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.headings'].execute = function (params) {
      var doc = wps.WpsApplication().ActiveDocument;
      var levels = (params && params.levels) || [1, 2, 3, 4];
      var results = {};
      for (var l = 0; l < levels.length; l++) {
        results[levels[l]] = _detect_headings_level(doc, levels[l]);
      }
      return results;
    };
    registry.skills['detect.abstract'].execute = function () { return _detect_abstract(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.keywords'].execute = function () { return _detect_keywords(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.toc_area'].execute = function () { return _detect_toc_area(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.body_paragraphs'].execute = function () { return _detect_body_paragraphs(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.figure_captions'].execute = function () { return _detect_figure_captions(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.table_captions'].execute = function () { return _detect_table_captions(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.images'].execute = function () { return { role: 'images', paragraphIndices: [], confidence: 'low' }; };
    registry.skills['detect.tables'].execute = function () { return { role: 'tables', paragraphIndices: [], confidence: 'low' }; };
    registry.skills['detect.references'].execute = function () { return _detect_references(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.acknowledgment'].execute = function () { return _detect_acknowledgment(wps.WpsApplication().ActiveDocument); };
    registry.skills['detect.footnotes'].execute = function () { return _detect_footnotes(wps.WpsApplication().ActiveDocument); };
  }

  window.DocumentStructureCache = DocumentStructureCache;
})();
