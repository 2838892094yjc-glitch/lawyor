#!/usr/bin/env node

/**
 * Local AI bridge for WPS add-in.
 * Accepts POST /analyze { text, chunk_size } and forwards to Kimi/Moonshot API.
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const {
  validateAIOutput,
  parseAIOutput,
  VALID_FORMAT_FNS,
  VALID_TYPES,
  VALID_MODES
} = require('./ai-parser.js');

function sanitizeEnv(value) {
  if (typeof value !== 'string') return '';
  const trimmed = value.trim().replace(/^['\"]+|['\"]+$/g, '');
  if (!trimmed) return '';
  // Guard against accidental multi-command paste into env vars.
  return trimmed.split(/\s+/)[0];
}

function parseNumberEnv(name, fallback) {
  const value = Number(sanitizeEnv(process.env[name] || String(fallback)));
  return Number.isFinite(value) ? value : fallback;
}

function parseBoolEnv(name, fallback) {
  const raw = sanitizeEnv(process.env[name] || '');
  if (!raw) return fallback;
  if (['1', 'true', 'yes', 'on'].includes(raw.toLowerCase())) return true;
  if (['0', 'false', 'no', 'off'].includes(raw.toLowerCase())) return false;
  return fallback;
}

function loadEnvFiles() {
  const candidates = [
    path.join(__dirname, '.env.kimi.local'),
    path.join(__dirname, '.env.kimi')
  ];

  for (const filePath of candidates) {
    if (!fs.existsSync(filePath)) continue;

    const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
    for (const rawLine of lines) {
      const line = String(rawLine || '').trim();
      if (!line || line.startsWith('#')) continue;

      const match = line.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
      if (!match) continue;

      const key = match[1];
      let value = match[2].trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }

      if (!process.env[key]) {
        process.env[key] = value;
      }
    }
  }
}

loadEnvFiles();

const PORT = parseNumberEnv('PORT', parseNumberEnv('AI_BRIDGE_PORT', 8765));
const MAX_INPUT_CHARS = parseNumberEnv('MAX_INPUT_CHARS', 24000);
const SKILL_PROMPT_PATH = process.env.SKILL_PROMPT_PATH || path.join(__dirname, 'contract-template-transformation-skill', 'SKILL.md');
const KIMI_API_KEY = sanitizeEnv(process.env.KIMI_API_KEY || process.env.MOONSHOT_API_KEY);
const KIMI_API_URL = process.env.KIMI_API_URL || 'https://api.moonshot.cn/v1/chat/completions';

const LEGACY_MODEL = sanitizeEnv(process.env.KIMI_MODEL || process.env.MOONSHOT_MODEL) || 'moonshot-v1-32k';
const LEGACY_TEMPERATURE = parseNumberEnv('KIMI_TEMPERATURE', 0);
const LEGACY_TIMEOUT_MS = parseNumberEnv('KIMI_REQUEST_TIMEOUT_MS', 30000);
const LEGACY_MAX_TOKENS = parseNumberEnv('KIMI_MAX_TOKENS', 1200);

const PIPELINE_ENABLED = parseBoolEnv('KIMI_PIPELINE_ENABLED', true);
const PIPELINE_FALLBACK_DIRECT = parseBoolEnv('KIMI_PIPELINE_FALLBACK_DIRECT', true);

const SEMANTIC_MODEL = sanitizeEnv(process.env.KIMI_SEMANTIC_MODEL) || 'kimi-k2.5';
const SEMANTIC_TEMPERATURE = parseNumberEnv('KIMI_SEMANTIC_TEMPERATURE', 1);
const SEMANTIC_MAX_TOKENS = parseNumberEnv('KIMI_SEMANTIC_MAX_TOKENS', 1800);
const SEMANTIC_TIMEOUT_MS = parseNumberEnv('KIMI_SEMANTIC_TIMEOUT_MS', 70000);

const STRUCT_MODEL = sanitizeEnv(process.env.KIMI_STRUCT_MODEL || process.env.KIMI_MODEL || process.env.MOONSHOT_MODEL) || 'moonshot-v1-32k';
const STRUCT_TEMPERATURE = parseNumberEnv('KIMI_STRUCT_TEMPERATURE', LEGACY_TEMPERATURE);
const STRUCT_MAX_TOKENS = parseNumberEnv('KIMI_STRUCT_MAX_TOKENS', LEGACY_MAX_TOKENS);
const STRUCT_TIMEOUT_MS = parseNumberEnv('KIMI_STRUCT_TIMEOUT_MS', LEGACY_TIMEOUT_MS);

const FORMAT_MODEL = sanitizeEnv(process.env.KIMI_FORMAT_MODEL || STRUCT_MODEL);
const FORMAT_TIMEOUT_MS = parseNumberEnv('KIMI_FORMAT_TIMEOUT_MS', 60000);
const NETWORK_RETRIES = Math.max(0, parseNumberEnv('KIMI_NETWORK_RETRIES', 1));

const bridgeMetrics = {
  startedAt: Date.now(),
  requestsTotal: 0,
  analyzeSuccess: 0,
  analyzeFailure: 0,
  lastAnalyzeAt: null,
  lastLatencyMs: null,
  lastVariablesCount: 0,
  lastVariableTags: [],
  lastSchemaWarnings: 0,
  lastAnalyzeIso: null,
  lastPipelineMode: null,
  lastStageLatencies: null,
  lastError: ''
};

const temperatureWarnedModels = new Set();
const EXPLICIT_PLACEHOLDER_RE = /(【[^】]{0,80}】|_{2,}|\[\s*\]|\(\s*\)|（\s*\)|﹍{2,}|\*{2,})/;
const DYNAMIC_HINT_RE = /(合同编号|订单编号|编号|日期|年|月|日|期限|工作日|小时|金额|价款|单价|税点|税号|账户|账号|开户行|收款人|联系人|手机|电话|邮箱|电子邮件|地址|交货|批次|比例|违约金|购货方|采购人员|审批人)/;

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function sendJson(res, statusCode, payload) {
  setCors(res);
  res.statusCode = statusCode;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeModelName(model) {
  return String(model || '').trim().toLowerCase();
}

function resolveTemperatureForModel(model, requestedTemperature, stageLabel) {
  const normalized = normalizeModelName(model);
  const isKimiK25 = normalized.includes('kimi-k2.5');

  if (isKimiK25 && Number(requestedTemperature) !== 1) {
    if (!temperatureWarnedModels.has(normalized)) {
      console.warn(`[Kimi Bridge] ${stageLabel} forcing temperature=1 for model=${model}`);
      temperatureWarnedModels.add(normalized);
    }
    return 1;
  }

  return requestedTemperature;
}

function shouldRetryRequest(err, statusCode) {
  if (typeof statusCode === 'number') {
    return statusCode === 408 || statusCode === 409 || statusCode === 429 || (statusCode >= 500 && statusCode <= 599);
  }

  const msg = String((err && err.message) || err || '').toLowerCase();
  if (!msg) return false;
  return msg.includes('fetch failed') || msg.includes('econnreset') || msg.includes('etimedout') || msg.includes('network');
}

function readSkillPrompt() {
  return fs.readFileSync(SKILL_PROMPT_PATH, 'utf8');
}

function extractTextContent(chatResponse) {
  if (!chatResponse || !Array.isArray(chatResponse.choices) || chatResponse.choices.length === 0) {
    throw new Error('Kimi response missing choices');
  }

  const msg = chatResponse.choices[0] && chatResponse.choices[0].message;
  if (!msg || msg.content === undefined || msg.content === null) {
    throw new Error('Kimi response missing message content');
  }

  if (typeof msg.content === 'string') {
    return msg.content.trim();
  }

  if (Array.isArray(msg.content)) {
    const text = msg.content
      .map((item) => {
        if (!item) return '';
        if (typeof item === 'string') return item;
        if (typeof item.text === 'string') return item.text;
        return '';
      })
      .join('\n')
      .trim();

    if (text) return text;
  }

  throw new Error('Unsupported Kimi message content format');
}

function tryParseJsonObject(rawText) {
  const text = String(rawText || '').trim();

  try {
    return JSON.parse(text);
  } catch (_err) {
    // Continue.
  }

  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch) {
    const fenced = fenceMatch[1].trim();
    try {
      return JSON.parse(fenced);
    } catch (_err) {
      // Continue.
    }
  }

  for (let start = 0; start < text.length; start += 1) {
    if (text[start] !== '{') continue;

    let depth = 0;
    let inString = false;
    let escaped = false;

    for (let i = start; i < text.length; i += 1) {
      const ch = text[i];

      if (inString) {
        if (escaped) {
          escaped = false;
        } else if (ch === '\\') {
          escaped = true;
        } else if (ch === '"') {
          inString = false;
        }
        continue;
      }

      if (ch === '"') {
        inString = true;
        continue;
      }

      if (ch === '{') depth += 1;
      if (ch === '}') depth -= 1;

      if (depth === 0) {
        const candidate = text.slice(start, i + 1);
        try {
          return JSON.parse(candidate);
        } catch (_err) {
          break;
        }
      }
    }
  }

  throw new Error('Unable to parse JSON from Kimi response');
}

function validateOutputShape(parsed) {
  if (!parsed || typeof parsed !== 'object' || !Array.isArray(parsed.variables)) {
    throw new Error('Output must be JSON object with variables array');
  }
  return parsed;
}


function firstNonEmpty(...values) {
  for (const value of values) {
    if (typeof value === 'string' && value.trim()) return value.trim();
  }
  return '';
}

function sanitizeText(value, maxLen = 2000) {
  if (typeof value !== 'string') return '';
  return value.replace(/\u0000/g, '').trim().slice(0, maxLen);
}

function normalizeListFromOutput(rawParsed) {
  if (Array.isArray(rawParsed)) return rawParsed;
  if (!rawParsed || typeof rawParsed !== 'object') return [];

  if (Array.isArray(rawParsed.variables)) return rawParsed.variables;
  if (Array.isArray(rawParsed.items)) return rawParsed.items;
  if (Array.isArray(rawParsed.fields)) return rawParsed.fields;
  if (Array.isArray(rawParsed.candidates)) return rawParsed.candidates;

  if (rawParsed.output && typeof rawParsed.output === 'object') {
    return normalizeListFromOutput(rawParsed.output);
  }

  return [];
}

function detectPlaceholderMarker(text) {
  const raw = String(text || '');
  const match = raw.match(EXPLICIT_PLACEHOLDER_RE);
  if (match) {
    return {
      marker: match[0],
      explicit: true
    };
  }

  if (/[:：]\s{2,}/.test(raw) || /[:：]\s*$/.test(raw)) {
    return {
      marker: '____',
      explicit: true
    };
  }

  return {
    marker: '',
    explicit: false
  };
}

function guessType(text) {
  const value = String(text || '');
  if (!value) return 'text';

  if (/(日期|年月日|签署日|生效日)/.test(value)) return 'date';
  if (/(金额|价款|税点|税率|数量|比例|百分|违约金|小时|日|月|年|批次)/.test(value)) return 'number';
  if (/(信息如下|说明|方案|整改|附件|账户信息|开票信息)/.test(value)) return 'textarea';
  return 'text';
}

function normalizeType(value, hintText) {
  const candidate = String(value || '').trim();
  if (VALID_TYPES.includes(candidate)) return candidate;
  return guessType(hintText);
}

function normalizeFormatFn(value, hintText) {
  const candidate = String(value || '').trim();
  if (VALID_FORMAT_FNS.includes(candidate)) return candidate;
  if (/(日期|年月日|签署日|生效日)/.test(String(hintText || ''))) return 'dateUnderline';
  return 'none';
}

function normalizeMode(value) {
  const candidate = String(value || '').trim();
  if (VALID_MODES.includes(candidate)) return candidate;
  return 'insert';
}

function normalizeTag(value, index, usedTags) {
  let tag = sanitizeText(String(value || ''), 64).replace(/[^A-Za-z0-9_]/g, '_');
  if (!tag) tag = `Field_${index + 1}`;
  if (!/^[A-Za-z_]/.test(tag)) tag = `Field_${tag}`;

  let unique = tag;
  let counter = 2;
  while (usedTags.has(unique)) {
    unique = `${tag}_${counter}`;
    counter += 1;
  }
  usedTags.add(unique);
  return unique;
}

function inferLabel(rawLabel, rawPrefix, rawContext, index) {
  const label = firstNonEmpty(rawLabel);
  if (label) return sanitizeText(label, 120);

  const fromPrefix = sanitizeText(rawPrefix, 120).replace(/[：:]+$/, '');
  if (fromPrefix) return fromPrefix;

  const fromContext = sanitizeText(rawContext, 120).slice(0, 24);
  if (fromContext) return fromContext;

  return `字段${index + 1}`;
}

function normalizeOptions(value) {
  if (!Array.isArray(value)) return [];
  return value.map((item) => sanitizeText(String(item || ''), 80)).filter(Boolean).slice(0, 50);
}

function shouldKeepVariable({ explicitMarker, label, context, placeholder, mode }) {
  if (mode === 'paragraph') return true;

  const hintText = [label, context].join(' ');
  const hasDynamicHint = DYNAMIC_HINT_RE.test(hintText);
  const compactPlaceholder = String(placeholder || '').replace(/\s+/g, ' ').trim();
  const compactContext = String(context || '').replace(/\s+/g, ' ').trim();

  if (explicitMarker) return true;

  if (hasDynamicHint && compactPlaceholder.length <= 120 && compactContext.length <= 120) {
    return true;
  }

  return false;
}

function normalizeVariable(raw, index, sourceText, usedTags) {
  const source = raw && typeof raw === 'object' ? raw : {};

  const context = sanitizeText(firstNonEmpty(source.context, source.anchor, source.sentence, source.text), 600);
  const prefix = sanitizeText(firstNonEmpty(source.prefix, source.left, source.head), 220);
  const suffix = sanitizeText(firstNonEmpty(source.suffix, source.right, source.tail), 220);
  const label = inferLabel(source.label || source.name || source.field, prefix, context, index);

  const rawPlaceholder = sanitizeText(firstNonEmpty(source.placeholder, source.value, source.blank, source.variable), 600);
  const markerFromText = detectPlaceholderMarker([context, prefix, suffix].join(' '));
  const placeholder = rawPlaceholder || markerFromText.marker || '____';

  const mode = normalizeMode(firstNonEmpty(source.mode, source.mode_hint));
  let type = normalizeType(firstNonEmpty(source.type, source.type_hint), [label, context, prefix].join(' '));
  const options = normalizeOptions(source.options);
  if ((type === 'select' || type === 'radio') && options.length === 0) {
    type = 'text';
  }

  const variable = {
    context: context || sanitizeText([prefix, placeholder, suffix].join(''), 600) || '未命名上下文',
    prefix,
    placeholder,
    suffix,
    label,
    tag: normalizeTag(firstNonEmpty(source.tag, source.tag_hint, source.name), index, usedTags),
    type,
    formatFn: normalizeFormatFn(firstNonEmpty(source.formatFn, source.format_hint), [label, context, placeholder].join(' ')),
    mode
  };

  if (options.length > 0 && (type === 'select' || type === 'radio')) {
    variable.options = options;
  }

  const keep = shouldKeepVariable({
    explicitMarker: markerFromText.explicit || EXPLICIT_PLACEHOLDER_RE.test(rawPlaceholder),
    label: variable.label,
    context: variable.context,
    placeholder: variable.placeholder,
    mode: variable.mode
  });

  return keep ? variable : null;
}

function fallbackExtractVariablesFromText(text) {
  const lines = String(text || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  const usedTags = new Set();
  const result = [];

  for (const line of lines) {
    if (result.length >= 80) break;

    const markerInfo = detectPlaceholderMarker(line);
    const hasHint = DYNAMIC_HINT_RE.test(line);
    if (!markerInfo.explicit) {
      if (!hasHint) continue;
      if (line.length > 60) continue;
      if (!/[:：]/.test(line)) continue;
    }

    const marker = markerInfo.marker || '____';
    const idx = line.indexOf(marker);
    let prefix = '';
    let suffix = '';
    if (idx >= 0) {
      prefix = line.slice(0, idx);
      suffix = line.slice(idx + marker.length);
    } else {
      prefix = line;
    }

    const label = inferLabel('', prefix, line, result.length);
    const type = normalizeType('', `${label} ${line}`);
    const variable = {
      context: sanitizeText(line, 600),
      prefix: sanitizeText(prefix, 220),
      placeholder: marker || '____',
      suffix: sanitizeText(suffix, 220),
      label,
      tag: normalizeTag('', result.length, usedTags),
      type,
      formatFn: normalizeFormatFn('', `${label} ${line}`),
      mode: 'insert'
    };

    result.push(variable);
  }

  return result;
}

function coerceOutputShape(rawParsed, sourceText) {
  const list = normalizeListFromOutput(rawParsed);
  const usedTags = new Set();
  const normalized = list
    .map((item, index) => normalizeVariable(item, index, sourceText, usedTags))
    .filter(Boolean);

  if (normalized.length > 0) {
    return { variables: normalized };
  }

  const fallback = fallbackExtractVariablesFromText(sourceText);
  if (fallback.length > 0) {
    return { variables: fallback };
  }

  return { variables: [] };
}

function enforcePluginCompatibility(parsed) {
  const validation = validateAIOutput(parsed);
  if (!validation.valid) {
    const detail = (validation.errors || []).slice(0, 5).join(' | ');
    throw new Error('Plugin schema validation failed: ' + detail);
  }

  const parseResult = parseAIOutput(parsed, { validateOnly: true });
  if (!parseResult.success) {
    const detail = (parseResult.errors || []).slice(0, 5).join(' | ');
    throw new Error('Plugin parse validation failed: ' + detail);
  }

  const warnings = validation.warnings || [];
  if (warnings.length > 0) {
    console.warn('[Kimi Bridge] schema warnings: ' + warnings.slice(0, 5).join(' | '));
  }

  return { parsed, warnings };
}

async function callChatCompletion({ messages, model, temperature, maxTokens, timeoutMs, stageLabel }) {
  const startedAt = Date.now();
  const resolvedTemperature = resolveTemperatureForModel(model, temperature, stageLabel);
  let lastError;

  for (let attempt = 0; attempt <= NETWORK_RETRIES; attempt += 1) {
    const attemptLabel = `${stageLabel} attempt ${attempt + 1}/${NETWORK_RETRIES + 1}`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const response = await fetch(KIMI_API_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${KIMI_API_KEY}`
        },
        signal: controller.signal,
        body: JSON.stringify({
          model,
          temperature: resolvedTemperature,
          max_tokens: maxTokens,
          messages
        })
      });

      const raw = await response.text();
      if (!response.ok) {
        const err = new Error(`${stageLabel} API error ${response.status}: ${raw}`);
        const retryable = shouldRetryRequest(err, response.status);
        if (retryable && attempt < NETWORK_RETRIES) {
          await delay(400 * (attempt + 1));
          continue;
        }
        throw err;
      }

      let parsedResponse;
      try {
        parsedResponse = JSON.parse(raw);
      } catch (err) {
        const jsonErr = new Error(`${stageLabel} invalid API JSON: ${err.message}`);
        if (attempt < NETWORK_RETRIES) {
          await delay(250 * (attempt + 1));
          continue;
        }
        throw jsonErr;
      }

      return {
        contentText: extractTextContent(parsedResponse),
        latencyMs: Date.now() - startedAt,
        finishReason: parsedResponse.choices && parsedResponse.choices[0] ? parsedResponse.choices[0].finish_reason : null,
        attempts: attempt + 1
      };
    } catch (err) {
      clearTimeout(timeoutId);

      if (err && err.name === 'AbortError') {
        lastError = new Error(`${attemptLabel} timeout after ${timeoutMs}ms`);
      } else {
        const message = err && err.message ? err.message : String(err);
        lastError = new Error(`${attemptLabel} failed: ${message}`);
      }

      const retryable = shouldRetryRequest(err);
      if (!retryable || attempt >= NETWORK_RETRIES) {
        throw lastError;
      }

      await delay(500 * (attempt + 1));
    } finally {
      clearTimeout(timeoutId);
    }
  }

  throw lastError || new Error(`${stageLabel} failed unexpectedly`);
}

function normalizeSemanticVariables(rawParsed) {
  if (!rawParsed) return [];

  let list = [];
  if (Array.isArray(rawParsed.variables)) {
    list = rawParsed.variables;
  } else if (Array.isArray(rawParsed.candidates)) {
    list = rawParsed.candidates;
  } else if (Array.isArray(rawParsed.items)) {
    list = rawParsed.items;
  } else if (Array.isArray(rawParsed)) {
    list = rawParsed;
  }

  return list
    .filter((item) => item && typeof item === 'object')
    .map((item) => ({
      context: typeof item.context === 'string' ? item.context : '',
      prefix: typeof item.prefix === 'string' ? item.prefix : '',
      placeholder: typeof item.placeholder === 'string' ? item.placeholder : '',
      suffix: typeof item.suffix === 'string' ? item.suffix : '',
      label: typeof item.label === 'string' ? item.label : '',
      tag_hint: typeof item.tag_hint === 'string' ? item.tag_hint : (typeof item.tag === 'string' ? item.tag : ''),
      type_hint: typeof item.type_hint === 'string' ? item.type_hint : (typeof item.type === 'string' ? item.type : ''),
      format_hint: typeof item.format_hint === 'string' ? item.format_hint : (typeof item.formatFn === 'string' ? item.formatFn : ''),
      mode_hint: typeof item.mode_hint === 'string' ? item.mode_hint : (typeof item.mode === 'string' ? item.mode : ''),
      options: Array.isArray(item.options) ? item.options : [],
      confidence: typeof item.confidence === 'string' ? item.confidence : 'medium',
      reason: typeof item.reason === 'string' ? item.reason : ''
    }))
    .filter((item) => item.placeholder || item.label || item.context);
}

function buildDirectAnalyzePrompt({ skillPrompt, clippedText }) {
  const system = [
    'You are a contract templating engine.',
    'Return ONLY a valid JSON object: {"variables": [...]}',
    'Do not include markdown or explanations.',
    `Follow this skill spec exactly:\n\n${skillPrompt}`
  ].join('\n\n');

  return {
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: `Contract text:\n${clippedText}` }
    ],
    model: LEGACY_MODEL,
    temperature: LEGACY_TEMPERATURE,
    maxTokens: LEGACY_MAX_TOKENS,
    timeoutMs: LEGACY_TIMEOUT_MS,
    stageLabel: 'Kimi direct-analyze'
  };
}

async function runDirectAnalyze({ skillPrompt, clippedText }) {
  const chatResult = await callChatCompletion(buildDirectAnalyzePrompt({ skillPrompt, clippedText }));
  const parsedOutput = tryParseJsonObject(chatResult.contentText);
  const shaped = coerceOutputShape(parsedOutput, clippedText);
  validateOutputShape(shaped);
  const checked = enforcePluginCompatibility(shaped);
  return {
    parsed: checked.parsed,
    warnings: checked.warnings,
    meta: {
      mode: 'direct',
      stageLatencies: {
        directMs: chatResult.latencyMs
      },
      finishReasons: {
        direct: chatResult.finishReason
      }
    }
  };
}

async function runPipelineAnalyze({ skillPrompt, clippedText }) {
  const semanticSystem = [
    'You are a senior contract variable mining expert.',
    'Task: identify as many template variables as possible from the contract text.',
    'Output ONLY JSON object with key "variables" (array).',
    'Each variable may include: context,prefix,placeholder,suffix,label,tag_hint,type_hint,format_hint,mode_hint,options,confidence,reason.',
    'This stage focuses on recall and evidence. Do not force strict plugin schema yet.',
    `Skill strategy to follow:\n\n${skillPrompt}`
  ].join('\n\n');

  const semanticResult = await callChatCompletion({
    messages: [
      { role: 'system', content: semanticSystem },
      { role: 'user', content: `Contract text:\n${clippedText}` }
    ],
    model: SEMANTIC_MODEL,
    temperature: SEMANTIC_TEMPERATURE,
    maxTokens: SEMANTIC_MAX_TOKENS,
    timeoutMs: SEMANTIC_TIMEOUT_MS,
    stageLabel: 'Kimi semantic-stage'
  });

  let semanticRepairMs = null;
  let semanticParsed;
  let semanticTextForNormalize = semanticResult.contentText;
  try {
    semanticParsed = tryParseJsonObject(semanticResult.contentText);
  } catch (_err) {
    const semanticRepair = await callChatCompletion({
      messages: [
        {
          role: 'system',
          content: 'Convert the provided semantic analysis text into JSON object with key variables only. Output JSON only.'
        },
        {
          role: 'user',
          content: [
            'Contract text:\n' + clippedText,
            'Semantic analysis text:\n' + semanticResult.contentText
          ].join('\n\n')
        }
      ],
      model: STRUCT_MODEL,
      temperature: STRUCT_TEMPERATURE,
      maxTokens: STRUCT_MAX_TOKENS,
      timeoutMs: STRUCT_TIMEOUT_MS,
      stageLabel: 'Kimi semantic-repair-stage'
    });
    semanticRepairMs = semanticRepair.latencyMs;
    semanticTextForNormalize = semanticRepair.contentText;
    semanticParsed = tryParseJsonObject(semanticRepair.contentText);
  }

  const semanticVariables = normalizeSemanticVariables(semanticParsed);
  if (semanticVariables.length === 0) {
    console.warn('[Kimi Bridge] semantic-stage yielded zero normalized candidates; normalize-stage will use raw semantic text');
  }

  const normalizeSystem = [
    'You are a strict schema normalizer for WPS contract add-in.',
    'Convert semantic candidates into STRICT plugin JSON object: {"variables": [...]}',
    'Output only JSON, no markdown, no extra text.',
    `Allowed type: ${VALID_TYPES.join(', ')}`,
    `Allowed formatFn: ${VALID_FORMAT_FNS.join(', ')}`,
    `Allowed mode: ${VALID_MODES.join(', ')}`,
    'Required keys for each variable: context,prefix,placeholder,suffix,label,tag,type,formatFn,mode.',
    'For select/radio type, options must be non-empty array.',
    'Keep outputs concise and deterministic.'
  ].join('\n\n');

  const semanticPayload = JSON.stringify(
    {
      variables: semanticVariables.slice(0, 120),
      rawSemanticText: String(semanticTextForNormalize || '').slice(0, 6000)
    },
    null,
    2
  );

  const normalizeResult = await callChatCompletion({
    messages: [
      { role: 'system', content: normalizeSystem },
      {
        role: 'user',
        content: [
          `Contract text:\n${clippedText}`,
          `Semantic candidates JSON:\n${semanticPayload}`
        ].join('\n\n')
      }
    ],
    model: STRUCT_MODEL,
    temperature: STRUCT_TEMPERATURE,
    maxTokens: STRUCT_MAX_TOKENS,
    timeoutMs: STRUCT_TIMEOUT_MS,
    stageLabel: 'Kimi normalize-stage'
  });

  let normalizeRepairMs = null;
  let normalizedParsed;
  try {
    normalizedParsed = tryParseJsonObject(normalizeResult.contentText);
  } catch (_err) {
    const normalizeRepair = await callChatCompletion({
      messages: [
        {
          role: 'system',
          content: 'Convert the provided text to STRICT plugin JSON object {"variables": [...]} only. No markdown.'
        },
        {
          role: 'user',
          content: [
            'Contract text:\n' + clippedText,
            'Raw normalize-stage output:\n' + normalizeResult.contentText
          ].join('\n\n')
        }
      ],
      model: STRUCT_MODEL,
      temperature: STRUCT_TEMPERATURE,
      maxTokens: STRUCT_MAX_TOKENS,
      timeoutMs: STRUCT_TIMEOUT_MS,
      stageLabel: 'Kimi normalize-repair-stage'
    });
    normalizeRepairMs = normalizeRepair.latencyMs;
    normalizedParsed = tryParseJsonObject(normalizeRepair.contentText);
  }

  const shaped = coerceOutputShape(normalizedParsed, clippedText);
  validateOutputShape(shaped);
  const checked = enforcePluginCompatibility(shaped);

  return {
    parsed: checked.parsed,
    warnings: checked.warnings,
    meta: {
      mode: 'pipeline',
      stageLatencies: {
        semanticMs: semanticResult.latencyMs,
        semanticRepairMs,
        normalizeMs: normalizeResult.latencyMs,
        normalizeRepairMs
      },
      finishReasons: {
        semantic: semanticResult.finishReason,
        normalize: normalizeResult.finishReason
      },
      semanticCandidates: semanticVariables.length
    }
  };
}

function runHeuristicAnalyze({ clippedText, reason, mode = 'heuristic_fallback' }) {
  const shaped = { variables: fallbackExtractVariablesFromText(clippedText) };
  validateOutputShape(shaped);
  const checked = enforcePluginCompatibility(shaped);
  return {
    parsed: checked.parsed,
    warnings: checked.warnings,
    meta: {
      mode,
      stageLatencies: {
        heuristicMs: 0
      },
      fallbackReason: reason || ''
    }
  };
}

async function callKimiAnalyze({ text }) {
  if (!KIMI_API_KEY) {
    throw new Error('Missing KIMI_API_KEY (or MOONSHOT_API_KEY)');
  }

  const skillPrompt = readSkillPrompt();
  const clippedText = String(text || '').slice(0, MAX_INPUT_CHARS);

  if (!PIPELINE_ENABLED) {
    try {
      return await runDirectAnalyze({ skillPrompt, clippedText });
    } catch (directErr) {
      console.warn(`[Kimi Bridge] direct failed, fallback to heuristic: ${directErr.message}`);
      return runHeuristicAnalyze({ clippedText, reason: directErr.message, mode: 'direct_fallback_heuristic' });
    }
  }

  try {
    return await runPipelineAnalyze({ skillPrompt, clippedText });
  } catch (pipelineErr) {
    if (!PIPELINE_FALLBACK_DIRECT) {
      console.warn(`[Kimi Bridge] pipeline failed, fallback to heuristic: ${pipelineErr.message}`);
      return runHeuristicAnalyze({ clippedText, reason: pipelineErr.message, mode: 'pipeline_fallback_heuristic' });
    }

    console.warn(`[Kimi Bridge] pipeline failed, fallback to direct: ${pipelineErr.message}`);
    try {
      const direct = await runDirectAnalyze({ skillPrompt, clippedText });
      return {
        ...direct,
        meta: {
          ...direct.meta,
          mode: 'pipeline_fallback_direct',
          pipelineError: pipelineErr.message
        }
      };
    } catch (directErr) {
      console.warn(`[Kimi Bridge] direct fallback failed, fallback to heuristic: ${directErr.message}`);
      return runHeuristicAnalyze({
        clippedText,
        reason: `pipeline: ${pipelineErr.message}; direct: ${directErr.message}`,
        mode: 'pipeline_direct_fallback_heuristic'
      });
    }
  }
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';

    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 5 * 1024 * 1024) {
        reject(new Error('Payload too large'));
      }
    });

    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(new Error(`Invalid request JSON: ${err.message}`));
      }
    });

    req.on('error', reject);
  });
}

function buildHealthPayload() {
  const uptimeSec = Math.floor((Date.now() - bridgeMetrics.startedAt) / 1000);
  const total = bridgeMetrics.analyzeSuccess + bridgeMetrics.analyzeFailure;
  const successRate = total > 0 ? Number((bridgeMetrics.analyzeSuccess / total).toFixed(4)) : 1;

  return {
    ok: true,
    provider: 'kimi',
    model: STRUCT_MODEL,
    temperature: STRUCT_TEMPERATURE,
    requestTimeoutMs: STRUCT_TIMEOUT_MS,
    maxTokens: STRUCT_MAX_TOKENS,
    skillPromptPath: SKILL_PROMPT_PATH,
    pipeline: {
      enabled: PIPELINE_ENABLED,
      fallbackDirect: PIPELINE_FALLBACK_DIRECT,
      networkRetries: NETWORK_RETRIES,
      semantic: {
        model: SEMANTIC_MODEL,
        temperature: SEMANTIC_TEMPERATURE,
        timeoutMs: SEMANTIC_TIMEOUT_MS,
        maxTokens: SEMANTIC_MAX_TOKENS
      },
      normalize: {
        model: STRUCT_MODEL,
        temperature: STRUCT_TEMPERATURE,
        timeoutMs: STRUCT_TIMEOUT_MS,
        maxTokens: STRUCT_MAX_TOKENS
      },
      direct: {
        model: LEGACY_MODEL,
        temperature: LEGACY_TEMPERATURE,
        timeoutMs: LEGACY_TIMEOUT_MS,
        maxTokens: LEGACY_MAX_TOKENS
      }
    },
    metrics: {
      ...bridgeMetrics,
      uptimeSec,
      successRate
    }
  };
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') {
    setCors(res);
    res.statusCode = 204;
    res.end();
    return;
  }

  if (req.method === 'GET' && req.url === '/health') {
    sendJson(res, 200, buildHealthPayload());
    return;
  }

  if (req.method === 'GET' && req.url === '/stats') {
    const health = buildHealthPayload();
    sendJson(res, 200, {
      ok: true,
      provider: health.provider,
      model: health.model,
      pipeline: health.pipeline,
      metrics: health.metrics
    });
    return;
  }

  // POST /format: AI 排版指令生成
  if (req.method === 'POST' && req.url === '/format') {
    try {
      const payload = await readJsonBody(req);
      const text = payload && typeof payload.text === 'string' ? payload.text : '';
      const context = payload && payload.context ? payload.context : {};
      // 接收会话上下文（任务A新增）
      const chatContext = payload && Array.isArray(payload.chatContext) ? payload.chatContext : [];

      if (!text.trim()) {
        sendJson(res, 400, { error: 'text is required' });
        return;
      }

      if (!KIMI_API_KEY) {
        sendJson(res, 500, { error: 'Missing KIMI_API_KEY' });
        return;
      }

      const formatSkillPath = path.join(__dirname, 'formatting-skill', 'SKILL.md');
      let formatSkillPrompt;
      try {
        formatSkillPrompt = fs.readFileSync(formatSkillPath, 'utf8');
      } catch (readErr) {
        sendJson(res, 500, { error: `Cannot read formatting skill: ${readErr.message}` });
        return;
      }

      // 构建消息数组（执行指令不需要会话历史）
      const messages = [
        { role: 'system', content: formatSkillPrompt },
        {
          role: 'user',
          content: [
            `文档上下文: ${JSON.stringify(context)}`,
            `用户需求: ${text}`
          ].join('\n\n')
        }
      ];

      console.log(`[/format] executing without chatContext, messages: ${messages.length}`);

      const chatResult = await callChatCompletion({
        messages: messages,
        model: FORMAT_MODEL,
        temperature: STRUCT_TEMPERATURE,
        maxTokens: 4096,
        timeoutMs: FORMAT_TIMEOUT_MS,
        stageLabel: 'Kimi format-stage'
      });

      setCors(res);
      res.statusCode = 200;
      res.setHeader('Content-Type', 'text/plain; charset=utf-8');
      res.end(chatResult.contentText);
    } catch (err) {
      sendJson(res, 500, { error: err.message });
    }
    return;
  }

  // POST /agent/dispatch: 意图路由（任务B新增）
  if (req.method === 'POST' && req.url === '/agent/dispatch') {
    try {
      const payload = await readJsonBody(req);
      const text = payload && typeof payload.text === 'string' ? payload.text : '';
      const chatContext = payload && Array.isArray(payload.chatContext) ? payload.chatContext : [];

      if (!text.trim()) {
        sendJson(res, 400, { error: 'text is required' });
        return;
      }

      // 构建消息，使用轻量级模型进行意图识别
      const dispatchSystemPrompt = `你是一个意图分类器。请分析用户输入，识别其意图类型。

意图类型：
1. FORMAT - 格式修改（如：字体、段落、表格、页面样式、论文格式如APA/MLA等排版操作）
2. CONTENT_EDIT - 基础文本编辑（如：插入、删除、替换、添加具体文本内容）
3. GENERAL_QA - 问答回复（不涉及文档修改，仅回答问题）

判断规则：
- "改成XXX格式"、"排版成XXX"、"设置成XXX"、"应用XXX格式" → FORMAT
- "改成APA"、"设置APA格式" → FORMAT（这是格式需求）
- "插入XXX文字"、"删除XXX内容"、"替换XXX为YYY" → CONTENT_EDIT
- "什么是XXX"、"如何做XXX"、"解释一下" → GENERAL_QA

请直接返回JSON格式：
{"intent": "意图类型", "confidence": 0.0-1.0, "reason": "简短原因"}

注意：只要涉及文档格式、排版、样式，全部归类为 FORMAT。
`;

      const dispatchMessages = [
        { role: 'system', content: dispatchSystemPrompt },
        ...chatContext.slice(-6), // 只传最近6轮上下文
        { role: 'user', content: text }
      ];

      const dispatchResult = await callChatCompletion({
        messages: dispatchMessages,
        model: STRUCT_MODEL, // 使用结构化模型
        temperature: 0.3,
        maxTokens: 256,
        timeoutMs: 30000,
        stageLabel: 'intent-dispatch'
      });

      // 解析结果
      let intentResult = {
        intent: 'FORMAT',
        confidence: 0.5,
        reason: '模型响应解析失败，使用默认'
      };

      try {
        const parsed = JSON.parse(dispatchResult.contentText.trim());
        if (parsed.intent && ['FORMAT', 'CONTENT_EDIT', 'GENERAL_QA'].includes(parsed.intent)) {
          intentResult = {
            intent: parsed.intent,
            confidence: Math.min(1, Math.max(0, parsed.confidence || 0.5)),
            reason: parsed.reason || ''
          };
        }
      } catch (parseErr) {
        console.log('[/agent/dispatch] JSON parse failed, using default:', parseErr.message);
      }

      console.log(`[/agent/dispatch] intent=${intentResult.intent}, confidence=${intentResult.confidence}`);

      setCors(res);
      sendJson(res, 200, intentResult);
    } catch (err) {
      console.error('[/agent/dispatch] error:', err.message);
      // 出错时返回默认意图，不阻断流程
      setCors(res);
      sendJson(res, 200, {
        intent: 'FORMAT',
        confidence: 0.3,
        reason: '服务端错误: ' + err.message
      });
    }
    return;
  }

  if (req.method === 'POST' && req.url === '/analyze') {
    bridgeMetrics.requestsTotal += 1;
    const startedAt = Date.now();

    try {
      const payload = await readJsonBody(req);
      const text = payload && typeof payload.text === 'string' ? payload.text : '';

      if (!text.trim()) {
        bridgeMetrics.analyzeFailure += 1;
        bridgeMetrics.lastAnalyzeAt = Date.now();
        bridgeMetrics.lastLatencyMs = Date.now() - startedAt;
        bridgeMetrics.lastError = 'text is required';
        sendJson(res, 400, { error: 'text is required' });
        return;
      }

      const result = await callKimiAnalyze({ text });

      bridgeMetrics.analyzeSuccess += 1;
      bridgeMetrics.lastAnalyzeAt = Date.now();
      bridgeMetrics.lastAnalyzeIso = new Date().toISOString();
      bridgeMetrics.lastLatencyMs = Date.now() - startedAt;
      bridgeMetrics.lastVariablesCount = Array.isArray(result.parsed.variables) ? result.parsed.variables.length : 0;
      bridgeMetrics.lastVariableTags = Array.isArray(result.parsed.variables)
        ? result.parsed.variables
          .map((item) => (item && typeof item.tag === 'string' ? item.tag.trim() : ''))
          .filter(Boolean)
          .slice(0, 8)
        : [];
      bridgeMetrics.lastSchemaWarnings = Array.isArray(result.warnings) ? result.warnings.length : 0;
      bridgeMetrics.lastPipelineMode = result.meta && result.meta.mode ? result.meta.mode : null;
      bridgeMetrics.lastStageLatencies = result.meta && result.meta.stageLatencies ? result.meta.stageLatencies : null;
      bridgeMetrics.lastError = '';

      sendJson(res, 200, result.parsed);
    } catch (err) {
      bridgeMetrics.analyzeFailure += 1;
      bridgeMetrics.lastAnalyzeAt = Date.now();
      bridgeMetrics.lastAnalyzeIso = new Date().toISOString();
      bridgeMetrics.lastLatencyMs = Date.now() - startedAt;
      bridgeMetrics.lastError = err.message;
      sendJson(res, 500, { error: err.message });
    }
    return;
  }

  sendJson(res, 404, { error: 'Not found' });
});

if (require.main === module) {
  server.on('error', (err) => {
    console.error(`[Kimi Bridge] startup failed: ${err.message}`);
    process.exit(1);
  });

  const BIND_HOST = sanitizeEnv(process.env.BIND_HOST) || '0.0.0.0';
  server.listen(PORT, BIND_HOST, () => {
    console.log(`[Kimi Bridge] listening on http://${BIND_HOST}:${PORT}`);
    console.log(`[Kimi Bridge] pipeline=${PIPELINE_ENABLED ? 'on' : 'off'}, semantic=${SEMANTIC_MODEL}, normalize=${STRUCT_MODEL}`);
    console.log(`[Kimi Bridge] skill=${SKILL_PROMPT_PATH}`);
  });
}

module.exports = {
  server,
  buildHealthPayload
};
