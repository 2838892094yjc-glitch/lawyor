#!/usr/bin/env node

const fs = require('fs');

const VALID_TYPES = new Set(['text', 'number', 'date', 'select', 'radio', 'textarea']);
const VALID_FORMATS = new Set([
  'none',
  'dateUnderline',
  'dateYearMonth',
  'chineseNumber',
  'chineseNumberWan',
  'amountWithChinese',
  'articleNumber',
  'percentageChinese',
]);
const VALID_MODES = new Set(['insert', 'paragraph']);
const REQUIRED_FIELDS = ['context', 'prefix', 'placeholder', 'suffix', 'label', 'tag', 'type', 'formatFn', 'mode'];

function fail(msg) {
  console.error(`ERROR: ${msg}`);
  process.exitCode = 1;
}

function main() {
  const file = process.argv[2];
  if (!file) {
    console.error('Usage: node scripts/validate_output.js <output.json>');
    process.exit(2);
  }

  let parsed;
  try {
    parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (err) {
    console.error(`Failed to parse JSON: ${err.message}`);
    process.exit(2);
  }

  if (!parsed || typeof parsed !== 'object' || !Array.isArray(parsed.variables)) {
    fail('Root object must contain variables array');
    return;
  }

  const seenTags = new Set();
  parsed.variables.forEach((v, i) => {
    const name = `variables[${i}]`;

    for (const key of REQUIRED_FIELDS) {
      if (!(key in v)) fail(`${name} missing required field: ${key}`);
    }

    if (typeof v.tag === 'string' && seenTags.has(v.tag)) {
      fail(`${name} duplicate tag: ${v.tag}`);
    }
    if (typeof v.tag === 'string') seenTags.add(v.tag);

    if (v.type && !VALID_TYPES.has(v.type)) fail(`${name} invalid type: ${v.type}`);
    if (v.formatFn && !VALID_FORMATS.has(v.formatFn)) fail(`${name} invalid formatFn: ${v.formatFn}`);
    if (v.mode && !VALID_MODES.has(v.mode)) fail(`${name} invalid mode: ${v.mode}`);

    if ((v.type === 'select' || v.type === 'radio') && (!Array.isArray(v.options) || v.options.length === 0)) {
      fail(`${name} requires non-empty options for type=${v.type}`);
    }
  });

  if (process.exitCode) return;
  console.log(`OK: ${parsed.variables.length} variables validated`);
}

main();
