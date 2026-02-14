#!/usr/bin/env node

/**
 * Configure cloud bootstrap URLs used by main.js runtime loader.
 * Usage:
 *   CLOUD_MANIFEST_URL=https://... CLOUD_TASKPANE_URL=https://... node scripts/set-cloud-bootstrap.js
 */

const fs = require('fs');
const path = require('path');

const mainPath = path.join(__dirname, '..', 'main.js');
const manifestUrl = (process.env.CLOUD_MANIFEST_URL || '').trim();
const taskpaneUrl = (process.env.CLOUD_TASKPANE_URL || '').trim();

if (!manifestUrl && !taskpaneUrl) {
  console.error('Please set CLOUD_MANIFEST_URL and/or CLOUD_TASKPANE_URL');
  process.exit(1);
}

let content = fs.readFileSync(mainPath, 'utf8');

if (manifestUrl) {
  content = content.replace(/manifestUrl:\s*"[^"]*"/, `manifestUrl: "${manifestUrl}"`);
}
if (taskpaneUrl) {
  content = content.replace(/taskpaneUrl:\s*"[^"]*"/, `taskpaneUrl: "${taskpaneUrl}"`);
}

fs.writeFileSync(mainPath, content);
console.log('[set-cloud-bootstrap] Updated main.js');
if (manifestUrl) console.log(' manifestUrl =', manifestUrl);
if (taskpaneUrl) console.log(' taskpaneUrl =', taskpaneUrl);
