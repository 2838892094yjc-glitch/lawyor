#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const taskpaneUrl = (process.env.CLOUD_TASKPANE_URL || '').trim();
const apiBaseUrl = (process.env.CLOUD_API_BASE_URL || '').trim();
const autoRefreshMs = Number(process.env.CLOUD_AUTO_REFRESH_MS || 300000);
const version = (process.env.CLOUD_MANIFEST_VERSION || new Date().toISOString().replace(/[-:TZ.]/g, '').slice(0, 14)).trim();
const outputPath = process.env.CLOUD_MANIFEST_OUT || path.join('cloud', 'manifest.json');

if (!taskpaneUrl || !apiBaseUrl) {
  console.error('Missing required env vars: CLOUD_TASKPANE_URL and CLOUD_API_BASE_URL');
  process.exit(1);
}

const manifest = {
  version,
  taskpaneUrl,
  apiBaseUrl,
  autoRefreshMs: Number.isFinite(autoRefreshMs) && autoRefreshMs >= 30000 ? autoRefreshMs : 300000,
  client: {
    taskpaneUrl,
    autoRefreshMs: Number.isFinite(autoRefreshMs) && autoRefreshMs >= 30000 ? autoRefreshMs : 300000
  },
  services: {
    apiBaseUrl
  }
};

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, JSON.stringify(manifest, null, 2));

console.log('[generate-cloud-manifest] Wrote', outputPath);
console.log(' version    =', manifest.version);
console.log(' taskpane   =', manifest.taskpaneUrl);
console.log(' apiBaseUrl =', manifest.apiBaseUrl);
console.log(' autoRefreshMs =', manifest.autoRefreshMs);
