# Judge Fast Deploy (One-time Install + Auto Cloud Sync)

Goal:
- Judges install plugin once.
- Later updates are pulled automatically from cloud.
- AI service is always online from cloud.

## 1) Deploy cloud endpoints

You need two stable HTTPS URLs:
- App URL (taskpane frontend), e.g. `https://your-app.vercel.app/taskpane.html`
- API URL (AI gateway), e.g. `https://your-api.onrender.com`

Required API endpoints:
- `GET /health`
- `POST /analyze`
- `POST /format`

## 2) Publish manifest

Host a manifest JSON at a stable URL, for example:
`https://your-app.vercel.app/manifest.json`

Template: `cloud/manifest.example.json`

Key fields:
- `taskpaneUrl`: remote taskpane URL
- `apiBaseUrl`: remote AI API base URL
- `autoRefreshMs`: config refresh interval in ms

## 3) Configure bootstrap in plugin (one-time before packaging)

Option A (recommended):

```bash
CLOUD_MANIFEST_URL="https://your-app.vercel.app/manifest.json" \
CLOUD_TASKPANE_URL="" \
npm run cloud:bootstrap:set
```

Option B:
- Edit `main.js` and set `PEVC_DEFAULT_BOOTSTRAP.manifestUrl`.

## 4) Build and send to judges (only once)

```bash
./build_update.command
```

Send the generated zip in `dist_updates/`.
Judges run `Install_PEVC_WPS.command` once.

## 5) How updates work afterwards

- You update cloud frontend/API/skills.
- You update `manifest.json` if needed.
- Judges open plugin and it resolves remote taskpane URL from manifest.
- Runtime config auto-refreshes and switches API endpoint automatically.

No re-install needed.

## 6) Keepalive (must for evaluation)

- Run API with process manager (`pm2`/platform auto-restart).
- Add external health probe to `GET /health` every 1 minute.
- Enable alerts (email/IM) on downtime.

## 7) Quick verification checklist

1. On judge machine install once.
2. Open plugin, check console has remote URL in `ShowDialog URL`.
3. In plugin, AI status should be online from cloud API.
4. Change one text in cloud frontend and refresh plugin: new content appears.
