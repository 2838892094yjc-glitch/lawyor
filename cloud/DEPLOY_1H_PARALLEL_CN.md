# 1 小时并行上线方案（评委一次安装，后续自动更新）

目标：
- 评委只安装一次插件。
- 你后续改前端/AI 服务后，评委端自动拉取最新版本。
- 服务持续在线，尽量不下线。

## 并行分工（可 2~3 人同时做）

- 轨道 A（前端托管，10~20 分钟）
  - 把插件前端静态文件部署到 Vercel（或 Netlify）。
  - 拿到 `taskpane.html` 的 HTTPS 地址。

- 轨道 B（AI 服务托管，15~30 分钟）
  - 把 `kimi-analyze-server.js` 部署到 Render/Railway/Fly.io。
  - 确保接口可用：`GET /health`、`POST /analyze`、`POST /format`。

- 轨道 C（引导配置 + 打包，10 分钟）
  - 生成云端 `manifest.json`。
  - 将 `main.js` 的 bootstrap 指向 manifest。
  - 打包并发给评委（仅一次）。

## 不需要域名吗？

- 不需要。先用平台自带域名即可：
  - `https://xxx.vercel.app`
  - `https://xxx.onrender.com`
- 评审阶段最省时间的做法就是先不上自定义域名。

## 第 1 步：生成并发布 manifest

在本地执行（把 URL 换成你的）：

```bash
CLOUD_TASKPANE_URL="https://your-app.vercel.app/taskpane.html" \
CLOUD_API_BASE_URL="https://your-api.onrender.com" \
CLOUD_MANIFEST_VERSION="2026.02.14.1" \
npm run cloud:manifest:gen
```

会生成：`cloud/manifest.json`

把 `cloud/manifest.json` 上传到固定 HTTPS 地址，例如：
`https://your-app.vercel.app/manifest.json`

## 第 2 步：把插件 bootstrap 指到 manifest

```bash
CLOUD_MANIFEST_URL="https://your-app.vercel.app/manifest.json" \
npm run cloud:bootstrap:set
```

> 这一步只做一次，打给评委的安装包里就带上这个地址。

## 第 3 步：打包并发评委

```bash
./build_update.command
```

把 `dist_updates/` 里的最新包发评委，让他们执行一次安装脚本。

## 后续更新方式（无需评委重装）

- 你改前端：重新部署前端。
- 你改 AI 服务：重新部署 API。
- 如有地址/版本变更：更新 `manifest.json`。
- 评委下次打开插件会自动读取 manifest，并切换到最新云端配置。

## 保活建议（评审期间强烈建议）

- Render/Railway 开启自动重启（默认有）。
- 使用 UptimeRobot/Better Stack 每 1 分钟探测一次：
  - `GET https://your-api.onrender.com/health`
- 设置告警（邮件/飞书/钉钉），一旦挂了立即重启。

## 5 分钟验收清单

1. 评委机安装一次后能打开插件。
2. 控制台看到 `ShowDialog URL` 指向云端地址。
3. 控制台看到 `[RuntimeConfig] synced` 且 `api=` 为云端 API。
4. 改云端 taskpane 文案，评委刷新插件能看到变化。
5. 改 AI 服务逻辑，评委再次调用时生效。
