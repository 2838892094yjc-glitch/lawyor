# 无需 Render：仅用 Vercel 部署（无信用卡版本）

适用场景：Render 要求支付信息，无法继续。

结果：
- 前端 + API 都在 Vercel。
- 评委只安装一次，后续自动从云端更新。

## 1) 在 Vercel 导入 GitHub 仓库

仓库：`https://github.com/2838892094yjc-glitch/lawyor`

创建项目时只需要设置 1 个环境变量：
- `KIMI_API_KEY` = 你的 Moonshot/Kimi Key

> 不需要 Render，不需要单独后端平台。

## 2) 部署完成后拿到一个域名

假设是：`https://lawyor-demo.vercel.app`

那么：
- taskpaneUrl = `https://lawyor-demo.vercel.app/taskpane.html`
- apiBaseUrl = `https://lawyor-demo.vercel.app/api`
- manifestUrl = `https://cdn.jsdelivr.net/gh/2838892094yjc-glitch/lawyor@main/cloud/manifest.json`

## 3) 回到本地，更新 manifest 里的真实云地址

```bash
CLOUD_TASKPANE_URL="https://lawyor-demo.vercel.app/taskpane.html" \
CLOUD_API_BASE_URL="https://lawyor-demo.vercel.app/api" \
CLOUD_MANIFEST_VERSION="2026.02.14.2" \
npm run cloud:manifest:gen
```

然后提交并推送 `cloud/manifest.json` 到 `main` 分支。

## 4) 评委发包（一次）

```bash
./build_update.command
```

发 `dist_updates/` 的安装包，评委装一次即可。

## 5) 后续更新

- 改代码并 push 到 GitHub。
- Vercel 自动部署。
- 如有接口/版本变化，更新 `cloud/manifest.json` 并 push。
- 评委端自动拉新，不用重装。

