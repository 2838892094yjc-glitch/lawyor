# 合同变量识别 Skill（WPS 版）

## 概述

这是当前项目里“普通合同 → 可填写模板”的核心识别能力说明文档。
它通过本地 AI 服务（默认 `http://127.0.0.1:8765/analyze`）识别变量，再由 WPS 适配层完成埋点。

核心目标：
1. 从合同正文中抽取可变量化内容
2. 生成结构化 JSON（`variables`）
3. 转成表单字段并批量埋点到 WPS 文档

---

## 当前实现架构

识别与落地链路：

1. `taskpane.js` 调用 `window.wpsAdapter.getDocumentText()` 读取全文
2. `ai-skill.js` 分片并调用本地 AI 接口
3. `ai-parser.js` 校验并转换 AI 输出
4. `taskpane.js` 合并配置，渲染动态表单
5. `wps_adapter.js` 按锚点搜索并插入 Content Control

相关文件：

- `contract-variable-skill/Skill.md`：提示词规范（历史命名）
- `ai-skill.js`：AI 调用、分片、去重
- `ai-parser.js`：输出校验与字段转换
- `formatters.js`：格式化函数白名单
- `taskpane.js`：AI 识别入口与批量埋点
- `wps_adapter.js`：WPS 文档读写与埋点

## 当前使用的 API（你现在的配置）

- 前端插件调用的是本地桥接接口：`http://127.0.0.1:8765/analyze`（见 `ai-skill.js`）
- 也就是说，实际用哪个大模型由你本地 `8765` 服务决定，而不是由前端固定死

## 切换到 Kimi

1. 配置环境变量（示例见仓库 `.env.kimi.example`）
2. 启动本地 Kimi 桥接服务（会自动做自检）：

```bash
cd /Users/yangjingchi/Desktop/pevc-wps-addin
export KIMI_API_KEY="你的key"
export KIMI_MODEL="moonshot-v1-32k"
npm run ai:kimi
```

3. 打开健康检查：`http://127.0.0.1:8765/health`
4. 在 WPS 插件中点击“AI 智能识别”即可

可单独执行自检：

```bash
npm run ai:kimi:test
```

如需“开机自动启动桥接服务”（无感启动）：

```bash
# 先在当前终端 export KIMI_API_KEY / KIMI_MODEL
npm run ai:kimi:agent:install

# 查看状态
npm run ai:kimi:agent:status

# 卸载自动启动
npm run ai:kimi:agent:uninstall
```

插件侧会显示 AI 服务面板（在线状态、模型、成功/失败计数、最近耗时、最近错误）。

---

## 输出协议（AI 必须返回）

```json
{
  "variables": [
    {
      "context": "前后文锚点",
      "prefix": "变量前固定文本",
      "placeholder": "变量本体",
      "suffix": "变量后固定文本",
      "label": "中文字段名",
      "tag": "PinYinTag",
      "type": "text|number|date|select|radio|textarea",
      "options": ["可选项"],
      "formatFn": "none|dateUnderline|...",
      "mode": "insert|paragraph",
      "layer": 1,
      "confidence": "high|medium|low",
      "reason": "识别依据"
    }
  ]
}
```

---

## 快速验证（WPS 环境）

在 WPS 打开插件后，可在控制台执行：

```javascript
async function testRecognizeOnce() {
  const text = await window.wpsAdapter.getDocumentText();
  const aiOutput = await window.AISkill.analyzeDocument(text);
  const validation = window.AIParser.validateAIOutput(aiOutput);

  console.log('变量数量:', aiOutput?.variables?.length || 0);
  console.log('验证结果:', validation);

  if (validation.valid) {
    const parsed = window.AIParser.parseAIOutput(aiOutput);
    console.log('解析结果:', parsed.success, parsed.stats);
  }
}

testRecognizeOnce();
```

---

## 批量埋点策略（当前版本）

批量埋点会按候选锚点顺序查找（从更长、更唯一到更短）：

1. `context`
2. `prefix + placeholder + suffix`
3. `placeholder`

这比仅使用 `placeholder` 更稳健，能降低重复占位符导致的误定位概率。

---

## API 参考

### `window.AISkill`

```javascript
const aiOutput = await window.AISkill.analyzeDocument(documentText, onProgress);
const chunkOutput = await window.AISkill.callAIWithSkill(text, 0, 1);
const promptText = await window.AISkill.loadSkillPrompt();
```

### `window.AIParser`

```javascript
const check = window.AIParser.validateAIOutput(aiOutput);
const parsed = window.AIParser.parseAIOutput(aiOutput);
const embedInfo = window.AIParser.generateEmbedInfo(aiOutput);
```

### `window.Formatters`

```javascript
const value = window.Formatters.applyFormat(input, formatFn);
const ok = window.Formatters.isValidFormatFn(formatFn);
const all = window.Formatters.getValidFormatFns();
```

---

## 已知优化方向

1. 继续增强多候选锚点命中率（表格、换行、全角符号场景）
2. 增加识别质量回传日志（按合同类型统计）
3. 细化 layer2/layer3 的负样本规则，降低误识别
