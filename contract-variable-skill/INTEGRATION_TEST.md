# 合同变量识别 Skill - 集成测试指南（WPS）

## 测试目标

验证“识别 → 解析 → 埋点 → 填写回写”全链路可用：

1. AI 返回结构化 `variables`
2. Parser 校验并转换成功
3. 批量埋点可在文档中创建 Content Control
4. 表单填写可回写到文档

---

## 环境准备

### 1) 启动 WPS 加载项调试服务

```bash
cd /Users/yangjingchi/Desktop/pevc-wps-addin
wpsjs debug
```

### 2) 确认本地 AI 服务可用

默认接口：`http://127.0.0.1:8765/analyze`

若使用 Kimi，可在另一个终端启动：

```bash
cd /Users/yangjingchi/Desktop/pevc-wps-addin
export KIMI_API_KEY="你的key"
export KIMI_MODEL="moonshot-v1-32k"
npm run ai:kimi
```

说明：`npm run ai:kimi` 会在启动时自动执行一次 `ai:kimi:test` 自检。

插件内可在“AI 服务状态”面板查看在线状态、模型、成功/失败统计与最近错误。

若不可用，`AI 智能识别` 会直接报错。

### 3) 打开测试合同

- 任一真实合同（建议包含显式占位符 + 固定文本单位 + 可选条款）

---

## 测试 1：识别输出

在插件控制台执行：

```javascript
async function testAIRecognize() {
  const docText = await window.wpsAdapter.getDocumentText();
  console.log('文档长度:', docText.length);

  const aiOutput = await window.AISkill.analyzeDocument(docText);
  console.log('识别结果:', aiOutput);

  const validation = window.AIParser.validateAIOutput(aiOutput);
  console.log('验证:', validation);

  window.__test_ai_output = aiOutput;
}

testAIRecognize();
```

预期：
- `aiOutput.variables` 非空
- `validation.valid === true`
- 字段具备 `context/prefix/placeholder/suffix/label/tag/type/formatFn/mode`

---

## 测试 2：解析转换

```javascript
const parsed = window.AIParser.parseAIOutput(window.__test_ai_output);
console.log(parsed.success, parsed.stats, parsed.warnings);
window.__test_parsed = parsed;
```

预期：
- `parsed.success === true`
- `parsed.config` 可用于动态表单渲染
- `stats.total > 0`

---

## 测试 3：批量埋点

方式 A（推荐）：点击 UI 的“📌 批量埋点”按钮。  
方式 B：控制台调用：

```javascript
batchEmbedAIFields();
```

预期：
- 成功数 > 0
- 失败项可在 Console 里看到原因
- 文档中出现对应 `ContentControl`（可手动点选检查）

重点检查：
- 固定文本未被破坏（如“万元”“个工作日”仍保留）
- 重复占位符场景，优先命中更长锚点

---

## 测试 4：表单回写

1. 在右侧表单填写几个字段
2. 观察文档对应控件文本是否同步更新
3. 关闭并重新打开文档后，检查状态是否可恢复

预期：
- 回写成功
- 文档属性状态保存成功（`DocumentStateManager`）

---

## 测试 5：回归项（本次改动相关）

1. `window.AISkill.loadSkillPrompt()` 可返回提示词文本
2. `insertControlBySearch` 支持候选数组输入
3. 无候选时给出“搜索文本为空”错误

快速校验：

```javascript
(async () => {
  const prompt = await window.AISkill.loadSkillPrompt();
  console.log('prompt length:', prompt.length);
})();
```

---

## 常见问题

### 1) AI 返回空结果
- 检查本地 AI 服务是否启动
- 检查文档是否有可识别文本

### 2) 埋点失败较多
- 多为锚点不唯一或文本差异（全角/换行/标点）
- 优先检查失败字段的 `context` 和 `prefix+placeholder+suffix`

### 3) 解析报错
- 查看 `validation.errors`
- 常见原因：`type` 非白名单、`select/radio` 缺少 `options`
