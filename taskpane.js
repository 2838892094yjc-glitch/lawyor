/**
 * PEVC合同助手 - WPS版 (完整表单管理 v20260129)
 * 100% 文档驱动，支持 WPS 环境隔离
 * 新增：完整的表单管理功能（添加/编辑/删除字段）
 */

/* ==================================================================
 * 1. 基础配置与全局状态
 * ================================================================== */

const CURRENT_CONFIG_VERSION = "v20260129_V12_AIProgress";

let contractConfig = [];
let pendingFields = [];
let savedFormData = {};
let enabledShareholders = {};
let enabledInvestors = {};
let activeDocKey = "";
let docWatchTimer = null;
let isDocRefreshing = false;
let docWatchInFlight = false;
let aiStatusTimer = null;
const AUTO_EMBED_STORAGE_KEY = "pevc_auto_embed_after_ai";
let autoEmbedAfterAI = true;

function getRuntimeApiUrl(path, fallbackUrl) {
    try {
        if (window.PEVCRuntimeConfig && window.PEVCRuntimeConfig.buildApiUrl) {
            return window.PEVCRuntimeConfig.buildApiUrl(path);
        }
    } catch (_err) {}
    return fallbackUrl;
}

async function syncRuntimeConfigFromManifest() {
    try {
        if (!window.PEVCRuntimeConfig || !window.PEVCRuntimeConfig.loadRemoteConfig) return;
        const result = await window.PEVCRuntimeConfig.loadRemoteConfig({ timeoutMs: 2500 });
        if (result && result.ok) {
            const cfg = result.config || {};
            console.log(`[RuntimeConfig] synced: version=${cfg.version || '-'}, api=${cfg.apiBaseUrl || '-'}`);
            window.PEVCRuntimeConfig.startAutoSync && window.PEVCRuntimeConfig.startAutoSync();
        } else if (result && result.reason && result.reason !== 'manifest_url_empty') {
            console.warn('[RuntimeConfig] sync skipped:', result.reason);
        }
    } catch (err) {
        console.warn('[RuntimeConfig] sync failed:', err && err.message ? err.message : err);
    }
}

// 硬编码的基础配置 (PEVC 默认表单)
const DEFAULT_CONTRACT_CONFIG = [
    {
        id: "section_files",
        header: { label: "1. 所需文件", tag: "Section_Files" },
        fields: [ { type: "html_placeholder", targetId: "local-sync-section" } ]
    },
    {
        id: "section_company_info",
        header: { label: "2. 公司基本信息", tag: "Section_CompanyInfo" },
        fields: [
            { id: "signingDate", label: "签订时间", tag: "SigningDate", type: "date", formatFn: "dateUnderline", placeholder: "选择日期", hasParagraphToggle: true },
            { id: "signingPlace", label: "签订地点", tag: "SigningPlace", type: "text", placeholder: "如：北京" },
            { id: "lawyerRep", label: "律师代表", tag: "LawyerRepresenting", type: "radio", options: ["公司", "投资方", "公司/投资方"] },
            { id: "projectShortName", label: "项目简称", tag: "ProjectShortName", type: "text" },
            { id: "companyName", label: "目标公司名称", tag: "CompanyName", type: "text" }
        ]
    }
];

/* ==================================================================
 * 2. 文档状态管理器
 * ================================================================== */

const DocumentStateManager = {
    _isSaving: false,
    _pendingSave: false,
    _lastSavedFingerprint: "",

    _getFingerprint(state) {
        return JSON.stringify(state).length + "";
    },

    async load() {
        console.log("[StateMgr] Loading state from document...");
        if (!window.wpsAdapter) return null;
        const state = await window.wpsAdapter.loadAddinState();
        if (state) {
            this._lastSavedFingerprint = this._getFingerprint(state);
            return state;
        }
        return null;
    },

    async save(force = false) {
        if (this._isSaving) { this._pendingSave = true; return; }
        const state = {
            config: contractConfig,
            pendingFields: pendingFields,
            formData: collectForm(),
            rounds: { enabledShareholders, enabledInvestors },
            version: CURRENT_CONFIG_VERSION,
            timestamp: Date.now()
        };
        const fp = this._getFingerprint(state);
        if (!force && fp === this._lastSavedFingerprint) return;

        this._isSaving = true;
        try {
            await window.wpsAdapter.saveAddinState(state);
            this._lastSavedFingerprint = fp;
            console.log("[StateMgr] ✅ Saved");
        } finally {
            this._isSaving = false;
            if (this._pendingSave) { this._pendingSave = false; this.save(); }
        }
    },

    async clear() {
        if (!window.wpsAdapter) return;
        await window.wpsAdapter.clearAddinState();
        window.location.reload();
    }
};

const scheduleSmartSave = debounce(() => DocumentStateManager.save(), 2000);

/* ==================================================================
 * 3. 业务逻辑
 * ================================================================== */

async function loadFormConfig() {
    contractConfig = JSON.parse(JSON.stringify(DEFAULT_CONTRACT_CONFIG));
}

async function loadPEVCTemplate() {
    if (typeof PEVC_DEFAULT_TEMPLATE === 'undefined') {
        showNotification("模板未加载", "error");
        return;
    }
    contractConfig = JSON.parse(JSON.stringify(PEVC_DEFAULT_TEMPLATE));
    contractConfig.forEach(s => { if(s.fields) s.fields.forEach(f => f.isCustom = true); });
    await DocumentStateManager.save(true);
    buildForm();
}

async function syncFormToSelectedDocuments() {
    const formData = collectForm();
    if (!formData || Object.keys(formData).length === 0) {
        showNotification("没有可同步的数据", "warning");
        return;
    }
    
    if (!window.wpsAdapter || !window.wpsAdapter.pickTargetDocuments) {
        showNotification("WPS 适配器未就绪", "error");
        return;
    }
    
    try {
        const targets = await window.wpsAdapter.pickTargetDocuments();
        if (!targets || targets.length === 0) return;
        
        showNotification(`正在向 ${targets.length} 个文档同步数据...`, "info");
        const labelMap = buildLabelMap();
        const results = await window.wpsAdapter.syncFormDataToDocuments(formData, labelMap, targets);
        
        let msg = `同步完成！成功: ${results.successCount}/${targets.length}`;
        if (results.missingCount > 0) msg += `，跳过缺失字段: ${results.missingCount} 个`;
        showNotification(msg, results.successCount > 0 ? "success" : "warning", 5000);
    } catch (err) {
        console.error("[Sync] Error:", err);
        showNotification(`同步失败: ${err.message}`, "error");
    }
}

async function refreshAIServiceStatus(options = {}) {
    const { silent = true } = options;
    const statusEl = document.getElementById("ai-service-status");
    const providerEl = document.getElementById("ai-service-provider");
    const modelEl = document.getElementById("ai-service-model");
    const successEl = document.getElementById("ai-service-success");
    const failureEl = document.getElementById("ai-service-failure");
    const lastVarsEl = document.getElementById("ai-service-last-vars");
    const lastLatencyEl = document.getElementById("ai-service-last-latency");
    const lastModeEl = document.getElementById("ai-service-last-mode");
    const stageLatenciesEl = document.getElementById("ai-service-stage-latencies");
    const lastTagsEl = document.getElementById("ai-service-last-tags");
    const lastTimeEl = document.getElementById("ai-service-last-time");
    const lastErrorEl = document.getElementById("ai-service-last-error");

    if (!statusEl) return;

    try {
        if (!silent) {
            statusEl.textContent = "刷新中...";
            statusEl.style.color = "#2563eb";
        }
        const healthUrl = getRuntimeApiUrl("/health", "http://127.0.0.1:8765/health");
        const response = await fetch(healthUrl, { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const health = await response.json();
        const metrics = health.metrics || {};

        statusEl.textContent = "在线";
        statusEl.style.color = "#16a34a";
        if (providerEl) providerEl.textContent = health.provider || "-";
        if (modelEl) modelEl.textContent = health.model || "-";
        if (successEl) successEl.textContent = String(metrics.analyzeSuccess || 0);
        if (failureEl) failureEl.textContent = String(metrics.analyzeFailure || 0);
        if (lastVarsEl) lastVarsEl.textContent = metrics.lastVariablesCount !== undefined ? String(metrics.lastVariablesCount) : "-";
        if (lastLatencyEl) lastLatencyEl.textContent = metrics.lastLatencyMs !== null && metrics.lastLatencyMs !== undefined ? String(metrics.lastLatencyMs) : "-";
        if (lastModeEl) lastModeEl.textContent = metrics.lastPipelineMode || "-";
        if (stageLatenciesEl) {
            const latencies = metrics.lastStageLatencies && typeof metrics.lastStageLatencies === "object"
                ? Object.entries(metrics.lastStageLatencies)
                    .filter(([, value]) => value !== null && value !== undefined)
                    .map(([key, value]) => `${key}:${value}ms`)
                : [];
            stageLatenciesEl.textContent = latencies.length > 0 ? latencies.join(" | ") : "-";
        }
        if (lastTagsEl) {
            const tags = Array.isArray(metrics.lastVariableTags) ? metrics.lastVariableTags : [];
            lastTagsEl.textContent = tags.length > 0 ? tags.join(", ") : "-";
        }
        if (lastTimeEl) {
            const raw = metrics.lastAnalyzeIso || metrics.lastAnalyzeAt;
            if (!raw) {
                lastTimeEl.textContent = "-";
            } else {
                const date = typeof raw === "number" ? new Date(raw) : new Date(String(raw));
                lastTimeEl.textContent = Number.isNaN(date.getTime()) ? "-" : date.toLocaleString();
            }
        }
        if (lastErrorEl) lastErrorEl.textContent = metrics.lastError ? String(metrics.lastError).slice(0, 120) : "-";
    } catch (err) {
        statusEl.textContent = "离线";
        statusEl.style.color = "#dc2626";
        if (providerEl) providerEl.textContent = "-";
        if (modelEl) modelEl.textContent = "-";
        if (successEl) successEl.textContent = "0";
        if (failureEl) failureEl.textContent = "0";
        if (lastVarsEl) lastVarsEl.textContent = "-";
        if (lastLatencyEl) lastLatencyEl.textContent = "-";
        if (lastModeEl) lastModeEl.textContent = "-";
        if (stageLatenciesEl) stageLatenciesEl.textContent = "-";
        if (lastTagsEl) lastTagsEl.textContent = "-";
        if (lastTimeEl) lastTimeEl.textContent = "-";
        if (lastErrorEl) lastErrorEl.textContent = err.message;
        if (!silent) {
            showNotification(`AI 服务不可用: ${err.message}`, "warning");
        }
    }
}

async function refreshAIServiceStatusManual() {
    await refreshAIServiceStatus({ silent: false });
    showNotification("AI 服务状态已刷新", "info", 1200);
}

function startAIStatusWatcher() {
    if (aiStatusTimer) clearInterval(aiStatusTimer);
    refreshAIServiceStatus({ silent: true });
    aiStatusTimer = setInterval(() => {
        refreshAIServiceStatus({ silent: true });
    }, 5000);
}

async function autoGenerateForm() {
    // 立即显示进度条
    const progressPanel = document.getElementById("ai-progress-panel");
    const progressBar = document.getElementById("ai-progress-bar");
    const progressStatus = document.getElementById("ai-progress-status");
    
    const updateProgress = (current, total, status) => {
        if (progressBar) {
            const percent = total > 0 ? Math.round((current / total) * 100) : 0;
            progressBar.style.width = percent + "%";
        }
        if (progressStatus) progressStatus.textContent = status;
        console.log(`[AI Progress] ${current}/${total}: ${status}`);
    };
    
    try {
        if (!window.AISkill || !window.AIParser) {
            showNotification("AI 模块未加载", "error");
            return;
        }
        if (!window.wpsAdapter || !window.wpsAdapter.getDocumentText) {
            showNotification("WPS 适配器未就绪", "error");
            return;
        }

        await refreshAIServiceStatus({ silent: true });
        
        // 立即显示进度面板
        if (progressPanel) progressPanel.classList.add("show");
        updateProgress(0, 1, "正在读取文档...");
        
        const docText = await window.wpsAdapter.getDocumentText();
        if (!docText || !docText.trim()) {
            if (progressPanel) progressPanel.classList.remove("show");
            showNotification("未读取到文档内容", "warning");
            return;
        }
        
        updateProgress(0, 1, "文档读取完成，准备分析...");

        const aiOutput = await window.AISkill.analyzeDocument(docText, updateProgress);

        const parsed = window.AIParser.parseAIOutput(aiOutput);
        if (!parsed.success) {
            const errMsg = parsed.errors && parsed.errors.length ? parsed.errors[0] : "AI 解析失败";
            showNotification(errMsg, "error");
            return;
        }
        
        const aiSections = parsed.config || [];
        aiSections.forEach(section => {
            if (section.fields) {
                section.fields.forEach(f => { f.isCustom = true; });
            }
        });
        
        contractConfig = [...DEFAULT_CONTRACT_CONFIG, ...aiSections];
        savedFormData = {};
        await DocumentStateManager.save(true);
        buildForm();

        const count = parsed.stats?.total || aiSections.reduce((sum, sec) => sum + (sec.fields?.length || 0), 0);
        let embedSummary = null;

        if (isAutoEmbedEnabled()) {
            updateProgress(0, 1, "识别完成，正在自动执行批量埋点...");
            embedSummary = await batchEmbedAIFields({ suppressNotifications: true, source: "auto" });
        }

        if (progressPanel) {
            setTimeout(() => {
                progressPanel.classList.remove("show");
            }, 1000);
        }

        if (embedSummary) {
            if (embedSummary.total <= 0) {
                showNotification(`AI 识别完成，生成 ${count} 个字段；未发现可自动埋点字段。`, "warning");
            } else if (embedSummary.failCount === 0) {
                showNotification(`AI 识别完成，生成 ${count} 个字段；自动埋点成功 ${embedSummary.successCount} 个。`, "success");
            } else {
                showNotification(`AI 识别完成，生成 ${count} 个字段；自动埋点成功 ${embedSummary.successCount}，失败 ${embedSummary.failCount}。`, embedSummary.failCount > embedSummary.successCount ? "warning" : "info");
            }
        } else {
            showNotification(`AI 识别完成，生成 ${count} 个字段。请检查后点击"批量埋点"`, "success");
        }

        await refreshAIServiceStatus({ silent: true });
    } catch (err) {
        console.error("[AI] Error:", err);
        // 隐藏进度条
        if (progressPanel) progressPanel.classList.remove("show");
        showNotification(`AI 识别失败: ${err.message}`, "error");
        await refreshAIServiceStatus({ silent: true });
    }
}

/**
 * 批量埋点可视化
 */
function escapeEmbedHTML(value) {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function getEmbedEntryStatus(entry) {
    if (!entry) return "fail";
    if (entry.status) return entry.status;
    return entry.success ? "success" : "fail";
}

function renderBatchEmbedReport(summary, entries) {
    const panel = document.getElementById("embed-report-panel");
    const summaryEl = document.getElementById("embed-report-summary");
    const listEl = document.getElementById("embed-report-list");
    if (!panel || !summaryEl || !listEl) return;

    if (!summary || !entries || entries.length === 0) {
        panel.style.display = "none";
        summaryEl.textContent = "";
        listEl.innerHTML = "";
        return;
    }

    const sourceLabel = summary.source === "auto" ? "自动" : "手动";
    const filteredCount = summary.filteredCount || 0;
    const skipCount = summary.skipCount || 0;
    summaryEl.textContent = `来源：${sourceLabel} ｜ 新增 ${summary.successCount}/${summary.total} ｜ 跳过 ${skipCount} ｜ 失败 ${summary.failCount} ｜ 过滤 ${filteredCount}`;

    const html = entries.map((entry, idx) => {
        const status = getEmbedEntryStatus(entry);
        const statusClass = status;
        const statusMap = {
            success: "成功",
            skip: "跳过",
            filtered: "过滤",
            fail: "失败"
        };
        const statusText = statusMap[status] || "失败";
        const strategy = entry.strategy || "unknown";
        const score = typeof entry.score === "number" ? entry.score : "-";
        const reason = entry.success ? "" : (entry.reason || "未命中");
        const snippet = entry.snippet ? `<div class="embed-row-snippet">${escapeEmbedHTML(entry.snippet)}</div>` : "";
        const anchorInfo = (entry.prefixOk === undefined && entry.suffixOk === undefined)
            ? ""
            : `<span class="embed-chip">前缀:${entry.prefixOk ? "✓" : "✗"}</span><span class="embed-chip">后缀:${entry.suffixOk ? "✓" : "✗"}</span>`;

        return `
            <div class="embed-row ${statusClass}">
                <div class="embed-row-main">
                    <span class="embed-row-index">${idx + 1}.</span>
                    <span class="embed-row-label">${escapeEmbedHTML(entry.label)}</span>
                    <span class="embed-row-status ${statusClass}">${statusText}</span>
                </div>
                <div class="embed-row-meta">
                    <span class="embed-chip">策略:${escapeEmbedHTML(strategy)}</span>
                    <span class="embed-chip">得分:${score}</span>
                    ${anchorInfo}
                </div>
                ${reason ? `<div class="embed-row-reason">${escapeEmbedHTML(reason)}</div>` : ""}
                ${snippet}
            </div>
        `;
    }).join("");

    listEl.innerHTML = html;
    panel.style.display = "block";
}

function getAICandidateFilterReason(field) {
    if (!field) return "空字段";
    const tag = String(field.tag || "");
    const label = String(field.label || "");
    const aiContext = field.aiContext || {};
    const context = String(aiContext.context || "");
    const placeholder = String(aiContext.placeholder || "");

    if (/^Field_\d+$/i.test(tag)) {
        return "通用临时字段（Field_x）";
    }

    const samplePattern = /请输入|示例|样例|example/i;
    if (samplePattern.test(context) || samplePattern.test(placeholder) || samplePattern.test(label)) {
        return "样例占位字段（包含“请输入/示例”）";
    }

    const durationPattern = /agreement_duration|effective.*date|validity.*date/i;
    if ((durationPattern.test(tag) || /协议有效期至|有效期至/.test(label)) && /年\s*月\s*日/.test(placeholder.replace(/[_＿\s]+/g, ""))) {
        return "泛化日期字段（已由年/月等更细粒度字段覆盖）";
    }

    return "";
}

/**
 * 批量埋点：根据 AI 识别的 placeholder 在文档中搜索并插入 Content Control
 */
async function batchEmbedAIFields(options = {}) {
    const { suppressNotifications = false, source = "manual" } = options;

    if (!window.wpsAdapter || !window.wpsAdapter.insertControlBySearch) {
        if (!suppressNotifications) showNotification("WPS 适配器未就绪", "error");
        const summary = {
            source,
            total: 0,
            attempted: 0,
            successCount: 0,
            skipCount: 0,
            filteredCount: 0,
            failCount: 0,
            failedFields: [{ label: "系统", reason: "WPS 适配器未就绪" }],
            skippedFields: [],
            filteredFields: [],
            entries: []
        };
        renderBatchEmbedReport(summary, []);
        return summary;
    }

    // 收集所有 AI 识别的字段（isCustom 且有 _aiContext）
    const aiFieldsRaw = [];
    contractConfig.forEach(section => {
        if (section.fields) {
            section.fields.forEach(field => {
                if (field.isCustom && field._aiContext && field._aiContext.placeholder) {
                    const aiContext = field._aiContext;
                    const composed = `${aiContext.prefix || ''}${aiContext.placeholder || ''}${aiContext.suffix || ''}`;
                    const searchCandidates = [
                        aiContext.context,
                        composed,
                        aiContext.placeholder
                    ].filter(v => typeof v === "string" && v.trim());

                    aiFieldsRaw.push({
                        tag: field.tag,
                        label: field.label,
                        searchCandidates,
                        aiContext
                    });
                }
            });
        }
    });

    if (aiFieldsRaw.length === 0) {
        if (!suppressNotifications) showNotification("没有可埋点的 AI 识别字段", "warning");
        const summary = {
            source,
            total: 0,
            attempted: 0,
            successCount: 0,
            skipCount: 0,
            filteredCount: 0,
            failCount: 0,
            failedFields: [],
            skippedFields: [],
            filteredFields: [],
            entries: []
        };
        renderBatchEmbedReport(summary, []);
        return summary;
    }

    const filteredFields = [];
    const aiFields = [];
    aiFieldsRaw.forEach((field) => {
        const filterReason = getAICandidateFilterReason(field);
        if (filterReason) {
            filteredFields.push({ label: field.label, reason: filterReason });
            return;
        }
        aiFields.push(field);
    });

    const entries = filteredFields.map((item) => ({
        tag: "",
        label: item.label,
        success: false,
        status: "filtered",
        reason: item.reason,
        strategy: "filter",
        score: null,
        prefixOk: undefined,
        suffixOk: undefined,
        snippet: ""
    }));

    const progressPanel = document.getElementById("ai-progress-panel");
    const progressBar = document.getElementById("ai-progress-bar");
    const progressStatus = document.getElementById("ai-progress-status");

    if (progressPanel) progressPanel.classList.add("show");
    if (progressBar) progressBar.style.width = "0%";
    if (progressStatus) progressStatus.textContent = aiFields.length ? "准备批量埋点..." : "没有需要执行的字段";

    let successCount = 0;
    let skipCount = 0;
    let failCount = 0;
    const failedFields = [];
    const skippedFields = [];

    for (let i = 0; i < aiFields.length; i++) {
        const field = aiFields[i];
        const progress = Math.round(((i + 1) / aiFields.length) * 100);

        if (progressBar) progressBar.style.width = progress + "%";
        if (progressStatus) progressStatus.textContent = `正在埋点: ${field.label} (${i + 1}/${aiFields.length})`;

        try {
            const result = await window.wpsAdapter.insertControlBySearch(
                field.tag,
                field.label,
                {
                    searchCandidates: field.searchCandidates,
                    aiContext: field.aiContext
                }
            );

            const detail = result && result.detail ? result.detail : {};
            const alreadyExists = detail.status === "already_exists" || (result && typeof result.message === "string" && result.message.indexOf("已存在埋点") === 0);
            const overlapConflict = detail.status === "conflict";
            const status = (result && result.success)
                ? "success"
                : ((alreadyExists || overlapConflict) ? "skip" : "fail");

            const entry = {
                tag: field.tag,
                label: field.label,
                success: !!(result && result.success),
                status,
                reason: result && result.message ? result.message : "未知结果",
                strategy: detail.strategy || ((alreadyExists || overlapConflict) ? "existing" : "unknown"),
                score: typeof detail.score === "number" ? detail.score : null,
                prefixOk: detail.prefixOk,
                suffixOk: detail.suffixOk,
                snippet: detail.snippet || ""
            };
            entries.push(entry);

            if (status === "success") {
                successCount++;
                console.log(`[Batch Embed] ✓ ${field.tag}: ${entry.reason}`);
            } else if (status === "skip") {
                skipCount++;
                const skipReason = overlapConflict ? `与已有埋点位置重叠，自动跳过: ${entry.reason}` : entry.reason;
                if (overlapConflict) entry.reason = skipReason;
                skippedFields.push({ label: field.label, reason: entry.reason });
                console.log(`[Batch Embed] ↷ ${field.tag}: ${entry.reason}`);
            } else {
                failCount++;
                failedFields.push({ label: field.label, reason: entry.reason, strategy: entry.strategy, score: entry.score });
                console.warn(`[Batch Embed] ✗ ${field.tag}: ${entry.reason}`);
            }
        } catch (err) {
            failCount++;
            const reason = err && err.message ? err.message : "未知错误";
            failedFields.push({ label: field.label, reason });
            entries.push({
                tag: field.tag,
                label: field.label,
                success: false,
                status: "fail",
                reason,
                strategy: "exception",
                score: null,
                prefixOk: undefined,
                suffixOk: undefined,
                snippet: ""
            });
            console.error(`[Batch Embed] Error ${field.tag}:`, err);
        }

        await new Promise(resolve => setTimeout(resolve, 100));
    }

    if (progressPanel) {
        setTimeout(() => progressPanel.classList.remove("show"), 1000);
    }

    const summary = {
        source,
        total: aiFieldsRaw.length,
        attempted: aiFields.length,
        successCount,
        skipCount,
        filteredCount: filteredFields.length,
        failCount,
        failedFields,
        skippedFields,
        filteredFields,
        entries
    };
    renderBatchEmbedReport(summary, entries);

    if (!suppressNotifications) {
        if (failCount === 0) {
            if (successCount === 0) {
                showNotification(`批量埋点完成：跳过 ${skipCount}，过滤 ${filteredFields.length}。`, "info");
            } else {
                showNotification(`批量埋点完成：新增 ${successCount}，跳过 ${skipCount}，过滤 ${filteredFields.length}。`, "success");
            }
        } else {
            showNotification(`批量埋点完成：新增 ${successCount}，跳过 ${skipCount}，失败 ${failCount}。`, failCount > successCount ? "error" : "warning");
            console.log("[Batch Embed] 失败的字段:", failedFields);
        }
    } else if (failCount > 0) {
        console.log("[Batch Embed] 失败的字段:", failedFields);
    }

    return summary;
}

async function updateContent(tag, value, label) {
    if (window.wpsAdapter) {
        await window.wpsAdapter.updateContent(tag, value, label);
        savedFormData[tag] = value;
        scheduleSmartSave();
    }
}

function insertControl(tag, label, isWrapper = false) {
    if (window.wpsAdapter) window.wpsAdapter.insertControl(tag, label, isWrapper);
}

function toggleVisibility(tag, visible, label) {
    if (window.wpsAdapter) window.wpsAdapter.toggleVisibility(tag, visible, label);
}

/* ==================================================================
 * 4. 渲染引擎
 * ================================================================== */

function buildForm() {
    const container = document.getElementById("form-container");
    const navList = document.getElementById("nav-list");
    if (!container) return;
    container.innerHTML = "";
    if (navList) navList.innerHTML = "";
    
    const visibleSections = contractConfig.filter((section) => !isPlaceholderOnlySection(section));

    if (!visibleSections.length) {
        container.innerHTML = "<p style='color:#999;text-align:center;padding:40px;'>暂无表单配置</p>";
        return;
    }
    
    visibleSections.forEach((section, idx) => {
        if (navList) {
            const navItem = document.createElement("li");
            navItem.className = "nav-item" + (idx === 0 ? " active" : "");
            const cleanLabel = section.header?.label ? section.header.label.replace(/^\d+\.\s*/, "") : `第${idx + 1}节`;
            navItem.innerHTML = `<span class="nav-num">${idx + 1}</span><span class="nav-text">${cleanLabel}</span>`;
            navItem.onclick = () => {
                document.querySelectorAll(".nav-item").forEach(el => el.classList.remove("active"));
                navItem.classList.add("active");
                const targetId = `sec-${section.id || idx}`;
                const target = document.getElementById(targetId);
                if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
            };
            navList.appendChild(navItem);
        }
        
        const secHeader = document.createElement("div");
        secHeader.className = "section-header";
        secHeader.id = `sec-${section.id || idx}`;
        secHeader.textContent = section.header.label;
        container.appendChild(secHeader);

        const fieldsDiv = document.createElement("div");
        fieldsDiv.className = "section-fields";
        (section.fields || []).forEach(f => createField(f, fieldsDiv, section.id));
        container.appendChild(fieldsDiv);
    });
    
    // 渲染自定义字段面板
    renderCustomFieldsPanel();
}

function createField(field, parent, sectionId) {
    if (field.type === "divider") {
        const div = document.createElement("div");
        div.className = "divider";
        div.textContent = field.label;
        parent.appendChild(div);
        return;
    }
    if (field.type === "html_placeholder") {
        const target = document.getElementById(field.targetId);
        if (target && field.targetId !== "local-sync-section") {
            parent.appendChild(target);
            target.style.display = "block";
        }
        return;
    }

    const wrap = document.createElement("div");
    wrap.className = "field" + (field.isCustom ? " custom-field-in-form" : "");
    wrap.style.position = "relative";
    if (sectionId) wrap.dataset.sectionId = sectionId;
    if (field.id) wrap.dataset.fieldId = field.id;
    
    // 添加编辑按钮
    const editBtn = document.createElement("button");
    editBtn.className = "field-edit-btn";
    editBtn.innerHTML = "⚙️";
    editBtn.title = "编辑字段";
    editBtn.onclick = (e) => {
        e.stopPropagation();
        editFieldInSection(sectionId, field.id);
    };
    wrap.appendChild(editBtn);
    
    const head = document.createElement("div");
    head.className = "field-header";
    head.innerHTML = `<span class="field-label">${field.label}</span>`;
    
    if (field.tag) {
        const btn = document.createElement("button");
        btn.className = "field-btn";
        btn.textContent = field.hasParagraphToggle ? "插入段落" : "插入";
        btn.onclick = () => insertControl(field.tag, field.label, field.hasParagraphToggle);
        head.appendChild(btn);
    }
    wrap.appendChild(head);

    const val = savedFormData[field.tag] || "";

    if (field.type === "radio") {
        const grp = document.createElement("div");
        grp.className = "radio-group";
        const groupName = `r_${sectionId || "sec"}_${field.id}`;
        (field.options || []).forEach(opt => {
            const item = document.createElement("label");
            item.className = "radio-item";
            const radio = document.createElement("input");
            radio.type = "radio";
            radio.name = groupName;
            radio.value = opt;
            if (field.tag) radio.dataset.tag = field.tag;
            if (val === opt) radio.checked = true;
            radio.onchange = () => {
                const show = ["是","有","适用","确认"].includes(opt);
                if (field.hasParagraphToggle) toggleVisibility(field.tag, show, field.label);
                else updateContent(field.tag, opt, field.label);
                if (field.tag) savedFormData[field.tag] = opt;
                scheduleSmartSave();
                
                // 处理 subFields 显示/隐藏
                if (field.subFields && field.subFields.length > 0) {
                    const subContainer = wrap.querySelector(".sub-fields-container");
                    if (subContainer) {
                        subContainer.style.display = show ? "block" : "none";
                    }
                }
            };
            item.appendChild(radio);
            item.appendChild(document.createTextNode(opt));
            grp.appendChild(item);
        });
        wrap.appendChild(grp);
        
        // 渲染 subFields（如果有）
        if (field.subFields && field.subFields.length > 0) {
            const subContainer = document.createElement("div");
            subContainer.className = "sub-fields-container";
            subContainer.style.marginLeft = "20px";
            subContainer.style.marginTop = "10px";
            subContainer.style.paddingLeft = "12px";
            subContainer.style.borderLeft = "2px solid var(--border)";
            const show = val && ["是","有","适用","确认"].includes(val);
            subContainer.style.display = show ? "block" : "none";
            
            field.subFields.forEach(sf => createField(sf, subContainer, sectionId));
            wrap.appendChild(subContainer);
        }
    } else {
        const inp = document.createElement("input");
        inp.className = "field-input";
        inp.type = field.type || "text";
        if (field.tag) inp.dataset.tag = field.tag;
        if (field.formatFn) inp.dataset.formatFn = field.formatFn;
        if (field.placeholder) inp.placeholder = field.placeholder;
        inp.value = val;
        inp.oninput = debounce(() => {
            let finalVal = inp.value;
            if (window.Formatters && field.formatFn) finalVal = window.Formatters.applyFormat(finalVal, field.formatFn);
            if (field.tag) savedFormData[field.tag] = finalVal;
            updateContent(field.tag, finalVal, field.label);
        }, 600);
        inp.onblur = () => DocumentStateManager.save();
        wrap.appendChild(inp);
    }
    parent.appendChild(wrap);
}

/* ==================================================================
 * 5. 未填写字段检查器 / 交付流程
 * ================================================================== */

let checkerState = {
    unfilledFields: [],
    currentIndex: 0,
    isOpen: false
};

function collectUnfilledFields() {
    const container = document.getElementById("form-container");
    if (!container) return [];
    
    const unfilled = [];
    container.querySelectorAll("input[data-tag], select[data-tag]").forEach(el => {
        const tag = el.dataset.tag;
        const value = el.value;
        if (el.type === "radio") return;
        if (!value || value.trim() === "") {
            const fieldEl = el.closest(".field");
            const labelEl = fieldEl?.querySelector(".field-label");
            const label = labelEl?.textContent || tag;
            unfilled.push({ tag, label, element: fieldEl });
        }
    });
    return unfilled;
}

function openChecker(fields) {
    checkerState.unfilledFields = fields;
    checkerState.currentIndex = 0;
    checkerState.isOpen = true;
    
    document.getElementById("checker-bar").classList.add("show");
    updateCheckerUI();
    highlightCurrentField();
}

window.closeChecker = function() {
    checkerState.isOpen = false;
    document.getElementById("checker-bar").classList.remove("show");
    document.querySelectorAll(".highlight").forEach(el => el.classList.remove("highlight"));
};

function updateCheckerUI() {
    const { unfilledFields, currentIndex } = checkerState;
    const total = unfilledFields.length;
    const current = currentIndex + 1;
    document.getElementById("checker-count").textContent = `${current}/${total}`;
    document.getElementById("checker-field").textContent = unfilledFields[currentIndex]?.label || "";
    document.getElementById("checker-prev").disabled = currentIndex === 0;
    document.getElementById("checker-next").disabled = currentIndex >= total - 1;
}

function highlightCurrentField() {
    document.querySelectorAll(".field.highlight").forEach(el => el.classList.remove("highlight"));
    const { unfilledFields, currentIndex } = checkerState;
    const field = unfilledFields[currentIndex];
    if (field && field.element) {
        field.element.classList.add("highlight");
        field.element.scrollIntoView({ behavior: "smooth", block: "center" });
    }
}

window.checkerPrev = function() {
    if (checkerState.currentIndex > 0) {
        checkerState.currentIndex--;
        updateCheckerUI();
        highlightCurrentField();
    }
};

window.checkerNext = function() {
    if (checkerState.currentIndex < checkerState.unfilledFields.length - 1) {
        checkerState.currentIndex++;
        updateCheckerUI();
        highlightCurrentField();
    }
};

function showDialog(title, content, options = {}) {
    return new Promise(resolve => {
        const overlay = document.getElementById("dialog-overlay");
        const box = document.getElementById("dialog-box");
        const titleEl = document.getElementById("dialog-title");
        const contentEl = document.getElementById("dialog-content");
        const confirmBtn = document.getElementById("dialog-confirm");
        const cancelBtn = document.getElementById("dialog-cancel");
        
        titleEl.textContent = title;
        contentEl.textContent = content;
        confirmBtn.textContent = options.confirmText || "确认";
        cancelBtn.textContent = options.cancelText || "取消";
        
        confirmBtn.className = options.danger ? "dialog-btn danger" : "dialog-btn confirm";
        
        overlay.classList.add("show");
        box.classList.add("show");
        
        const cleanup = (result) => {
            overlay.classList.remove("show");
            box.classList.remove("show");
            confirmBtn.onclick = null;
            cancelBtn.onclick = null;
            resolve(result);
        };
        
        confirmBtn.onclick = () => cleanup(true);
        cancelBtn.onclick = () => cleanup(false);
    });
}

const deliverSteps = [
    { id: "check", label: "检查未填写字段" },
    { id: "backup", label: "备份文档" },
    { id: "confirm", label: "确认清理" },
    { id: "cleanup", label: "清理隐藏段落" },
    { id: "done", label: "完成" }
];

function showProgress() {
    document.getElementById("dialog-overlay").classList.add("show");
    const panel = document.getElementById("progress-panel");
    const stepsDiv = document.getElementById("progress-steps");
    
    stepsDiv.innerHTML = deliverSteps.map(step => `
        <div class="progress-step" id="step-${step.id}">
            <span class="progress-step-icon">○</span>
            <span>${step.label}</span>
        </div>
    `).join("");
    
    panel.classList.add("show");
}

function updateStepStatus(stepId, status) {
    const stepEl = document.getElementById(`step-${stepId}`);
    if (!stepEl) return;
    stepEl.className = `progress-step ${status}`;
    const icon = stepEl.querySelector(".progress-step-icon");
    if (status === "active") icon.textContent = "◉";
    else if (status === "done") icon.textContent = "✓";
    else if (status === "error") icon.textContent = "✕";
    else icon.textContent = "○";
}

function updateProgressStatus(text) {
    document.getElementById("progress-status").textContent = text;
}

function hideProgress() {
    document.getElementById("dialog-overlay").classList.remove("show");
    document.getElementById("progress-panel").classList.remove("show");
}

function delay(ms) {
    return new Promise(r => setTimeout(r, ms));
}

window.proceedToBackup = async function() {
    closeChecker();
    await startDeliverProcess(true);
};

window.deliverContract = async function() {
    const unfilled = collectUnfilledFields();
    if (unfilled.length > 0) openChecker(unfilled);
    else await startDeliverProcess(true);
};

window.undoAllEmbeds = async function() {
    const confirmed = await showDialog(
        "⚠️ 撤销所有埋点",
        "此操作将删除文档中所有 Content Control（蓝色框），但保留里面的文字内容。\n\n" +
        "⚠️ 此操作不可撤销！\n\n" +
        "是否继续？",
        { confirmText: "确认撤销", cancelText: "取消", danger: true }
    );
    
    if (!confirmed) return;
    
    try {
        if (!window.wpsAdapter) {
            showNotification("WPS 适配器未初始化", "error");
            return;
        }
        
        console.log("[Undo] 开始撤销所有埋点...");
        const result = await window.wpsAdapter.undoAllEmbeds();
        console.log(`[Undo] 完成！删除了 ${result.deletedCount} 个埋点`);
        
        if (result.errors && result.errors.length > 0) {
            console.warn("[Undo] 部分失败:", result.errors);
            showNotification(`撤销完成，但有 ${result.errors.length} 个失败`, "warning");
        } else if (result.deletedCount > 0) {
            showNotification(`成功撤销 ${result.deletedCount} 个埋点`, "success");
        } else {
            showNotification("文档中没有埋点", "info");
        }
    } catch (err) {
        console.error("[Undo] 错误:", err);
        showNotification(`撤销失败: ${err.message}`, "error");
    }
};

async function startDeliverProcess(checkPassed = false) {
    showProgress();
    
    try {
        updateStepStatus("check", "active");
        updateProgressStatus("检查表单...");
        await delay(300);
        
        if (checkPassed) {
            updateStepStatus("check", "done");
        } else {
            const unfilled = collectUnfilledFields();
            if (unfilled.length > 0) {
                updateStepStatus("check", "error");
                updateProgressStatus(`发现 ${unfilled.length} 个未填写字段`);
                await delay(1000);
                hideProgress();
                openChecker(unfilled);
                return;
            }
            updateStepStatus("check", "done");
        }
        await delay(200);
        
        updateStepStatus("backup", "active");
        updateProgressStatus("准备备份...");
        await delay(300);
        hideProgress();
        
        const backupConfirm = await showDialog(
            "📁 备份提醒",
            "建议在交付前备份当前文档：\n\n" +
            "• 点击「自动备份」创建副本\n" +
            "• 或手动在文件夹中复制一份\n\n" +
            "备份后点击「继续」执行清理。",
            { confirmText: "自动备份", cancelText: "跳过备份" }
        );
        
        showProgress();
        
        if (backupConfirm && window.wpsAdapter) {
            updateProgressStatus("正在备份文档...");
            const backupResult = await window.wpsAdapter.backupDocument();
            
            if (backupResult.success) {
                updateStepStatus("backup", "done");
                updateProgressStatus(`✓ 备份已创建: ${backupResult.fileName}`);
                showNotification(`备份已创建: ${backupResult.fileName}`, "success");
            } else {
                updateStepStatus("backup", "error");
                updateProgressStatus(`备份失败: ${backupResult.error}`);
                await delay(1000);
                
                hideProgress();
                const manualConfirm = await showDialog(
                    "⚠️ 备份失败",
                    `自动备份失败：${backupResult.error}\n\n` +
                    "请手动备份文档后继续。\n" +
                    "确认已完成手动备份？",
                    { confirmText: "已备份，继续", cancelText: "取消" }
                );
                
                if (!manualConfirm) {
                    showNotification("操作已取消", "warning");
                    return;
                }
                showProgress();
                updateStepStatus("backup", "done");
            }
        } else {
            updateStepStatus("backup", "done");
            updateProgressStatus("跳过备份");
        }
        await delay(300);
        
        updateStepStatus("confirm", "active");
        updateProgressStatus("等待确认...");
        hideProgress();
        
        const hiddenCount = window.wpsAdapter ? window.wpsAdapter.getHiddenParagraphs().length : 0;
        const finalConfirm = await showDialog(
            "⚠️ 确认清理",
            `即将执行以下操作：\n\n` +
            `• 删除 ${hiddenCount} 个隐藏段落\n` +
            `• 生成干净的交付版本\n\n` +
            `⚠️ 此操作不可撤销！`,
            { confirmText: "确认清理", cancelText: "取消", danger: true }
        );
        
        if (!finalConfirm) {
            showNotification("操作已取消", "warning");
            return;
        }
        
        showProgress();
        updateStepStatus("confirm", "done");
        await delay(200);
        
        updateStepStatus("cleanup", "active");
        updateProgressStatus("正在清理隐藏段落...");
        
        let deletedCount = 0;
        if (window.wpsAdapter) {
            const result = await window.wpsAdapter.executeCleanup();
            deletedCount = result.deletedCount || 0;
        }
        
        updateStepStatus("cleanup", "done");
        updateProgressStatus(`已删除 ${deletedCount} 个隐藏段落`);
        await delay(300);
        
        updateStepStatus("done", "done");
        updateProgressStatus("🎉 合同已准备好交付！");
        await delay(1500);
        hideProgress();
        
        showNotification("🎉 合同已完成！文档已准备好交付。", "success");
        
    } catch (err) {
        console.error("[Deliver] Error:", err);
        updateProgressStatus(`错误: ${err.message}`);
        await delay(2000);
        hideProgress();
        showNotification(`交付失败: ${err.message}`, "error");
    }
}

function buildLabelMap() {
    const map = {};
    function walk(fields) {
        if (!fields) return;
        fields.forEach(f => {
            if (f.tag && f.label) map[f.tag] = f.label;
            if (f.subFields) walk(f.subFields);
        });
    }
    contractConfig.forEach(section => {
        walk(section.fields);
        walk(section.shareholderFields);
        walk(section.investorFields);
    });
    return map;
}

async function applyDocumentState(state) {
    if (state && state.config) contractConfig = state.config;
    else await loadFormConfig();

    savedFormData = state && state.formData ? state.formData : {};

    if (state && state.rounds) {
        enabledShareholders = state.rounds.enabledShareholders || {};
        enabledInvestors = state.rounds.enabledInvestors || {};
    } else {
        enabledShareholders = {};
        enabledInvestors = {};
    }

    if (window.wpsAdapter && window.wpsAdapter.readFormDataFromDocumentCCs) {
        const ccData = await window.wpsAdapter.readFormDataFromDocumentCCs();
        if (ccData && Object.keys(ccData).length > 0) {
            savedFormData = { ...savedFormData, ...ccData };
        }
    }

    buildForm();
}

async function refreshFromActiveDocument(reason = "") {
    if (isDocRefreshing) return;
    isDocRefreshing = true;
    try {
        await syncRuntimeConfigFromManifest();
        const state = await DocumentStateManager.load();
        if (reason) console.log(`[DocSwitch] Refresh from active document: ${reason}`);
        await applyDocumentState(state);
    } catch (err) {
        console.error("[DocSwitch] Refresh error:", err);
    } finally {
        isDocRefreshing = false;
    }
}

function startDocumentWatcher() {
    if (!window.wpsAdapter || !window.wpsAdapter.getActiveDocumentKey) return;
    if (docWatchTimer) clearInterval(docWatchTimer);

    docWatchTimer = setInterval(async () => {
        if (docWatchInFlight) return;
        docWatchInFlight = true;
        try {
            const key = await window.wpsAdapter.getActiveDocumentKey();
            if (!key) return;
            if (key !== activeDocKey) {
                activeDocKey = key;
                await refreshFromActiveDocument("doc-changed");
            }
        } catch (err) {
            console.warn('[DocSwitch] watcher tick failed:', err && err.message ? err.message : err);
        } finally {
            docWatchInFlight = false;
        }
    }, 1500);
}

/* ==================================================================
 * 5. 初始化
 * ================================================================== */

async function init() {
    console.log("[Taskpane] WPS Clean Start v8...");
    const loading = document.getElementById("app-loading-overlay");
    if (loading) loading.style.display = "flex";

    try {
        await syncRuntimeConfigFromManifest();
        const state = await DocumentStateManager.load();
        if (state) {
            console.log("[Init] Restored from document");
        } else {
            console.log("[Init] Empty document, using base config");
        }
        await applyDocumentState(state);
        if (window.wpsAdapter && window.wpsAdapter.getActiveDocumentKey) {
            activeDocKey = await window.wpsAdapter.getActiveDocumentKey();
        }
        initAutoEmbedPreference();
        startDocumentWatcher();
        startAIStatusWatcher();
        initCustomFieldsManager();
    } catch (e) {
        console.error(e);
    } finally {
        if (loading) loading.style.display = "none";
    }
}

function debounce(fn, wait) {
    let t;
    return function(...args) {
        clearTimeout(t);
        t = setTimeout(() => fn.apply(this, args), wait);
    };
}

function initAutoEmbedPreference() {
    const checkbox = document.getElementById("auto-embed-after-ai");
    let enabled = true;

    try {
        const raw = localStorage.getItem(AUTO_EMBED_STORAGE_KEY);
        if (raw === "0" || raw === "false") enabled = false;
    } catch (err) {
        console.warn("[AutoEmbed] read preference failed:", err && err.message ? err.message : err);
    }

    autoEmbedAfterAI = enabled;
    if (checkbox) checkbox.checked = enabled;
}

function setAutoEmbedPreference(enabled) {
    autoEmbedAfterAI = !!enabled;
    try {
        localStorage.setItem(AUTO_EMBED_STORAGE_KEY, autoEmbedAfterAI ? "1" : "0");
    } catch (err) {
        console.warn("[AutoEmbed] save preference failed:", err && err.message ? err.message : err);
    }
}

function isAutoEmbedEnabled() {
    return autoEmbedAfterAI !== false;
}

function onAutoEmbedToggleChange(checked) {
    setAutoEmbedPreference(checked);
    showNotification(checked ? "已开启：识别后自动批量埋点" : "已关闭：识别后自动批量埋点", "info", 1200);
}

function isPlaceholderOnlySection(section) {
    const fields = section && Array.isArray(section.fields) ? section.fields : [];
    return fields.length > 0 && fields.every((field) => field && field.type === "html_placeholder");
}

function collectForm() {
    const res = { ...savedFormData };
    document.querySelectorAll(".field-input, input:checked").forEach(el => {
        const tag = el.dataset.tag;
        if (!tag) return;
        const formatFn = el.dataset.formatFn;
        if (formatFn && window.Formatters) {
            const cached = savedFormData[tag];
            res[tag] = cached !== undefined ? cached : window.Formatters.applyFormat(el.value, formatFn);
        } else {
            res[tag] = el.value;
        }
    });
    savedFormData = { ...res };
    return res;
}

/* ==================================================================
 * 表单管理功能 - 添加/编辑/删除字段
 * ================================================================== */

// 拼音映射表
const PINYIN_MAP = {
    '测': 'Ce', '试': 'Shi', '字': 'Zi', '段': 'Duan', '特': 'Te', '殊': 'Shu',
    '条': 'Tiao', '款': 'Kuan', '额': 'E', '外': 'Wai', '费': 'Fei', '用': 'Yong',
    '投': 'Tou', '资': 'Zi', '金': 'Jin', '比': 'Bi', '例': 'Li', '期': 'Qi',
    '限': 'Xian', '日': 'Ri', '合': 'He', '同': 'Tong', '价': 'Jia',
    '格': 'Ge', '数': 'Shu', '量': 'Liang', '总': 'Zong', '计': 'Ji', '备': 'Bei',
    '注': 'Zhu', '说': 'Shuo', '明': 'Ming', '描': 'Miao', '述': 'Shu', '名': 'Ming',
    '称': 'Cheng', '地': 'Di', '址': 'Zhi', '电': 'Dian', '话': 'Hua', '邮': 'You',
    '箱': 'Xiang', '联': 'Lian', '系': 'Xi', '人': 'Ren', '公': 'Gong', '司': 'Si',
    '股': 'Gu', '东': 'Dong', '权': 'Quan', '益': 'Yi', '法': 'Fa', '定': 'Ding',
    '代': 'Dai', '表': 'Biao', '册': 'Ce', '本': 'Ben',
    '实': 'Shi', '缴': 'Jiao', '认': 'Ren', '购': 'Gou', '买': 'Mai', '卖': 'Mai',
    '转': 'Zhuan', '让': 'Rang', '受': 'Shou', '增': 'Zeng', '减': 'Jian', '持': 'Chi',
    '有': 'You', '占': 'Zhan', '方': 'Fang', '甲': 'Jia', '乙': 'Yi', '丙': 'Bing',
    '丁': 'Ding', '戊': 'Wu', '己': 'Ji', '董': 'Dong', '事': 'Shi', '会': 'Hui',
    '监': 'Jian', '经': 'Jing', '理': 'Li', '财': 'Cai', '务': 'Wu',
    '行': 'Xing', '政': 'Zheng', '部': 'Bu', '门': 'Men', '销': 'Xiao', '售': 'Shou',
    '市': 'Shi', '场': 'Chang', '研': 'Yan', '发': 'Fa', '技': 'Ji', '术': 'Shu',
    '产': 'Chan', '品': 'Pin', '服': 'Fu', '项': 'Xiang', '目': 'Mu',
    '工': 'Gong', '程': 'Cheng', '建': 'Jian', '设': 'She', '开': 'Kai',
    '生': 'Sheng', '制': 'Zhi', '造': 'Zao',
    '加': 'Jia', '质': 'Zhi', '标': 'Biao', '准': 'Zhun',
    '规': 'Gui', '范': 'Fan', '流': 'Liu', '步': 'Bu', '骤': 'Zhou',
    '时': 'Shi', '间': 'Jian', '始': 'Shi', '结': 'Jie', '束': 'Shu',
    '完': 'Wan', '成': 'Cheng', '进': 'Jin', '度': 'Du', '状': 'Zhuang', '态': 'Tai',
    '正': 'Zheng', '常': 'Chang', '异': 'Yi', '错': 'Cuo', '误': 'Wu',
    '功': 'Gong', '失': 'Shi', '败': 'Bai', '取': 'Qu', '消': 'Xiao',
    '确': 'Que', '提': 'Ti', '交': 'Jiao', '保': 'Bao', '存': 'Cun',
    '删': 'Shan', '除': 'Chu', '修': 'Xiu', '改': 'Gai', '编': 'Bian', '辑': 'Ji',
    '查': 'Cha', '看': 'Kan', '详': 'Xiang', '情': 'Qing', '列': 'Lie',
    '搜': 'Sou', '索': 'Suo', '过': 'Guo', '滤': 'Lv', '排': 'Pai', '序': 'Xu',
    '分': 'Fen', '类': 'Lei', '签': 'Qian', '属': 'Shu', '性': 'Xing',
    '值': 'Zhi', '默': 'Mo', '选': 'Xuan', '择': 'Ze', '必': 'Bi',
    '填': 'Tian', '可': 'Ke', '只': 'Zhi', '读': 'Du', '隐': 'Yin',
    '藏': 'Cang', '显': 'Xian', '示': 'Shi', '启': 'Qi', '禁': 'Jin',
    '锁': 'Suo', '解': 'Jie', '配': 'Pei', '置': 'Zhi', '统': 'Tong', '管': 'Guan', '员': 'Yuan', '户': 'Hu',
    '角': 'Jiao', '色': 'Se', '访': 'Fang', '问': 'Wen',
    '登': 'Deng', '录': 'Lu', '退': 'Tui', '出': 'Chu',
    '密': 'Mi', '码': 'Ma', '账': 'Zhang', '号': 'Hao', '件': 'Jian',
    '手': 'Shou', '机': 'Ji', '验': 'Yan', '证': 'Zheng', '审': 'Shen', '核': 'He',
    '批': 'Pi', '拒': 'Ju', '绝': 'Jue', '通': 'Tong', '过': 'Guo',
    '待': 'Dai', '处': 'Chu', '已': 'Yi', '未': 'Wei', '新': 'Xin', '旧': 'Jiu',
    '创': 'Chuang', '更': 'Geng', '版': 'Ban',
    '上': 'Shang', '下': 'Xia', '左': 'Zuo', '右': 'You', '前': 'Qian', '后': 'Hou',
    '第': 'Di', '一': 'Yi', '二': 'Er', '三': 'San', '四': 'Si', '五': 'Wu',
    '六': 'Liu', '七': 'Qi', '八': 'Ba', '九': 'Jiu', '十': 'Shi', '百': 'Bai',
    '千': 'Qian', '万': 'Wan', '亿': 'Yi', '元': 'Yuan', '角': 'Jiao',
    '年': 'Nian', '月': 'Yue', '周': 'Zhou', '天': 'Tian', '小': 'Xiao', '大': 'Da',
    '长': 'Chang', '短': 'Duan', '高': 'Gao', '低': 'Di', '多': 'Duo', '少': 'Shao',
    '好': 'Hao', '坏': 'Huai', '对': 'Dui', '是': 'Shi', '否': 'Fou',
    '真': 'Zhen', '假': 'Jia', '空': 'Kong', '满': 'Man', '全': 'Quan',
    '其': 'Qi', '他': 'Ta', '她': 'Ta', '它': 'Ta', '们': 'Men', '的': 'De',
    '得': 'De', '和': 'He', '与': 'Yu', '或': 'Huo', '及': 'Ji',
    '但': 'Dan', '而': 'Er', '且': 'Qie', '因': 'Yin', '为': 'Wei', '所': 'Suo',
    '以': 'Yi', '如': 'Ru', '果': 'Guo', '则': 'Ze', '虽': 'Sui',
    '然': 'Ran', '仍': 'Reng', '还': 'Hai', '也': 'Ye', '都': 'Dou', '每': 'Mei',
    '各': 'Ge', '某': 'Mou', '些': 'Xie', '这': 'Zhe', '那': 'Na', '哪': 'Na',
    '谁': 'Shui', '什': 'Shen', '么': 'Me', '怎': 'Zen', '样': 'Yang',
    '里': 'Li', '何': 'He', '几': 'Ji', '许': 'Xu', '很': 'Hen', '非': 'Fei',
    '最': 'Zui', '再': 'Zai', '又': 'You', '才': 'Cai', '就': 'Jiu',
    '仅': 'Jin', '即': 'Ji', '便': 'Bian', '若': 'Ruo', '倘': 'Tang',
    '既': 'Ji', '由': 'You', '于': 'Yu', '至': 'Zhi', '到': 'Dao', '从': 'Cong',
    '向': 'Xiang', '往': 'Wang', '在': 'Zai', '被': 'Bei', '把': 'Ba', '将': 'Jiang',
    '给': 'Gei', '使': 'Shi', '令': 'Ling', '要': 'Yao', '需': 'Xu',
    '应': 'Ying', '该': 'Gai', '能': 'Neng', '须': 'Xu',
    '想': 'Xiang', '愿': 'Yuan', '希': 'Xi',
    '望': 'Wang', '请': 'Qing', '谢': 'Xie', '不': 'Bu', '起': 'Qi',
    '抱': 'Bao', '歉': 'Qian', '见': 'Jian', '您': 'Nin',
    '反': 'Fan', '稀': 'Xi', '释': 'Shi', '补': 'Bu', '偿': 'Chang', '优': 'You',
    '先': 'Xian', '回': 'Hui', '赎': 'Shu', '清': 'Qing',
    '算': 'Suan', '领': 'Ling', '跟': 'Gen', '决': 'Jue', '票': 'Piao',
    '托': 'Tuo', '信': 'Xin', '息': 'Xi', '报': 'Bao', '告': 'Gao',
    '季': 'Ji', '预': 'Yu',
    '知': 'Zhi', '披': 'Pi', '露': 'Lu', '陈': 'Chen',
    '障': 'Zhang', '违': 'Wei', '约': 'Yue', '责': 'Ze', '任': 'Ren', '赔': 'Pei',
    '争': 'Zheng', '议': 'Yi', '仲': 'Zhong', '裁': 'Cai', '诉': 'Su', '讼': 'Song'
};

/**
 * 将中文转换为拼音 (PascalCase)
 */
function toPinyin(chinese) {
    if (!chinese) return '';
    let result = '';
    for (const char of chinese) {
        if (PINYIN_MAP[char]) {
            result += PINYIN_MAP[char];
        } else if (/[a-zA-Z0-9]/.test(char)) {
            result += char;
        } else if (/[\u4e00-\u9fa5]/.test(char)) {
            // 未知汉字，使用 Unicode 编码
            result += 'U' + char.charCodeAt(0).toString(16).toUpperCase();
        }
        // 忽略其他字符（空格、标点等）
    }
    return result || 'CustomField';
}

/**
 * HTML 转义
 */
function escapeHtml(text) {
            const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

/**
 * 显示添加字段弹窗
 */
function showAddFieldModal() {
    const modal = document.getElementById("add-field-modal");
    if (modal) {
        modal.classList.add("show");
        // 清空表单
        const labelInput = document.getElementById("field-label");
        labelInput.value = "";
        document.getElementById("field-type").value = "text";
        document.getElementById("options-group").style.display = "none";
        document.getElementById("tag-preview").style.display = "none";
        document.getElementById("tag-preview-text").textContent = "";
        
        // 重置选项列表
        resetAddOptions();
        
        // 重置插入模式选择
        document.querySelectorAll("#add-field-modal .insert-mode-option").forEach(opt => {
            opt.classList.remove("selected");
            if (opt.dataset.mode === "insert") {
                opt.classList.add("selected");
                opt.querySelector("input").checked = true;
            }
        });
        
        // 设置弹窗标题
        document.getElementById("modal-title").textContent = "添加新字段";
        document.getElementById("modal-confirm").textContent = "创建字段";
        
        // 聚焦到名称输入框
        setTimeout(() => labelInput.focus(), 100);
    }
}

/**
 * 隐藏添加字段弹窗
 */
function hideAddFieldModal() {
    const modal = document.getElementById("add-field-modal");
    if (modal) {
        modal.classList.remove("show");
    }
}

/**
 * 更新 Tag 预览（基于拼音转换）
 */
function updateTagPreview() {
    const label = document.getElementById("field-label").value.trim();
    const tagPreview = document.getElementById("tag-preview");
    const tagPreviewText = document.getElementById("tag-preview-text");
    
    if (label) {
        const tag = toPinyin(label);
        tagPreviewText.textContent = tag;
        tagPreview.style.display = "block";
                    } else {
        tagPreview.style.display = "none";
    }
}

// ========== 选项管理（支持多行选项） ==========

// 临时选项存储（添加字段弹窗用）
let tempAddOptions = [];
// 临时选项存储（编辑字段弹窗用）
let tempEditOptions = [];

/**
 * 渲染选项列表
 * @param {string} containerId - 列表容器 ID
 * @param {Array} options - 选项数组
 * @param {string} mode - 'add' 或 'edit'
 */
function renderOptionsList(containerId, options, mode) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    if (options.length === 0) {
        container.innerHTML = '<div class="options-empty">暂无选项，点击下方按钮添加</div>';
        return;
    }
    
    container.innerHTML = options.map((opt, idx) => `
        <div class="option-item" data-index="${idx}">
            <span class="option-index">${idx + 1}</span>
            <div class="option-text" title="${escapeHtml(opt)}">${escapeHtml(opt)}</div>
            <button type="button" class="option-delete" onclick="removeOption(${idx}, '${mode}')">×</button>
        </div>
    `).join('');
}

/**
 * 显示添加选项的弹窗
 * @param {string} mode - 'add' 或 'edit'
 */
function showAddOptionModal(mode) {
    // 移除已存在的弹窗
    const existing = document.getElementById("option-input-modal");
    if (existing) existing.remove();
    
    const modal = document.createElement("div");
    modal.id = "option-input-modal";
    modal.className = "option-input-modal";
    modal.innerHTML = `
        <div class="option-input-box">
            <h4>添加选项</h4>
            <textarea id="new-option-text" placeholder="输入选项内容（支持多行）"></textarea>
            <div class="option-input-actions">
                <button type="button" class="btn-cancel" onclick="closeAddOptionModal()">取消</button>
                <button type="button" class="btn-confirm" onclick="confirmAddOption('${mode}')">确定</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
    
    // 自动聚焦
    setTimeout(() => {
        document.getElementById("new-option-text")?.focus();
    }, 100);
}

/**
 * 关闭添加选项弹窗
 */
function closeAddOptionModal() {
    const modal = document.getElementById("option-input-modal");
    if (modal) modal.remove();
}

/**
 * 确认添加选项
 */
function confirmAddOption(mode) {
    const textarea = document.getElementById("new-option-text");
    const value = textarea?.value.trim();
    
    if (!value) {
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification("请输入选项内容", "error");
        }
        return;
    }
    
    if (mode === 'add') {
        tempAddOptions.push(value);
        renderOptionsList("field-options-list", tempAddOptions, 'add');
    } else if (mode === 'edit') {
        tempEditOptions.push(value);
        renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
    }
    
    closeAddOptionModal();
}

/**
 * 移除选项
 */
function removeOption(index, mode) {
    if (mode === 'add') {
        tempAddOptions.splice(index, 1);
        renderOptionsList("field-options-list", tempAddOptions, 'add');
    } else if (mode === 'edit') {
        tempEditOptions.splice(index, 1);
        renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
    }
}

/**
 * 重置添加字段弹窗的选项
 */
function resetAddOptions() {
    tempAddOptions = [];
    renderOptionsList("field-options-list", tempAddOptions, 'add');
}

/**
 * 设置编辑字段弹窗的选项
 */
function setEditOptions(options) {
    tempEditOptions = [...(options || [])];
    renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
}

/**
 * 添加新字段（统一添加到 pendingFields）
 */
function addCustomFieldFromModal() {
    const label = document.getElementById("field-label").value.trim();
    const type = document.getElementById("field-type").value;
    const formatFn = document.getElementById("field-format")?.value || "none";
    const insertMode = document.querySelector('#add-field-modal input[name="insert-mode"]:checked')?.value || "insert";
    
    // 验证
    if (!label) {
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification("请输入字段名称", "error");
        }
        return;
    }
    
    // 自动生成 Tag（拼音）
    let tag = toPinyin(label);
    
    // 检查 tag 是否在 contractConfig 和 pendingFields 中重复
    let counter = 1;
    let originalTag = tag;
    const allTags = [];
    contractConfig.forEach(sec => {
        if (sec.fields) {
            sec.fields.forEach(f => {
                if (f.tag) allTags.push(f.tag);
            });
        }
    });
    pendingFields.forEach(f => {
        if (f.tag) allTags.push(f.tag);
    });
    while (allTags.includes(tag)) {
        tag = originalTag + counter;
        counter++;
    }
    
    // 从临时数组获取选项
    let options = [...tempAddOptions];
    
    // 选择类型需要至少一个选项
    if (type === "select" || type === "radio") {
        if (options.length === 0) {
            if (window.wpsAdapter && window.wpsAdapter.showNotification) {
                window.wpsAdapter.showNotification("请添加至少一个选项", "error");
            }
            return;
        }
    }
    
    // 创建字段对象
    const newField = {
        id: "pending_" + Date.now(),
        label,
        tag,
        type,
        options: options.length > 0 ? options : undefined,
        formatFn: (type !== "select" && type !== "radio" && formatFn !== "none") ? formatFn : undefined,
        hasParagraphToggle: insertMode === "paragraph" || insertMode === "both",
        isCustom: true  // 标记为自定义字段
    };
    
    // 添加到待放置区
    pendingFields.push(newField);
    
    // 保存到文档
    DocumentStateManager.save();
    
    // 重置选项
    resetAddOptions();
    
    // 重新渲染底部面板
    renderCustomFieldsPanel();
    
    hideAddFieldModal();
    if (window.wpsAdapter && window.wpsAdapter.showNotification) {
        window.wpsAdapter.showNotification(`已创建字段: ${label}`, "success");
    }
}

/**
 * 渲染自定义字段面板
 */
function renderCustomFieldsPanel() {
    const container = document.getElementById("custom-field-list");
    if (!container) return;
    
    container.innerHTML = "";
    
    // 添加字段按钮
    const addBtn = document.createElement("div");
    addBtn.className = "add-field-card";
    addBtn.id = "btn-add-field";
    addBtn.innerHTML = `<i class="ms-Icon ms-Icon--Add" aria-hidden="true"></i> <span>添加字段</span>`;
    addBtn.onclick = showAddFieldModal;
    container.appendChild(addBtn);
    
    // 渲染待放置字段
    pendingFields.forEach((field, idx) => {
        const card = document.createElement("div");
        card.className = "custom-field-card";
        card.dataset.fieldId = field.id;
        card.innerHTML = `
            <span class="field-label">${escapeHtml(field.label)}</span>
            <span class="field-meta">${field.type}</span>
            <button class="field-delete-btn" onclick="deletePendingField('${field.id}')">×</button>
        `;
                container.appendChild(card);
            });
        }

/**
 * 删除待放置字段
 */
function deletePendingField(fieldId) {
    const index = pendingFields.findIndex(f => f.id === fieldId);
    if (index >= 0) {
        pendingFields.splice(index, 1);
        DocumentStateManager.save();
        renderCustomFieldsPanel();
    }
}

/**
 * 编辑 Section 中的字段
 */
function editFieldInSection(sectionId, fieldIdOrIndex) {
    // 查找 section 和字段
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields) return;
    
    let fieldIndex = -1;
    if (typeof fieldIdOrIndex === 'number') {
        fieldIndex = fieldIdOrIndex;
    } else {
        fieldIndex = section.fields.findIndex(f => f.id === fieldIdOrIndex);
    }
    
    if (fieldIndex < 0 || fieldIndex >= section.fields.length) return;
    const field = section.fields[fieldIndex];
    
    // 显示编辑弹窗
    showFieldEditModal(sectionId, fieldIndex, field);
}

/**
 * 显示字段编辑弹窗
 */
function showFieldEditModal(sectionId, fieldIndex, field) {
    // 如果弹窗不存在，创建它
    let modal = document.getElementById("universal-field-edit-modal");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "universal-field-edit-modal";
        modal.className = "modal-overlay";
        modal.innerHTML = `
            <div class="modal-content" style="max-width: 480px;">
                <div class="modal-header">
                    <h3 id="ufm-title">编辑字段</h3>
                    <button class="modal-close" onclick="hideFieldEditModal()">×</button>
                </div>
                <div class="modal-body">
                    <input type="hidden" id="ufm-section-id">
                    <input type="hidden" id="ufm-field-index">
                    
                    <div class="modal-field-group">
                        <label>字段名称</label>
                        <input type="text" id="ufm-label" placeholder="如：签订时间">
                    </div>
                    
                    <div class="modal-field-group">
                        <label>Tag 标签 (只读)</label>
                        <input type="text" id="ufm-tag" readonly style="background:#f1f5f9;color:#64748b;">
                    </div>
                    
                    <div class="modal-field-group">
                        <label>字段类型</label>
                        <select id="ufm-type" onchange="onUfmTypeChange()">
                            <option value="text">文本</option>
                            <option value="number">数字</option>
                            <option value="date">日期</option>
                            <option value="select">下拉选择</option>
                            <option value="radio">单选按钮</option>
                        </select>
                    </div>
                    
                    <div class="modal-field-group" id="ufm-format-group">
                        <label>输出格式</label>
                        <select id="ufm-format">
                            <option value="none">无格式化（直接输出）</option>
                            <option value="dateUnderline">日期: 2024年01月15日</option>
                            <option value="dateYearMonth">日期: 2024年01月</option>
                            <option value="chineseNumber">中文大写: 壹佰（100）</option>
                            <option value="chineseNumberWan">万元: 壹佰（100）万元</option>
                            <option value="amountWithChinese">完整金额大写</option>
                            <option value="articleNumber">条款: 第五条</option>
                        </select>
                    </div>
                    
                    <div class="modal-field-group" id="ufm-options-group" style="display:none;">
                        <label>选项列表</label>
                        <div class="options-list" id="ufm-options-list"></div>
                        <button type="button" class="add-option-btn" onclick="showAddOptionModal('edit')">
                            + 添加选项
                        </button>
                    </div>
                    
                    <div class="modal-field-group">
                        <label>移动到 Section</label>
                        <select id="ufm-target-section">
                        </select>
                    </div>
                </div>
                <div class="modal-footer" style="display:flex;gap:10px;">
                    <button class="modal-btn danger" onclick="deleteFieldFromEditModal()" style="background:#ef4444;color:#fff;border:none;padding:9px 18px;border-radius:6px;cursor:pointer;">
                        删除字段
                    </button>
                    <div style="flex:1;"></div>
                    <button class="modal-cancel" onclick="hideFieldEditModal()">取消</button>
                    <button class="modal-confirm" onclick="saveFieldEdit()">保存</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }
    
    // 填充表单
    document.getElementById("ufm-section-id").value = sectionId;
    document.getElementById("ufm-field-index").value = fieldIndex;
    document.getElementById("ufm-label").value = field.label || "";
    document.getElementById("ufm-tag").value = field.tag || "";
    document.getElementById("ufm-type").value = field.type || "text";
    document.getElementById("ufm-format").value = field.formatFn || "none";
    
    // 选项 - 使用选项列表
    const optionsGroup = document.getElementById("ufm-options-group");
    const formatGroup = document.getElementById("ufm-format-group");
    if (field.type === "select" || field.type === "radio") {
        optionsGroup.style.display = "block";
        formatGroup.style.display = "none";
        setEditOptions(field.options || []);
    } else {
        optionsGroup.style.display = "none";
        formatGroup.style.display = "block";
        setEditOptions([]);
    }
    
    // 填充目标 Section 下拉框
    const targetSelect = document.getElementById("ufm-target-section");
    targetSelect.innerHTML = "";
    contractConfig.forEach(sec => {
        if (sec.fields) { // 只显示有 fields 的普通 section
            const opt = document.createElement("option");
            opt.value = sec.id;
            opt.textContent = sec.header.label;
            if (sec.id === sectionId) opt.selected = true;
            targetSelect.appendChild(opt);
        }
    });
    
    modal.classList.add("show");
}

/**
 * 编辑弹窗类型切换
 */
function onUfmTypeChange() {
    const type = document.getElementById("ufm-type").value;
    const optionsGroup = document.getElementById("ufm-options-group");
    const formatGroup = document.getElementById("ufm-format-group");
    
    if (type === "select" || type === "radio") {
        optionsGroup.style.display = "block";
        formatGroup.style.display = "none";
    } else {
        optionsGroup.style.display = "none";
        formatGroup.style.display = "block";
    }
}

/**
 * 隐藏字段编辑弹窗
 */
function hideFieldEditModal() {
    const modal = document.getElementById("universal-field-edit-modal");
    if (modal) {
        modal.classList.remove("show");
    }
}

/**
 * 保存字段编辑
 */
function saveFieldEdit() {
    const sectionId = document.getElementById("ufm-section-id").value;
    const fieldIndex = parseInt(document.getElementById("ufm-field-index").value);
    const newLabel = document.getElementById("ufm-label").value.trim();
    const newType = document.getElementById("ufm-type").value;
    const newFormat = document.getElementById("ufm-format").value;
    const targetSectionId = document.getElementById("ufm-target-section").value;
    
    if (!newLabel) {
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification("请输入字段名称", "error");
        }
        return;
    }
    
    // 查找并更新字段
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields || !section.fields[fieldIndex]) {
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification("字段不存在", "error");
        }
        return;
    }
    
    const field = section.fields[fieldIndex];
    
    // 更新字段属性
    field.label = newLabel;
    field.type = newType;
    
    // 更新格式
    if (newType !== "select" && newType !== "radio") {
        if (newFormat && newFormat !== "none") {
            field.formatFn = newFormat;
        } else {
            delete field.formatFn;
        }
    }
    
    // 更新选项 - 从临时数组读取
    if (newType === "select" || newType === "radio") {
        if (tempEditOptions.length === 0) {
            if (window.wpsAdapter && window.wpsAdapter.showNotification) {
                window.wpsAdapter.showNotification("请添加至少一个选项", "error");
            }
            return;
        }
        field.options = [...tempEditOptions];
    } else {
        // 非选项类型，清除选项
        delete field.options;
    }
    
    // 如果需要移动到其他 Section
    if (targetSectionId !== sectionId) {
        // 从原 section 移除
        section.fields.splice(fieldIndex, 1);
        
        // 添加到目标 section
        const targetSection = contractConfig.find(s => s.id === targetSectionId);
        if (targetSection && targetSection.fields) {
            targetSection.fields.push(field);
        }
    }
    
    // 保存配置
    DocumentStateManager.save();
    
    // 重新构建表单
    buildForm();
    
    hideFieldEditModal();
    if (window.wpsAdapter && window.wpsAdapter.showNotification) {
        window.wpsAdapter.showNotification(`字段 "${newLabel}" 已更新`, "success");
    }
}

/**
 * 从编辑弹窗删除字段
 */
function deleteFieldFromEditModal() {
    const sectionId = document.getElementById("ufm-section-id").value;
    const fieldIndex = parseInt(document.getElementById("ufm-field-index").value);
    
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields || !section.fields[fieldIndex]) {
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification("字段不存在", "error");
        }
        hideFieldEditModal();
        return;
    }
    
    const field = section.fields[fieldIndex];
    const fieldLabel = field.label;
    
    if (confirm(`确定要删除字段"${fieldLabel}"吗？此操作不可撤销。`)) {
        // 删除字段
        section.fields.splice(fieldIndex, 1);
        
        // 保存配置
        DocumentStateManager.save();
        
        // 重新构建表单
    buildForm();
    
        hideFieldEditModal();
        if (window.wpsAdapter && window.wpsAdapter.showNotification) {
            window.wpsAdapter.showNotification(`字段 "${fieldLabel}" 已删除`, "success");
        }
    }
}

/**
 * 删除 Section 中的字段（直接调用版本）
 */
function deleteFieldInSection(sectionId, fieldIndex) {
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields) return;
    
    if (fieldIndex >= 0 && fieldIndex < section.fields.length) {
        const field = section.fields[fieldIndex];
        if (confirm(`确定要删除字段"${field.label}"吗？`)) {
            section.fields.splice(fieldIndex, 1);
            DocumentStateManager.save();
            buildForm();
        }
    }
}

/**
 * 初始化自定义字段管理器
 */
function initCustomFieldsManager() {
    // FAB 按钮切换抽屉
    const fab = document.getElementById("custom-field-fab");
    const drawer = document.getElementById("custom-field-drawer");
    const drawerClose = document.getElementById("drawer-close");
    
    if (fab && drawer) {
        fab.addEventListener("click", () => {
            drawer.classList.toggle("open");
            fab.classList.toggle("active");
        });
    }
    
    if (drawerClose && drawer && fab) {
        drawerClose.addEventListener("click", () => {
            drawer.classList.remove("open");
            fab.classList.remove("active");
        });
    }
    
    // 弹窗关闭按钮
    const modalClose = document.getElementById("modal-close");
    const modalCancel = document.getElementById("modal-cancel");
    if (modalClose) {
        modalClose.addEventListener("click", hideAddFieldModal);
    }
    if (modalCancel) {
        modalCancel.addEventListener("click", hideAddFieldModal);
    }
    
    // 弹窗确认按钮
    const modalConfirm = document.getElementById("modal-confirm");
    if (modalConfirm) {
        modalConfirm.addEventListener("click", addCustomFieldFromModal);
    }
    
    // 字段类型切换时显示/隐藏选项组和格式组
    const fieldType = document.getElementById("field-type");
    const optionsGroup = document.getElementById("options-group");
    const formatGroup = document.getElementById("format-group");
    if (fieldType) {
        fieldType.addEventListener("change", () => {
            const type = fieldType.value;
            if (type === "select" || type === "radio") {
                if (optionsGroup) optionsGroup.style.display = "block";
                if (formatGroup) formatGroup.style.display = "none";
            } else {
                if (optionsGroup) optionsGroup.style.display = "none";
                if (formatGroup) formatGroup.style.display = "block";
            }
        });
    }
    
    // 字段名称输入时更新 Tag 预览
    const fieldLabel = document.getElementById("field-label");
    if (fieldLabel) {
        fieldLabel.addEventListener("input", updateTagPreview);
    }
    
    // 添加选项按钮
    const addOptionBtn = document.getElementById("add-option-btn");
    if (addOptionBtn) {
        addOptionBtn.addEventListener("click", () => showAddOptionModal('add'));
    }
    
    // 插入模式选择
    document.querySelectorAll(".insert-mode-option").forEach(opt => {
        opt.addEventListener("click", function() {
            document.querySelectorAll(".insert-mode-option").forEach(o => o.classList.remove("selected"));
            this.classList.add("selected");
            this.querySelector("input").checked = true;
        });
    });
    
    // 渲染自定义字段面板
    renderCustomFieldsPanel();
}

// 导出全局函数
window.loadPEVCTemplate = loadPEVCTemplate;
window.syncFormToSelectedDocuments = syncFormToSelectedDocuments;
window.autoGenerateForm = autoGenerateForm;
window.batchEmbedAIFields = batchEmbedAIFields;
window.onAutoEmbedToggleChange = onAutoEmbedToggleChange;
window.refreshAIServiceStatusManual = refreshAIServiceStatusManual;
window.showAddFieldModal = showAddFieldModal;
window.closeAddOptionModal = closeAddOptionModal;
window.confirmAddOption = confirmAddOption;
window.removeOption = removeOption;
window.deletePendingField = deletePendingField;
window.showFieldEditModal = showFieldEditModal;
window.hideFieldEditModal = hideFieldEditModal;
window.saveFieldEdit = saveFieldEdit;
window.deleteFieldFromEditModal = deleteFieldFromEditModal;
window.onUfmTypeChange = onUfmTypeChange;

document.addEventListener("DOMContentLoaded", init);
