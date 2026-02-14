/**
 * WPS 适配层 - 将 Word API 调用转换为 WPS API 调用
 * 核心功能：ContentControl、全状态存储
 */

const ADDIN_STATE_KEY = "ContractAddin_State_V2";
const WPS_DATA_KEY = "PEVC_FormData";
const WPS_HIDDEN_KEY = "PEVC_HiddenTags";
const ORIGINAL_TEXT_KEY = "PEVC_OriginalTexts";
const isWpsEnvironment = typeof wps !== 'undefined';
let hiddenParagraphs = new Set();

// 队列统计信息
const queueStats = {
    taskCount: 0,           // 总任务数
    completedCount: 0,      // 已完成任务数
    failedCount: 0,        // 失败任务数
    totalWaitMs: 0,        // 总等待时间（毫秒）
    totalTaskMs: 0,        // 总任务执行时间（毫秒）
    recentErrors: [],       // 最近错误（保留最后10个）
    lastTaskEndTime: 0,     // 上次任务结束时间
    maxRecentErrors: 10     // 保留最近错误数量
};

const wpsActionQueue = (() => {
    let queue = Promise.resolve();
    let currentTaskStartTime = 0;

    return {
        add(task, options = {}) {
            const cooldownMs = typeof options.cooldownMs === "number" ? options.cooldownMs : 2000;
            const taskStartTime = Date.now();
            queueStats.taskCount++;

            const run = queue.then(async () => {
                // 记录任务开始时间
                currentTaskStartTime = Date.now();
                // 计算上一个任务结束后到现在的等待时间
                if (queueStats.lastTaskEndTime > 0) {
                    queueStats.totalWaitMs += currentTaskStartTime - queueStats.lastTaskEndTime;
                }

                try {
                    const result = await task();
                    // 记录任务执行时间
                    const taskMs = Date.now() - currentTaskStartTime;
                    queueStats.totalTaskMs += taskMs;
                    queueStats.completedCount++;
                    queueStats.lastTaskEndTime = Date.now();
                    return result;
                }
                catch (err) {
                    console.error("[WPS Queue] Error:", err);
                    // 记录错误
                    queueStats.recentErrors.push({
                        time: new Date().toISOString(),
                        message: err.message || String(err)
                    });
                    if (queueStats.recentErrors.length > queueStats.maxRecentErrors) {
                        queueStats.recentErrors.shift();
                    }
                    queueStats.failedCount++;
                    throw err;
                }
                finally {
                    if (cooldownMs > 0) await new Promise(r => setTimeout(r, cooldownMs));
                }
            });
            // 保持队列链可继续，同时把错误抛给当前调用方处理。
            queue = run.catch(() => undefined);
            return run;
        },

        /**
         * 获取队列统计信息
         * 返回：{ length, avgWaitMs, avgTaskMs, completedCount, failedCount, recentErrors }
         */
        getStats() {
            const avgWaitMs = queueStats.taskCount > 1
                ? Math.round(queueStats.totalWaitMs / (queueStats.taskCount - 1))
                : 0;
            const avgTaskMs = queueStats.completedCount > 0
                ? Math.round(queueStats.totalTaskMs / queueStats.completedCount)
                : 0;

            return {
                length: 1, // 当前有1个pending的promise（但无法精确获取）
                totalTasks: queueStats.taskCount,
                completedCount: queueStats.completedCount,
                failedCount: queueStats.failedCount,
                avgWaitMs: avgWaitMs,
                avgTaskMs: avgTaskMs,
                totalWaitMs: queueStats.totalWaitMs,
                totalTaskMs: queueStats.totalTaskMs,
                recentErrors: queueStats.recentErrors.slice()
            };
        },

        /**
         * 获取队列健康状态
         * 返回：{ status: 'healthy' | 'warning' | 'critical', reasons: [] }
         */
        getHealth() {
            const stats = this.getStats();
            const reasons = [];
            let status = 'healthy';

            // 检查错误率
            if (stats.totalTasks > 0) {
                const errorRate = stats.failedCount / stats.totalTasks;
                if (errorRate > 0.1) {
                    reasons.push('错误率过高: ' + (errorRate * 100).toFixed(1) + '%');
                    status = 'critical';
                } else if (errorRate > 0.05) {
                    reasons.push('错误率略高: ' + (errorRate * 100).toFixed(1) + '%');
                    status = 'warning';
                }
            }

            // 检查平均等待时间
            if (stats.avgWaitMs > 30000) {
                reasons.push('平均等待时间过长: ' + Math.round(stats.avgWaitMs / 1000) + 's');
                status = 'warning';
            }

            // 检查最近是否有错误
            if (stats.recentErrors.length > 0) {
                const lastError = stats.recentErrors[stats.recentErrors.length - 1];
                const lastErrorTime = new Date(lastError.time).getTime();
                const now = Date.now();
                if (now - lastErrorTime < 60000) { // 1分钟内
                    reasons.push('最近1分钟有错误: ' + lastError.message);
                    status = 'warning';
                }
            }

            return { status, reasons, stats };
        },

        /**
         * 重置统计信息
         */
        resetStats() {
            queueStats.taskCount = 0;
            queueStats.completedCount = 0;
            queueStats.failedCount = 0;
            queueStats.totalWaitMs = 0;
            queueStats.totalTaskMs = 0;
            queueStats.recentErrors = [];
            queueStats.lastTaskEndTime = 0;
        }
    };
})();

// ==================== 核心全状态存储 (WPS 专用) ====================

async function wpsSaveAddinState(state) {
    if (!isWpsEnvironment) return;
    return wpsActionQueue.add(async () => {
        const app = wps.WpsApplication();
        const doc = app.ActiveDocument;
        if (!doc) return;
        const jsonString = JSON.stringify(state);
        const props = doc.CustomDocumentProperties;
        let found = false;
        for (let i = 1; i <= props.Count; i++) {
            if (props.Item(i).Name === ADDIN_STATE_KEY) {
                props.Item(i).Value = jsonString;
                found = true;
                break;
            }
        }
        if (!found) props.Add(ADDIN_STATE_KEY, false, 4, jsonString);
        console.log("[WPS Storage] ✅ 全状态已存入文档属性");
    });
}

async function wpsLoadAddinState() {
    if (!isWpsEnvironment) return null;
    return wpsActionQueue.add(async () => {
        const app = wps.WpsApplication();
        const doc = app.ActiveDocument;
        if (!doc) return null;
        const props = doc.CustomDocumentProperties;
        for (let i = 1; i <= props.Count; i++) {
            if (props.Item(i).Name === ADDIN_STATE_KEY) {
                return JSON.parse(props.Item(i).Value);
            }
        }
        return null;
    });
}

async function wpsClearAddinState() {
    if (!isWpsEnvironment) return;
    return wpsActionQueue.add(async () => {
        const props = wps.WpsApplication().ActiveDocument.CustomDocumentProperties;
        for (let i = 1; i <= props.Count; i++) {
            if (props.Item(i).Name === ADDIN_STATE_KEY) {
                props.Item(i).Delete();
                break;
            }
        }
    });
}

/**
 * 获取当前活动文档的唯一标识
 * FullName 为空时回退到 Name
 */
async function wpsGetActiveDocumentKey() {
    if (!isWpsEnvironment) return "";
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return "";
            return doc.FullName || doc.Name || "";
        } catch (err) {
            console.error("[WPS] Get active document key error:", err);
            return "";
        }
    }, { cooldownMs: 0 });
}

/**
 * 从文档 Content Control 读取已填写的值
 * @returns {object} { tag: value }
 */
async function wpsReadFormDataFromDocumentCCs() {
    if (!isWpsEnvironment) return {};
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return {};

            const result = {};
            const contentControls = doc.ContentControls;
            const total = contentControls.Count;

            for (let i = 1; i <= total; i++) {
                const cc = contentControls.Item(i);
                const tag = cc.Tag;
                if (!tag) continue;
                const rawText = (cc.Range.Text || "").replace(/\u200B/g, "").trim();
                if (!rawText) continue;
                const displayLabel = cc.Title || tag;
                const placeholderFull = `【${displayLabel}】`;
                const placeholderLabelAlt = `[${displayLabel}]`;
                const placeholderTagAlt = `[${tag}]`;
                const placeholderTagFull = `【${tag}】`;
                if (rawText === placeholderFull || rawText === placeholderLabelAlt || rawText === placeholderTagAlt || rawText === placeholderTagFull) {
                    continue;
                }
                let valueText = rawText;
                if (/^【.*】$/.test(rawText)) {
                    valueText = rawText.slice(1, -1);
                }
                result[tag] = valueText;
            }

            return result;
        } catch (err) {
            console.error("[WPS] Read CC form data error:", err);
            return {};
        }
    });
}

/**
 * 弹出文件选择器，让用户多选目标文档
 */
async function wpsPickTargetDocuments() {
    if (!isWpsEnvironment) return [];
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            // 3 = msoFileDialogFilePicker
            const fd = app.FileDialog(3); 
            fd.Filters.Clear();
            fd.Filters.Add("Word 文档", "*.doc;*.docx");
            fd.AllowMultiSelect = true;
            fd.Title = "请选择要同步的目标文档";
            
            if (fd.Show() === -1) { // -1 = 按下“打开”按钮
                const items = [];
                for (let i = 1; i <= fd.SelectedItems.Count; i++) {
                    items.push(fd.SelectedItems.Item(i));
                }
                return items;
            }
            return [];
        } catch (err) {
            console.error("[WPS] FileDialog error:", err);
            return [];
        }
    });
}

/**
 * 跨文档同步数据
 * @param {object} formData - 表单数据
 * @param {object} labelMap - 标签映射
 * @param {string[]} targets - 目标文件路径数组
 */
async function wpsSyncFormDataToDocuments(formData, labelMap, targets) {
    if (!isWpsEnvironment) return { successCount: 0, missingCount: 0 };
    
    return wpsActionQueue.add(async () => {
        const app = wps.WpsApplication();
        let successCount = 0;
        let totalMissing = 0;
        
        // 记录原文档路径，以便最后切回
        const originalDoc = app.ActiveDocument;
        const originalPath = originalDoc ? originalDoc.FullName : "";

        for (const path of targets) {
            try {
                // 如果是当前正在编辑的文档，直接跳过（同步逻辑应在 taskpane.js 实时处理）
                if (path === originalPath) continue;

                // 打开文档 (路径, ConfirmConversions, ReadOnly)
                const doc = app.Documents.Open(path, false, false);
                if (!doc) continue;

                let docMissing = 0;
                const ccs = doc.ContentControls;
                
                // 遍历数据进行同步
                for (const [tag, value] of Object.entries(formData)) {
                    const displayLabel = labelMap[tag] || tag;
                    const rawValue = (value === null || value === undefined) ? "" : String(value);
                    const trimmedValue = rawValue.trim();
                    let coreValue = trimmedValue;
                    if (rawValue === "") {
                        coreValue = "";
                    } else if (/^【.*】$/.test(trimmedValue)) {
                        coreValue = trimmedValue.slice(1, -1);
                    } else if (/^\[.*\]$/.test(trimmedValue)) {
                        coreValue = trimmedValue.slice(1, -1);
                    }
                    const textToInsert = coreValue === "" ? `【${displayLabel}】` : `【${coreValue}】`;
                    
                    let foundInDoc = false;
                    for (let i = 1; i <= ccs.Count; i++) {
                        const cc = ccs.Item(i);
                        if (cc.Tag === tag) {
                            // 保留格式更新文本
                            const range = cc.Range;
                            const font = range.Font;
                            const savedFont = {
                                Name: font.Name,
                                Size: font.Size,
                                Bold: font.Bold,
                                Italic: font.Italic,
                                Color: font.Color
                            };
                            
                            range.Text = textToInsert;
                            
                            if (savedFont.Name) font.Name = savedFont.Name;
                            if (savedFont.Size) font.Size = savedFont.Size;
                            font.Bold = savedFont.Bold;
                            font.Italic = savedFont.Italic;
                            font.Color = savedFont.Color;
                            
                            foundInDoc = true;
                        }
                    }
                    
                    if (!foundInDoc) {
                        // 尝试替换文本占位符
                        const patterns = [`【${tag}】`, `【${displayLabel}】`, `[${tag}]`, `[${displayLabel}]`];
                        for (const pattern of patterns) {
                            const find = doc.Content.Find;
                            find.ClearFormatting();
                            find.Text = pattern;
                            find.Replacement.ClearFormatting();
                            find.Replacement.Text = textToInsert;
                            const replaced = find.Execute(pattern, false, false, false, false, false, true, 1, false, textToInsert, 2);
                            if (replaced) {
                                foundInDoc = true;
                                break;
                            }
                        }
                    }
                    
                    if (!foundInDoc) docMissing++;
                }
                
                doc.Save();
                doc.Close();
                successCount++;
                totalMissing += docMissing;
                
                // 强制等待 2 秒冷却，保证 WPS 稳定性
                await new Promise(r => setTimeout(r, 2000));
                
            } catch (docErr) {
                console.error(`[Sync] Error processing ${path}:`, docErr);
            }
        }
        
        // 最后切回原文档（如果还开着）
        try {
            if (originalPath) app.Documents.Item(originalPath).Activate();
        } catch (e) {}

        return { successCount, missingCount: totalMissing };
    });
}

/**
 * 保存 CC 的原始文本和格式快照到文档属性
 * 格式: { "tag1": { text: "原始文本1", font: {...} }, "tag2": {...} }
 * @param {string} tag - CC 的 Tag
 * @param {string} originalText - 原始文本内容
 * @param {object|null} fontSnapshot - 字体快照 { Name, Size, Bold, Italic, Color, Underline }
 */
async function saveOriginalText(tag, originalText, fontSnapshot = null) {
    if (!isWpsEnvironment) return;
    
    console.log(`[WPS Storage] 开始保存: ${tag}, 文本长度: ${originalText?.length || 0}, 格式快照: ${fontSnapshot ? 'Yes' : 'No'}`);

    try {
        const app = wps.WpsApplication();
        const doc = app.ActiveDocument;
        if (!doc) return;
        
        const customProps = doc.CustomDocumentProperties;
        let originalTexts = {};
        
        // 遍历读取现有属性
        for (let i = 1; i <= customProps.Count; i++) {
            try {
                const prop = customProps.Item(i);
                if (prop.Name === ORIGINAL_TEXT_KEY) {
                    originalTexts = JSON.parse(prop.Value || "{}");
                    console.log(`[WPS Storage] 读取到现有数据，共 ${Object.keys(originalTexts).length} 个tag`);
                    break;
                }
            } catch (e) {
                console.warn(`[WPS Storage] 读取属性 ${i} 失败:`, e.message);
            }
        }
        
        // 保存新的数据结构：{ text, font }
        originalTexts[tag] = {
            text: originalText,
            font: fontSnapshot
        };
        
        const jsonString = JSON.stringify(originalTexts);
        console.log(`[WPS Storage] JSON 大小: ${jsonString.length} 字节`);
        
        // 查找并更新，或新建属性
        let found = false;
        for (let i = 1; i <= customProps.Count; i++) {
            try {
                const prop = customProps.Item(i);
                if (prop.Name === ORIGINAL_TEXT_KEY) {
                    prop.Value = jsonString;
                    found = true;
                    console.log(`[WPS Storage] 更新已有属性`);
                    break;
                }
            } catch (e) {
                console.warn(`[WPS Storage] 更新属性 ${i} 失败:`, e.message);
            }
        }
        
        if (!found) {
            customProps.Add(ORIGINAL_TEXT_KEY, false, 4 /* msoPropertyTypeString */, jsonString);
            console.log(`[WPS Storage] 创建新属性`);
        }
        
        // 验证保存结果
        let verified = false;
        for (let i = 1; i <= customProps.Count; i++) {
            try {
                const prop = customProps.Item(i);
                if (prop.Name === ORIGINAL_TEXT_KEY) {
                    const saved = JSON.parse(prop.Value || "{}");
                    verified = saved[tag] && saved[tag].text === originalText;
                    break;
                }
            } catch (e) {}
        }
        
        console.log(`[WPS Storage] ✅ 保存完成: ${tag}, 验证: ${verified ? '通过' : '失败'}`);
    } catch (err) {
        console.error("[WPS Storage] 保存原始文本失败:", err);
    }
}

/**
 * 获取 CC 的原始文本和格式快照
 * @param {string} tag - CC 的 Tag
 * @returns {object|null} { text, font } 或 null（如果不存在）
 */
async function getOriginalText(tag) {
    if (!isWpsEnvironment) return null;
    
    try {
        const app = wps.WpsApplication();
        const doc = app.ActiveDocument;
        if (!doc) return null;
        
        const customProps = doc.CustomDocumentProperties;
        
        // 遍历查找属性
        for (let i = 1; i <= customProps.Count; i++) {
            try {
                const prop = customProps.Item(i);
                if (prop.Name === ORIGINAL_TEXT_KEY) {
                    const originalTexts = JSON.parse(prop.Value || "{}");
                    const snapshot = originalTexts[tag];
                    
                    if (snapshot) {
                        // 兼容旧格式（纯字符串）和新格式（{text, font}）
                        if (typeof snapshot === 'string') {
                            console.log(`[WPS Storage] 读取到旧格式数据: ${tag}`);
                            return { text: snapshot, font: null };
                        } else {
                            console.log(`[WPS Storage] 读取成功: ${tag}, 文本长度: ${snapshot.text?.length || 0}, 格式: ${snapshot.font ? 'Yes' : 'No'}`);
                            return snapshot;
                        }
                    }
                }
            } catch (e) {
                console.warn(`[WPS Storage] 读取属性 ${i} 失败:`, e.message);
            }
        }
        
        // 未找到时，打印所有属性名称用于诊断
        console.warn(`[WPS Storage] ⚠️ 未找到 tag: ${tag}`);
        console.log(`[WPS Storage] 文档中现有的自定义属性（共 ${customProps.Count} 个）:`);
        for (let i = 1; i <= customProps.Count; i++) {
            try {
                const prop = customProps.Item(i);
                console.log(`  - [${i}] ${prop.Name}`);
            } catch (e) {}
        }
        
        return null;
    } catch (err) {
        console.error("[WPS Storage] 获取原始文本失败:", err);
        return null;
    }
}

// ==================== 原有埋点操作适配 ====================

async function wpsUpdateContent(tag, value, label = null) {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Update CC: ${tag} = ${value}`);
        return;
    }
    const enqueueAt = Date.now();
    const displayLabel = label || tag;
    const valueStr = (value === null || value === undefined) ? "" : String(value);
    const valueTrim = valueStr.trim();
    const valueLen = valueStr.length;
    const valueIsEmpty = valueStr === "";
    const valueIsBracketed = /^【.*】$/.test(valueTrim);
    let valueCore = valueTrim;
    if (valueIsEmpty) {
        valueCore = "";
    } else if (/^【.*】$/.test(valueTrim)) {
        valueCore = valueTrim.slice(1, -1);
    } else if (/^\[.*\]$/.test(valueTrim)) {
        valueCore = valueTrim.slice(1, -1);
    }
    const textToInsert = valueCore === ""
        ? `【${displayLabel}】`
        : valueCore;
    
    return wpsActionQueue.add(async () => {
        const taskStartAt = Date.now();
        const queueDelayMs = taskStartAt - enqueueAt;
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) {
                console.warn("[WPS] No active document");
                return;
            }
            
            // 方法1: 尝试通过 Tag 查找所有 ContentControl（批量更新）
            const contentControls = doc.ContentControls;
            let updatedCount = 0;
            
            for (let i = 1; i <= contentControls.Count; i++) {
                const cc = contentControls.Item(i);
                if (cc.Tag === tag) {
                    const range = cc.Range;
                    
                    // 1. 保存字体格式
                    let savedFont = null;
                    try {
                        savedFont = {
                            name: range.Font.Name,
                            size: range.Font.Size,
                            bold: range.Font.Bold,
                            italic: range.Font.Italic,
                            color: range.Font.Color
                        };
                    } catch (e) {}
                    
                    // 2. 更新文本
                    range.Text = textToInsert;
                    
                    // 3. 恢复字体格式
                    if (savedFont) {
                        try {
                            if (savedFont.name) range.Font.Name = savedFont.name;
                            if (savedFont.size) range.Font.Size = savedFont.size;
                            if (savedFont.bold !== undefined) range.Font.Bold = savedFont.bold;
                            if (savedFont.italic !== undefined) range.Font.Italic = savedFont.italic;
                            if (savedFont.color) range.Font.Color = savedFont.color;
                        } catch (e) {}
                    }
                    
                    updatedCount++;
                }
            }
            
            let found = updatedCount > 0;
            if (found) {
                console.log(`[WPS] Updated ${updatedCount} CC(s) for tag: ${tag}`);
            }
            
            // 方法2: 如果没找到 CC，尝试查找替换文本占位符
            if (!found) {
                const patterns = [`【${tag}】`, `【${displayLabel}】`, `[${tag}]`, `[${displayLabel}]`];
                for (const pattern of patterns) {
                    const find = doc.Content.Find;
                    find.ClearFormatting();
                    find.Text = pattern;
                    find.Replacement.ClearFormatting();
                    find.Replacement.Text = textToInsert;
                    const replaced = find.Execute(
                        pattern, false, false, false, false, false, true, 1, false, 
                        textToInsert, 2 // wdReplaceAll
                    );
                    if (replaced) {
                        console.log(`[WPS] Replaced text: ${pattern} -> ${textToInsert}`);
                        found = true;
                        break;
                    }
                }
            }
            
            if (!found && updatedCount === 0) {
                console.log(`[WPS] Tag not found: ${tag}`);
            }
        } catch (err) {
            console.error(`[WPS] Update content error for ${tag}:`, err);
        }
    }, { cooldownMs: 100 });
}

async function wpsInsertControl(tag, title, isWrapper = false) {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Insert CC: ${tag}, wrapper=${isWrapper}`);
        return;
    }
    const enqueueAt = Date.now();
    return wpsActionQueue.add(async () => {
        const taskStartAt = Date.now();
        const queueDelayMs = taskStartAt - enqueueAt;
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            const selection = app.Selection;
            
            if (!doc || !selection) {
                showNotification("请先打开一个文档", "error");
                return;
            }
            
            // 统计已存在的同 tag CC 数量（用于提示）
            let existingCount = 0;
            const existingCCs = doc.ContentControls;
            for (let i = 1; i <= existingCCs.Count; i++) {
                if (existingCCs.Item(i).Tag === tag) {
                    existingCount++;
                }
            }
            
            let cc;
            let originalText = "";
            let fontSnapshot = null;
            const placeholder = `【${title}】`;
            
            // 检查是否有文本被选中
            const isTextSelected = selection.Start !== selection.End;
            const originalTextLen = isTextSelected ? (selection.Text ? selection.Text.length : 0) : 0;
            
            if (isWrapper) {
                // 包裹模式：必须有选中内容，包裹并保存
                if (isTextSelected) {
                    originalText = selection.Text;
                    
                    // **关键：在覆盖前捕获格式快照**
                    try {
                        const range = selection.Range;
                        const font = range.Font;
                        fontSnapshot = {
                            Name: font.Name,
                            Size: font.Size,
                            Bold: font.Bold,
                            Italic: font.Italic,
                            Color: font.Color,
                            Underline: font.Underline
                        };
                        console.log(`[WPS] 包裹段落: ${tag}, 原始文本长度: ${originalText.length}, 格式: ${fontSnapshot.Name}/${fontSnapshot.Size}`);
                    } catch (e) {
                        console.warn(`[WPS] 捕获格式失败:`, e.message);
                    }
                    
                    cc = selection.Range.ContentControls.Add(1); // wdContentControlRichText
                    cc.Range.Text = placeholder; // 统一显示为占位符
                } else {
                    showNotification("请先选中要包裹的内容", "warning");
                    return;
                }
            } else {
                // 插入模式：
                if (isTextSelected) {
                    // 有选中文本 → 保存后替换为占位符并包裹
                    originalText = selection.Text;
                    
                    // **关键：在覆盖前捕获格式快照**
                    try {
                        const range = selection.Range;
                        const font = range.Font;
                        fontSnapshot = {
                            Name: font.Name,
                            Size: font.Size,
                            Bold: font.Bold,
                            Italic: font.Italic,
                            Color: font.Color,
                            Underline: font.Underline
                        };
                        console.log(`[WPS] 选中文字插入: ${tag}, 原始文本: "${originalText.substring(0, 20)}...", 格式: ${fontSnapshot.Name}/${fontSnapshot.Size}`);
                    } catch (e) {
                        console.warn(`[WPS] 捕获格式失败:`, e.message);
                    }
                    
                    cc = selection.Range.ContentControls.Add(1);
                    cc.Range.Text = placeholder;
                } else {
                    // 无选中 → 直接插入占位符并包裹
                    originalText = "";
                    fontSnapshot = null; // 无原文，无格式
                    selection.TypeText(placeholder);
                    selection.MoveLeft(1, placeholder.length, true);
                    cc = selection.Range.ContentControls.Add(1);
                    console.log(`[WPS] 直接插入占位符: ${tag}`);
                }
            }
            
            if (cc) {
                cc.Tag = tag;
                cc.Title = title;
                
                // **关键：保存原始文本和格式快照到文档属性**
                await saveOriginalText(tag, originalText, fontSnapshot);
                
                const countInfo = existingCount > 0 ? ` (共${existingCount + 1}处)` : "";
                console.log(`[WPS] 已完成埋点: ${tag}${countInfo}, 初始显示: ${placeholder}`);
                showNotification(`已插入 [${title}]${countInfo}`, "success");
            }
        } catch (err) {
            console.error(`[WPS] Insert control error:`, err);
            showNotification(`插入失败: ${err.message}`, "error");
        }
    }, { cooldownMs: 100 });
}

async function wpsToggleVisibility(tag, isVisible, label = null) {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Toggle visibility: ${tag} = ${isVisible}, label=${label}`);
        return;
    }
    
    const displayLabel = label || tag;
    const placeholderText = `[${displayLabel}]`;
    // 用特殊标记区分占位符，便于识别和删除
    const placeholderMarker = `\u200B`; // 零宽空格作为标记
    const fullPlaceholder = placeholderMarker + placeholderText + placeholderMarker;
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return;
            
            // 通过 Tag 查找 ContentControl
            const contentControls = doc.ContentControls;
            let targetCC = null;
            
            for (let i = 1; i <= contentControls.Count; i++) {
                const cc = contentControls.Item(i);
                if (cc.Tag === tag) {
                    targetCC = cc;
                    break;
                }
            }
            
            if (!targetCC) {
                console.log(`[WPS] CC not found for toggle: ${tag}`);
                showNotification(`未找到 [${displayLabel}]，请先插入`, "warning");
                return;
            }
            
            const ccRange = targetCC.Range;
            const currentText = ccRange.Text || "";
            
            if (isVisible) {
                // ========== 显示（恢复）==========
                // 检查是否包含占位符标记
                if (currentText.includes(placeholderMarker)) {
                    // 找到占位符的位置
                    const markerStart = currentText.indexOf(placeholderMarker);
                    const markerEnd = currentText.lastIndexOf(placeholderMarker);
                    
                    if (markerStart !== -1 && markerEnd > markerStart) {
                        // 计算占位符长度
                        const placeholderLen = markerEnd - markerStart + 1;
                        
                        // 获取占位符范围并删除
                        const startPos = ccRange.Start + markerStart;
                        const endPos = ccRange.Start + markerEnd + 1;
                        const placeholderRange = doc.Range(startPos, endPos);
                        placeholderRange.Delete();
                        
                        // 取消原内容的隐藏
                        targetCC.Range.Font.Hidden = false;
                        
                        console.log(`[WPS] Show paragraph: ${tag} (placeholder removed)`);
                    }
                } else {
                    // 没有占位符标记，直接取消隐藏
                    ccRange.Font.Hidden = false;
                    console.log(`[WPS] Show paragraph: ${tag} (unhide)`);
                }
                
                // 从隐藏列表移除
                hiddenParagraphs.delete(tag);
                showNotification(`已显示: ${displayLabel}`, "success");
                
            } else {
                // ========== 隐藏 ==========
                // 检查是否已经隐藏（已有占位符）
                if (currentText.includes(placeholderMarker)) {
                    console.log(`[WPS] Already hidden: ${tag}`);
                    showNotification(`${displayLabel} 已经是隐藏状态`, "info");
                    return;
                }
                
                // 1. 在开头插入占位符
                const insertRange = doc.Range(ccRange.Start, ccRange.Start);
                insertRange.Text = fullPlaceholder;
                
                // 2. 将占位符后面的原内容设为隐藏
                const originalStart = ccRange.Start + fullPlaceholder.length;
                const originalEnd = ccRange.End;
                if (originalEnd > originalStart) {
                    const originalRange = doc.Range(originalStart, originalEnd);
                    originalRange.Font.Hidden = true;
                }
                
                // 添加到隐藏列表
                hiddenParagraphs.add(tag);
                console.log(`[WPS] Hide paragraph: ${tag} (content hidden, placeholder visible)`);
                showNotification(`已隐藏: ${displayLabel}`, "info");
            }
            
        } catch (err) {
            console.error(`[WPS] Toggle visibility error:`, err);
            showNotification(`操作失败: ${err.message}`, "error");
        }
    });
}

/**
 * 保存表单数据到文档属性
 * @param {object} formData - 表单数据对象
 */
async function wpsSaveFormDataToDocument(formData) {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Save form data to document`);
        return;
    }
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return;
            
            const jsonString = JSON.stringify(formData);
            const customProps = doc.CustomDocumentProperties;
            
            // 查找是否已存在
            let found = false;
            for (let i = 1; i <= customProps.Count; i++) {
                if (customProps.Item(i).Name === WPS_DATA_KEY) {
                    customProps.Item(i).Value = jsonString;
                    found = true;
                    break;
                }
            }
            
            if (!found) {
                customProps.Add(WPS_DATA_KEY, false, 4 /* msoPropertyTypeString */, jsonString);
            }
            
            console.log(`[WPS] Form data saved to document properties`);
            showNotification("数据已保存到文档", "success");
        } catch (err) {
            console.error(`[WPS] Save form data error:`, err);
            showNotification(`保存失败: ${err.message}`, "error");
        }
    });
}

/**
 * 从文档属性加载表单数据
 * @returns {object|null} 表单数据对象
 */
async function wpsLoadFormDataFromDocument() {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Load form data from document`);
        return null;
    }
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return null;
            
            const customProps = doc.CustomDocumentProperties;
            for (let i = 1; i <= customProps.Count; i++) {
                if (customProps.Item(i).Name === WPS_DATA_KEY) {
                    const jsonString = customProps.Item(i).Value;
                    console.log(`[WPS] Form data loaded from document properties`);
                    return JSON.parse(jsonString);
                }
            }
            
            console.log(`[WPS] No form data found in document`);
            return null;
        } catch (err) {
            console.error(`[WPS] Load form data error:`, err);
            return null;
        }
    });
}

/**
 * 批量应用表单数据到文档
 * @param {object} formData - 表单数据
 * @param {object} fieldLabelMap - tag -> label 的映射
 */
async function wpsApplyFormDataToDocument(formData, fieldLabelMap = {}) {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Apply form data to document`);
        return;
    }
    
    for (const [tag, value] of Object.entries(formData)) {
        const label = fieldLabelMap[tag] || tag;
        await wpsUpdateContent(tag, value, label);
    }
    
    showNotification(`已应用 ${Object.keys(formData).length} 个字段`, "success");
}

/**
 * 尝试自动备份当前文档
 * @returns {object} { success, fileName, error }
 */
async function wpsBackupDocument() {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Backup document`);
        return { success: true, fileName: "mock_backup.docx" };
    }
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) {
                return { success: false, error: "没有打开的文档" };
            }
            
            // 获取当前文档路径和名称
            const fullPath = doc.FullName;
            const docName = doc.Name;
            
            if (!fullPath || fullPath === "") {
                return { success: false, error: "文档尚未保存" };
            }
            
            // 生成备份文件名
            const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
            const timeStr = new Date().toTimeString().slice(0, 5).replace(':', '');
            const baseName = docName.replace(/\.docx?$/i, '');
            const backupName = `${baseName}_备份_${timestamp}_${timeStr}.docx`;
            
            // 获取文档所在目录
            const lastSlash = Math.max(fullPath.lastIndexOf('/'), fullPath.lastIndexOf('\\'));
            const dirPath = fullPath.substring(0, lastSlash + 1);
            const backupPath = dirPath + backupName;
            
            // 先保存当前文档
            doc.Save();
            
            // 另存为备份
            doc.SaveAs2(backupPath);
            
            console.log(`[WPS] Backup created: ${backupPath}`);
            return { success: true, fileName: backupName, path: backupPath };
            
        } catch (err) {
            console.error(`[WPS] Backup error:`, err);
            return { success: false, error: err.message };
        }
    });
}

/**
 * 执行清理操作（删除隐藏段落，包括占位符和隐藏的原内容）
 * @returns {object} { deletedCount }
 */
async function wpsExecuteCleanup() {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Execute cleanup`);
        return { deletedCount: hiddenParagraphs.size };
    }
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) {
                return { deletedCount: 0 };
            }
            
            let deletedCount = 0;
            const contentControls = doc.ContentControls;
            
            // 收集要删除的 CC 的 tag（不用索引，因为删除会改变索引）
            const tagsToDelete = [...hiddenParagraphs];
            
            // 遍历删除
            for (const tag of tagsToDelete) {
                // 每次重新查找，因为索引会变
                for (let i = contentControls.Count; i >= 1; i--) {
                    try {
                        const cc = contentControls.Item(i);
                        if (cc.Tag === tag) {
                            cc.Delete(true); // true = 删除内容
                            deletedCount++;
                            break; // 找到并删除后跳出内层循环
                        }
                    } catch (e) {
                        console.warn(`[WPS] Failed to delete CC ${tag}:`, e);
                    }
                }
            }
            
            // 清空隐藏列表
            hiddenParagraphs.clear();
            
            console.log(`[WPS] Cleanup completed: deleted=${deletedCount}`);
            return { deletedCount };
            
        } catch (err) {
            console.error(`[WPS] Cleanup error:`, err);
            return { deletedCount: 0 };
        }
    });
}

/**
 * 撤销所有埋点（删除所有 Content Control，恢复原始内容并保留格式）
 * @returns {object} { deletedCount, restoredCount, errors }
 */
async function wpsUndoAllEmbeds() {
    if (!isWpsEnvironment) {
        console.log(`[Mock] Undo all embeds`);
        return { deletedCount: 0, restoredCount: 0, errors: [] };
    }
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) {
                showNotification("没有打开的文档", "error");
                return { deletedCount: 0, restoredCount: 0, errors: [] };
            }
            
            console.log(`[WPS Undo] ========== 开始撤销所有埋点 ==========`);
            
            let deletedCount = 0;
            let restoredCount = 0;
            const errors = [];
            const contentControls = doc.ContentControls;
            const totalCount = contentControls.Count;
            
            console.log(`[WPS Undo] 找到 ${totalCount} 个 Content Control`);
            
            // 反向遍历删除（避免索引变化问题）
            for (let i = totalCount; i >= 1; i--) {
                try {
                    const cc = contentControls.Item(i);
                    const tag = cc.Tag || `(no-tag-${i})`;
                    
                    console.log(`[WPS Undo] [${totalCount - i + 1}/${totalCount}] 处理 CC: ${tag}`);
                    
                    // 1. 从文档属性获取原始文本和格式快照
                    const snapshot = await getOriginalText(tag);
                    
                    let restoreText = null;
                    if (snapshot !== null && snapshot.text !== undefined) {
                        restoreText = String(snapshot.text);
                    }

                    // 2. 采用标准的"还原内容 -> 还原格式 -> 剥离控件"流程
                    try {
                        const range = cc.Range;
                        
                        if (restoreText !== null && restoreText !== "") {
                            // 有原始文本：还原内容和格式
                            range.Text = restoreText;
                            
                            if (snapshot.font) {
                                try {
                                    const font = range.Font;
                                    if (snapshot.font.Name) font.Name = snapshot.font.Name;
                                    if (snapshot.font.NameFarEast) font.NameFarEast = snapshot.font.NameFarEast;
                                    if (snapshot.font.Size) font.Size = snapshot.font.Size;
                                    if (snapshot.font.Bold !== undefined) font.Bold = snapshot.font.Bold;
                                    if (snapshot.font.Italic !== undefined) font.Italic = snapshot.font.Italic;
                                    if (snapshot.font.Color !== undefined) font.Color = snapshot.font.Color;
                                    if (snapshot.font.Underline !== undefined) font.Underline = snapshot.font.Underline;
                                    
                                    console.log(`[WPS Undo]   ✅ 已还原原文内容与格式: ${tag}`);
                                } catch (formatErr) {
                                    console.warn(`[WPS Undo]   ⚠️ 格式还原异常: ${formatErr.message}`);
                                }
                            }
                            
                            // 剥离控件框架，保留内容 (Delete(false))
                            cc.Delete(false); 
                            restoredCount++;
                            console.log(`[WPS Undo]   ✅ 已剥离 CC 框架并保留内容: ${tag}`);
                        } else {
                            // 无原始文本：直接彻底删除控件及其内容 (Delete(true))
                            cc.Delete(true);
                            console.log(`[WPS Undo]   ✅ 无原始文本，已彻底删除 CC: ${tag}`);
                        }
                    } catch (undoErr) {
                        console.error(`[WPS Undo]   ❌ 撤销操作失败: ${undoErr.message}`);
                        errors.push(`撤销 ${tag} 失败: ${undoErr.message}`);
                    }
                    
                    deletedCount++;
                    
                    if (deletedCount % 10 === 0 || deletedCount === 1) {
                        console.log(`[WPS Undo] 进度: ${deletedCount}/${totalCount}`);
                    }
                } catch (e) {
                    const errMsg = `处理 CC[${i}] 失败: ${e.message}`;
                    errors.push(errMsg);
                    console.error(`[WPS Undo] ❌ ${errMsg}`);
                }
            }
            
            // 清空隐藏段落列表（因为 CC 都删除了）
            hiddenParagraphs.clear();
            
            console.log(`[WPS Undo] ========== 处理完成 ==========`);
            console.log(`[WPS Undo] 已成功撤销 ${deletedCount} 个埋点`);
            
            if (deletedCount > 0) {
                const msg = restoredCount > 0 
                    ? `已撤销 ${deletedCount} 个埋点，恢复 ${restoredCount} 个原始值`
                    : `已撤销 ${deletedCount} 个埋点`;
                showNotification(msg, "success");
            }
            
            return { deletedCount, restoredCount, errors };
            
        } catch (err) {
            console.error(`[WPS Undo] 错误:`, err);
            showNotification(`撤销失败: ${err.message}`, "error");
            return { deletedCount: 0, restoredCount: 0, errors: [err.message] };
        }
    });
}

/**
 * 获取文档中未填写的字段（内容为 [xxx] 格式）
 */
async function wpsGetUnfilledInDocument() {
    if (!isWpsEnvironment) {
        return [];
    }
    
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return [];
            
            const unfilled = [];
            const contentControls = doc.ContentControls;
            
            for (let i = 1; i <= contentControls.Count; i++) {
                const cc = contentControls.Item(i);
                const text = (cc.Range.Text || "").replace(/\u200B/g, "").trim();
                if (!text) continue;
                const displayLabel = cc.Title || cc.Tag;
                const placeholderFull = `【${displayLabel}】`;
                const placeholderLabelAlt = `[${displayLabel}]`;
                const placeholderTagAlt = `[${cc.Tag}]`;
                const placeholderTagFull = `【${cc.Tag}】`;
                const isPlaceholder = text === placeholderFull 
                    || text === placeholderLabelAlt 
                    || text === placeholderTagAlt 
                    || text === placeholderTagFull;
                
                // 检查是否未填写（内容为占位符，但不是隐藏的）
                if (isPlaceholder && !hiddenParagraphs.has(cc.Tag)) {
                    unfilled.push({ tag: cc.Tag, title: displayLabel });
                }
            }
            
            return unfilled;
        } catch (err) {
            console.error(`[WPS] Get unfilled error:`, err);
            return [];
        }
    });
}

/**
 * 获取当前文档纯文本
 */
async function wpsGetDocumentText() {
    if (!isWpsEnvironment) return "";
    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;
            if (!doc) return "";
            return doc.Content ? (doc.Content.Text || "") : "";
        } catch (err) {
            console.error("[WPS] Get document text error:", err);
            return "";
        }
    });
}

/**
 * 获取隐藏段落列表
 */
function getHiddenParagraphs() {
    return [...hiddenParagraphs];
}

// =====================================================================
// 页面内通知 (替代 alert)
// =====================================================================
function showNotification(message, type = "info") {
    const existingNotif = document.getElementById("app-notification");
    if (existingNotif) existingNotif.remove();
    
    const notif = document.createElement("div");
    notif.id = "app-notification";
    notif.style.cssText = `
        position: fixed;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        padding: 10px 16px;
        border-radius: 8px;
        font-size: 12px;
        font-weight: 500;
        z-index: 9999;
        max-width: 80%;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        animation: slideDown 0.3s ease;
    `;
    
    if (type === "error") {
        notif.style.background = "#fde8e8";
        notif.style.color = "#c53030";
        notif.style.border = "1px solid #fc8181";
    } else if (type === "warning") {
        notif.style.background = "#fefcbf";
        notif.style.color = "#744210";
        notif.style.border = "1px solid #f6e05e";
    } else if (type === "success") {
        notif.style.background = "#c6f6d5";
        notif.style.color = "#22543d";
        notif.style.border = "1px solid #68d391";
    } else {
        notif.style.background = "#bee3f8";
        notif.style.color = "#2a4365";
        notif.style.border = "1px solid #63b3ed";
    }
    
    notif.textContent = message;
    document.body.appendChild(notif);
    
    setTimeout(() => {
        if (notif.parentNode) {
            notif.style.opacity = "0";
            notif.style.transition = "opacity 0.3s";
            setTimeout(() => notif.remove(), 300);
        }
    }, 4000);
}

/**
 * 根据搜索文本在文档中查找并插入 Content Control（用于批量埋点）
 * @param {string} tag - 埋点标签
 * @param {string} title - 显示标题
 * @param {string|string[]|object} searchInput - 要搜索的文本（支持候选列表/上下文对象）
 * @returns {Promise<{success: boolean, message: string}>}
 */
async function wpsInsertControlBySearch(tag, title, searchInput) {
    const normalizeText = (value) => (typeof value === "string" ? value : "").replace(/\s+/g, "").trim();
    const isLikelyPlaceholderToken = (value) => {
        const token = normalizeText(value);
        if (!token) return false;
        return /^[【】\[\]（）()_＿*＊·.…—-]+$/.test(token) || token === "N/A";
    };

    const aiContext = (searchInput && typeof searchInput === "object" && !Array.isArray(searchInput))
        ? {
            context: typeof searchInput.context === "string" ? searchInput.context : (searchInput.aiContext && searchInput.aiContext.context) || "",
            prefix: typeof searchInput.prefix === "string" ? searchInput.prefix : (searchInput.aiContext && searchInput.aiContext.prefix) || "",
            placeholder: typeof searchInput.placeholder === "string" ? searchInput.placeholder : (searchInput.aiContext && searchInput.aiContext.placeholder) || "",
            suffix: typeof searchInput.suffix === "string" ? searchInput.suffix : (searchInput.aiContext && searchInput.aiContext.suffix) || "",
            confidence: (searchInput.aiContext && searchInput.aiContext.confidence) || searchInput.confidence || "medium"
        }
        : { context: "", prefix: "", placeholder: "", suffix: "", confidence: "medium" };

    const rawCandidates = (searchInput && typeof searchInput === "object" && !Array.isArray(searchInput))
        ? (Array.isArray(searchInput.searchCandidates) ? searchInput.searchCandidates : (Array.isArray(searchInput.candidates) ? searchInput.candidates : []))
        : (Array.isArray(searchInput) ? searchInput : [searchInput]);

    const composedAnchor = `${aiContext.prefix || ""}${aiContext.placeholder || ""}${aiContext.suffix || ""}`.trim();

    const candidateMetaMap = new Map();
    const pushCandidateMeta = (text, kind, baseScore) => {
        const candidate = typeof text === "string" ? text.trim() : "";
        if (!candidate) return;
        const key = normalizeText(candidate);
        if (!key) return;
        const previous = candidateMetaMap.get(key);
        if (!previous || baseScore > previous.baseScore) {
            candidateMetaMap.set(key, { candidate, kind, baseScore });
        }
    };

    const placeholderLikelySymbol = isLikelyPlaceholderToken(aiContext.placeholder);

    pushCandidateMeta(composedAnchor, "composed", 130);
    pushCandidateMeta(aiContext.context, "context", 100);
    pushCandidateMeta(aiContext.placeholder, "placeholder", 70);
    if (placeholderLikelySymbol && aiContext.prefix) pushCandidateMeta(aiContext.prefix, "prefix_anchor", 92);
    if (placeholderLikelySymbol && aiContext.suffix) pushCandidateMeta(aiContext.suffix, "suffix_anchor", 88);

    rawCandidates.forEach((item) => {
        const value = typeof item === "string" ? item : "";
        const normalized = normalizeText(value);
        if (!normalized) return;
        const kind = normalized === normalizeText(composedAnchor)
            ? "composed"
            : (normalized === normalizeText(aiContext.context)
                ? "context"
                : (normalized === normalizeText(aiContext.placeholder) ? "placeholder" : "candidate"));
        const baseScore = kind === "composed" ? 130 : (kind === "context" ? 100 : (kind === "placeholder" ? 70 : 80));
        pushCandidateMeta(value, kind, baseScore);
    });

    const candidates = Array.from(candidateMetaMap.values())
        .sort((a, b) => b.baseScore - a.baseScore || b.candidate.length - a.candidate.length);

    const prefixNorm = normalizeText(aiContext.prefix);
    const suffixNorm = normalizeText(aiContext.suffix);
    const hasAnchor = Boolean(prefixNorm || suffixNorm);

    if (!isWpsEnvironment) {
        console.log(`[Mock] Insert CC by search: ${tag}, candidates=${JSON.stringify(candidates.map((v) => v.candidate))}`);
        return { success: true, message: "Mock mode" };
    }

    if (candidates.length === 0) {
        return { success: false, message: "搜索文本为空" };
    }

    const getRangeText = (range) => {
        try {
            return typeof range.Text === "string" ? range.Text : "";
        } catch (_) {
            return "";
        }
    };

    const getRangeBound = (range, key, fallback = 0) => {
        try {
            const value = range[key];
            return typeof value === "number" ? value : fallback;
        } catch (_) {
            return fallback;
        }
    };

    const getNeighborText = (doc, start, end, beforeLength, afterLength) => {
        const content = doc.Content;
        const docStart = getRangeBound(content, "Start", 0);
        const docEnd = getRangeBound(content, "End", 0);
        const safeBefore = Math.max(80, beforeLength * 4);
        const safeAfter = Math.max(80, afterLength * 4);

        let before = "";
        let after = "";

        try {
            const from = Math.max(docStart, start - safeBefore);
            const to = Math.max(docStart, Math.min(start, docEnd));
            if (doc.Range && to >= from) before = getRangeText(doc.Range(from, to));
        } catch (_) {}

        try {
            const from = Math.max(docStart, Math.min(end, docEnd));
            const to = Math.min(docEnd, end + safeAfter);
            if (doc.Range && to >= from) after = getRangeText(doc.Range(from, to));
        } catch (_) {}

        return { before, after };
    };

    const hasConflictControl = (doc, start, end, selfTag) => {
        try {
            const controls = doc.ContentControls;
            for (let i = 1; i <= controls.Count; i++) {
                const cc = controls.Item(i);
                const ccStart = getRangeBound(cc.Range, "Start", -1);
                const ccEnd = getRangeBound(cc.Range, "End", -1);
                const overlap = ccStart < end && ccEnd > start;
                if (overlap && cc.Tag !== selfTag) return true;
            }
        } catch (err) {
            console.warn("[WPS] 检查埋点冲突失败:", err && err.message ? err.message : err);
        }
        return false;
    };

    const collectMatches = (doc, candidate, maxMatches = 16) => {
        const matches = [];
        const content = doc.Content;
        const docStart = getRangeBound(content, "Start", 0);
        const docEnd = getRangeBound(content, "End", 0);
        let cursor = docStart;
        let loopGuard = 0;

        while (cursor <= docEnd && matches.length < maxMatches && loopGuard < maxMatches * 4) {
            loopGuard += 1;
            let searchRange = null;

            try {
                if (doc.Range) {
                    searchRange = doc.Range(cursor, docEnd);
                }
            } catch (_) {
                searchRange = null;
            }

            if (!searchRange) {
                try {
                    searchRange = content.Duplicate ? content.Duplicate : content;
                    if (searchRange.SetRange) searchRange.SetRange(cursor, docEnd);
                } catch (_) {
                    searchRange = content;
                }
            }

            const find = searchRange.Find;
            if (!find) break;

            find.ClearFormatting();
            find.Text = candidate;
            find.Forward = true;
            find.Wrap = 0;
            find.MatchCase = false;
            find.MatchWholeWord = false;

            if (!find.Execute()) break;

            const matchedRange = searchRange.Duplicate ? searchRange.Duplicate : searchRange;
            const start = getRangeBound(matchedRange, "Start", cursor);
            const end = getRangeBound(matchedRange, "End", start + candidate.length);
            if (end <= start) break;

            matches.push({
                range: matchedRange,
                start,
                end,
                text: getRangeText(matchedRange)
            });

            cursor = Math.max(end, cursor + 1);
        }

        return matches;
    };

    const evaluateMatch = (doc, matched, meta) => {
        const around = getNeighborText(doc, matched.start, matched.end, prefixNorm.length, suffixNorm.length);
        const beforeNorm = normalizeText(around.before);
        const afterNorm = normalizeText(around.after);

        const isPrefixAnchor = meta.kind === "prefix_anchor";
        const isSuffixAnchor = meta.kind === "suffix_anchor";

        let prefixOk = !prefixNorm || beforeNorm.endsWith(prefixNorm);
        let suffixOk = !suffixNorm || afterNorm.startsWith(suffixNorm);

        if (isPrefixAnchor) {
            prefixOk = Boolean(prefixNorm);
            if (suffixNorm) suffixOk = afterNorm.includes(suffixNorm);
        }
        if (isSuffixAnchor) {
            suffixOk = Boolean(suffixNorm);
            if (prefixNorm) prefixOk = beforeNorm.includes(prefixNorm);
        }

        const conflict = hasConflictControl(doc, matched.start, matched.end, tag);

        let score = meta.baseScore;
        if (prefixNorm) score += prefixOk ? 45 : -30;
        if (suffixNorm) score += suffixOk ? 45 : -30;
        if (prefixNorm && suffixNorm && prefixOk && suffixOk) score += 20;
        if (aiContext.confidence === "high") score += 8;
        if (aiContext.confidence === "low") score -= 6;
        if (meta.kind === "candidate") score -= 5;
        if (isPrefixAnchor || isSuffixAnchor) score += 6;

        const contextNorm = normalizeText(aiContext.context);
        if (contextNorm) {
            const aroundNorm = normalizeText(`${around.before || ""}${matched.text || meta.candidate || ""}${around.after || ""}`);
            const contextHead = contextNorm.slice(0, Math.min(16, contextNorm.length));
            const contextTail = contextNorm.length > 20 ? contextNorm.slice(-16) : "";
            if ((contextHead && aroundNorm.includes(contextHead)) || (contextTail && aroundNorm.includes(contextTail))) {
                score += 12;
            }
        }

        if (conflict) score -= 120;

        return {
            ...matched,
            candidate: meta.candidate,
            kind: meta.kind,
            score,
            prefixOk,
            suffixOk,
            conflict
        };
    };

    return wpsActionQueue.add(async () => {
        try {
            const app = wps.WpsApplication();
            const doc = app.ActiveDocument;

            if (!doc) {
                return { success: false, message: "请先打开一个文档" };
            }

            const existingCCs = doc.ContentControls;
            for (let i = 1; i <= existingCCs.Count; i++) {
                if (existingCCs.Item(i).Tag === tag) {
                    console.log(`[WPS] Tag already exists: ${tag}`);
                    return {
                        success: false,
                        message: `已存在埋点: ${title}`,
                        detail: {
                            status: "already_exists",
                            strategy: "existing",
                            score: 999
                        }
                    };
                }
            }

            const allMatches = [];
            const dedupeByPos = new Map();

            for (const meta of candidates) {
                const matches = collectMatches(doc, meta.candidate);
                for (const item of matches) {
                    const key = `${item.start}:${item.end}`;
                    const scored = evaluateMatch(doc, item, meta);
                    const previous = dedupeByPos.get(key);
                    if (!previous || scored.score > previous.score) {
                        dedupeByPos.set(key, scored);
                    }
                }
            }

            dedupeByPos.forEach((value) => allMatches.push(value));
            allMatches.sort((a, b) => b.score - a.score || (b.end - b.start) - (a.end - a.start));

            if (!allMatches.length) {
                const preview = candidates[0].candidate.substring(0, 20);
                console.log(`[WPS] Text not found for ${tag}, tried ${candidates.length} candidates`);
                return {
                    success: false,
                    message: `未找到: ${preview}...`,
                    detail: {
                        status: "not_found",
                        strategy: "none",
                        score: 0,
                        candidatesTried: candidates.map((v) => v.candidate)
                    }
                };
            }

            const best = allMatches[0];
            const second = allMatches[1];
            const scoreGap = second ? (best.score - second.score) : 999;
            const minScore = hasAnchor ? 85 : 65;
            const relaxedPass = best.score >= 70 && scoreGap >= 22 && (best.prefixOk || best.suffixOk) && !best.conflict;

            if (best.conflict) {
                return {
                    success: false,
                    message: "命中位置已被其他埋点占用",
                    detail: {
                        status: "conflict",
                        strategy: best.kind,
                        score: best.score,
                        scoreGap,
                        candidate: best.candidate,
                        prefixOk: best.prefixOk,
                        suffixOk: best.suffixOk,
                        snippet: (best.text || "").substring(0, 50)
                    }
                };
            }

            if (best.score < minScore && !relaxedPass) {
                return {
                    success: false,
                    message: `定位置信度不足(${best.score})`,
                    detail: {
                        status: "low_confidence",
                        strategy: best.kind,
                        score: best.score,
                        scoreGap,
                        candidate: best.candidate,
                        prefixOk: best.prefixOk,
                        suffixOk: best.suffixOk,
                        snippet: (best.text || "").substring(0, 50)
                    }
                };
            }

            const ambiguousGap = (best.kind === "prefix_anchor" || best.kind === "suffix_anchor") ? 20 : 12;
            if (scoreGap < ambiguousGap && (best.kind === "placeholder" || best.kind === "prefix_anchor" || best.kind === "suffix_anchor") && hasAnchor && (!best.prefixOk || !best.suffixOk)) {
                return {
                    success: false,
                    message: "命中位置存在歧义，请手动埋点",
                    detail: {
                        status: "ambiguous",
                        strategy: best.kind,
                        score: best.score,
                        scoreGap,
                        candidate: best.candidate,
                        prefixOk: best.prefixOk,
                        suffixOk: best.suffixOk,
                        snippet: (best.text || "").substring(0, 50)
                    }
                };
            }

            let range = best.range;
            if ((best.kind === "prefix_anchor" || best.kind === "suffix_anchor") && doc.Range) {
                const collapsePos = best.kind === "prefix_anchor" ? best.end : best.start;
                try {
                    range = doc.Range(collapsePos, collapsePos);
                } catch (_) {
                    range = best.range;
                }
            }
            if (!range) {
                return { success: false, message: "获取匹配位置失败" };
            }

            let fontSnapshot = null;
            try {
                const font = range.Font;
                fontSnapshot = {
                    Name: font.Name,
                    Size: font.Size,
                    Bold: font.Bold,
                    Italic: font.Italic,
                    Color: font.Color,
                    Underline: font.Underline
                };
            } catch (e) {
                console.warn(`[WPS] 捕获格式失败:`, e.message);
            }

            const originalText = (best.kind === "prefix_anchor" || best.kind === "suffix_anchor") ? "" : (range.Text || best.candidate);

            const cc = range.ContentControls.Add(1);
            cc.Tag = tag;
            cc.Title = title;

            if (fontSnapshot) {
                try {
                    const ccFont = cc.Range.Font;
                    if (fontSnapshot.Name) ccFont.Name = fontSnapshot.Name;
                    if (fontSnapshot.Size) ccFont.Size = fontSnapshot.Size;
                    if (fontSnapshot.Bold !== undefined) ccFont.Bold = fontSnapshot.Bold;
                    if (fontSnapshot.Italic !== undefined) ccFont.Italic = fontSnapshot.Italic;
                    if (fontSnapshot.Underline !== undefined) ccFont.Underline = fontSnapshot.Underline;
                } catch (e) {
                    console.warn(`[WPS] 恢复格式失败:`, e.message);
                }
            }

            await saveOriginalText(tag, originalText, fontSnapshot);

            const usedRelaxed = best.score < minScore;
            console.log(`[WPS] Inserted CC by search: ${tag}, strategy=${best.kind}, score=${best.score}, matched="${(best.text || best.candidate).substring(0, 20)}...", relaxed=${usedRelaxed}`);
            return {
                success: true,
                message: usedRelaxed ? `已埋点(宽松置信): ${title}` : `已埋点: ${title}`,
                detail: {
                    status: usedRelaxed ? "success_relaxed" : "success",
                    strategy: best.kind,
                    score: best.score,
                    scoreGap,
                    candidate: best.candidate,
                    prefixOk: best.prefixOk,
                    suffixOk: best.suffixOk,
                    relaxed: usedRelaxed,
                    snippet: (best.text || "").substring(0, 50),
                    candidatesTried: candidates.map((v) => v.candidate),
                    totalMatches: allMatches.length
                }
            };

        } catch (err) {
            console.error(`[WPS] Insert CC by search error:`, err);
            return { success: false, message: err.message };
        }
    }, { cooldownMs: 200 });
}

window.wpsAdapter = {
    isWpsEnvironment,
    saveAddinState: wpsSaveAddinState,
    loadAddinState: wpsLoadAddinState,
    clearAddinState: wpsClearAddinState,
    getActiveDocumentKey: wpsGetActiveDocumentKey,
    readFormDataFromDocumentCCs: wpsReadFormDataFromDocumentCCs,
    updateContent: wpsUpdateContent,
    insertControl: wpsInsertControl,
    insertControlBySearch: wpsInsertControlBySearch,
    toggleVisibility: wpsToggleVisibility,
    saveFormData: wpsSaveFormDataToDocument,
    loadFormData: wpsLoadFormDataFromDocument,
    applyFormData: wpsApplyFormDataToDocument,
    backupDocument: wpsBackupDocument,
    executeCleanup: wpsExecuteCleanup,
    undoAllEmbeds: wpsUndoAllEmbeds,
    saveOriginalText: saveOriginalText,
    getOriginalText: getOriginalText,
    getUnfilledInDocument: wpsGetUnfilledInDocument,
    getHiddenParagraphs: getHiddenParagraphs,
    getDocumentText: wpsGetDocumentText,
    pickTargetDocuments: wpsPickTargetDocuments,
    syncFormDataToDocuments: wpsSyncFormDataToDocuments,
    // 队列统计接口（任务E新增）
    getQueueStats: function() { return wpsActionQueue.getStats(); },
    getQueueHealth: function() { return wpsActionQueue.getHealth(); },
    resetQueueStats: function() { return wpsActionQueue.resetStats(); }
};

window.showNotification = showNotification;

console.log(`[WPS Adapter] Initialized, WPS environment: ${isWpsEnvironment}`);
